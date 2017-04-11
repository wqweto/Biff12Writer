Attribute VB_Name = "mdGlobals"
'=========================================================================
'
' Biff12Writer (c) 2017 by wqweto@gmail.com
'
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' API
'=========================================================================

Private Const VT_I8                         As Long = &H14
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_vLastError                As Variant
Private m_uPeekArray                As UcsSafeArraySingleDimension
Private m_aPeekBuffer()             As Integer

Private Type UcsSafeArraySingleDimension
    cDims       As Integer      '--- usually 1
    fFeatures   As Integer      '--- leave 0
    cbElements  As Long         '--- bytes per element (2-int, 4-long)
    cLocks      As Long         '--- leave 0
    pvData      As Long         '--- ptr to data
    cElements   As Long         '--- UBound + 1
    lLbound     As Long         '--- LBound
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function PushError(Optional vLocalErr As Variant) As Variant
    vLocalErr = Array(Err.Number, Err.Source, Err.Description, Erl)
    m_vLastError = vLocalErr
    PushError = vLocalErr
End Function

Public Function PopRaiseError(sFunction As String, sModule As String, Optional vLocalErr As Variant)
    If Not IsMissing(vLocalErr) Then
        m_vLastError = vLocalErr
    End If
    Err.Raise m_vLastError(0), sModule & "." & sFunction & vbCrLf & m_vLastError(1), m_vLastError(2)
End Function

Public Function PopPrintError(sFunction As String, sModule As String, Optional vLocalErr As Variant)
    If Not IsMissing(vLocalErr) Then
        m_vLastError = vLocalErr
    End If
    Debug.Print "Error " & m_vLastError(0), sModule & "." & sFunction & vbCrLf & m_vLastError(1), m_vLastError(2)
End Function

Private Sub pvClose(nFile As Integer)
    On Error GoTo EH
    If nFile <> 0 Then
        Close nFile
    End If
EH:
    nFile = 0
End Sub

Public Function ReadBinaryFile(sFile As String) As Byte()
    Const FUNC_NAME     As String = "ReadBinaryFile"
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    baBuffer = ApiEmptyByteArray()
    nFile = FreeFile
    Open sFile For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    pvClose nFile
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Function

Public Sub WriteBinaryFile(sFile As String, baBuffer() As Byte)
    Const FUNC_NAME     As String = "WriteBinaryFile"
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    If InStrRev(sFile, "\") > 1 Then
        MkPath Left$(sFile, InStrRev(sFile, "\") - 1)
    End If
    DeleteFile sFile
    nFile = FreeFile
    Open sFile For Binary Access Write As nFile
    If Peek(ArrPtr(baBuffer)) <> 0 Then
        If UBound(baBuffer) >= 0 Then
            Put nFile, , baBuffer
        End If
    End If
    pvClose nFile
    Exit Sub
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Sub

Public Sub WriteTextFile(sFile As String, sText As String)
    Const FUNC_NAME     As String = "WriteTextFile"
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    If InStrRev(sFile, "\") > 1 Then
        MkPath Left$(sFile, InStrRev(sFile, "\") - 1)
    End If
    nFile = FreeFile
    Open sFile For Output As nFile
    Print #nFile, sText
    pvClose nFile
    Exit Sub
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Sub

Public Function FileAttr(sFile As String) As VbFileAttribute
    FileAttr = GetFileAttributes(sFile)
    If FileAttr = -1 Then
        FileAttr = &H8000
    End If
End Function

Public Function MkPath(sPath As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "MkPath"
    Dim vErr            As Variant
    
    On Error GoTo EH
    MkPath = (FileAttr(sPath) And vbDirectory) <> 0
    If Not MkPath Then
        If ApiCreateDirectory(sPath, 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError)
        End If
        MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        If Not MkPath And InStrRev(sPath, "\") <> 0 Then
            MkPath Left$(sPath, InStrRev(sPath, "\") - 1)
            Call ApiCreateDirectory(sPath, 0)
            MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        End If
    End If
    Exit Function
EH:
    PushError vErr
    PopRaiseError FUNC_NAME & "(sPath=" & sPath & ")", MODULE_NAME, vErr
End Function

Public Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim ret             As Long
   
    GetSystemMessage = Space$(2000)
    ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If ret > 2 Then
        If Mid$(GetSystemMessage, ret - 1, 2) = vbCrLf Then
            ret = ret - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, ret)
End Function


Public Function DeleteFile(sFileName As String) As Boolean
    Call ApiDeleteFile(sFileName)
End Function

Public Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long, Optional ByVal AddrPadding As Long = -1) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim cResult         As Collection
    Dim sPrefix         As String
    
    Set cResult = New Collection
    If lSize > 1 Then
        lIdx = Int(Log(lSize - 1) / Log(16) + 1)
    Else
        lIdx = 1
    End If
    If AddrPadding < 0 And AddrPadding < lIdx Then
        AddrPadding = lIdx
    End If
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(UnsignedAdd(lPtr, lIdx), 1) = 0 Then
                Call CopyMemory(lValue, ByVal UnsignedAdd(lPtr, lIdx), 1)
                sHex = sHex & Right$("00" & Hex$(lValue), 2) & " "
                If AddrPadding > 0 Then
                    If lValue >= 32 Then
                        sChar = sChar & Chr$(lValue)
                    Else
                        sChar = sChar & "."
                    End If
                End If
            Else
                sHex = sHex & "?? "
                If AddrPadding > 0 Then
                    sChar = sChar & "."
                End If
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            If AddrPadding > 0 Then
                sPrefix = Right$(String(AddrPadding, "0") & Hex$(lIdx - 15), AddrPadding) & ": "
            End If
            cResult.Add RTrim$(sPrefix & sHex & " " & sChar)
            sHex = vbNullString
            sChar = vbNullString
        End If
    Next
    DesignDumpMemory = ConcatCollection(cResult, vbCrLf)
End Function

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

Public Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long
    UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Public Function ToArray(oCol As Collection) As Variant
    Const FUNC_NAME     As String = "ToArray"
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    If oCol.Count > 0 Then
        ReDim vRetVal(0 To oCol.Count - 1) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = oCol(lIdx + 1)
        Next
        ToArray = vRetVal
    Else
        ToArray = Array()
    End If
    Exit Function
EH:
    PushError
    PopRaiseError FUNC_NAME, MODULE_NAME
End Function

Public Function CLngLng(vValue As Variant) As Variant
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function

Public Function ToLngLng(ByVal lLoDWord As Long, ByVal lHiDWord As Long) As Variant
    Call VariantChangeType(ToLngLng, ToLngLng, 0, VT_I8)
    Call CopyMemory(ByVal VarPtr(ToLngLng) + 8, lLoDWord, 4)
    Call CopyMemory(ByVal VarPtr(ToLngLng) + 12, lHiDWord, 4)
End Function

Public Function GetLoDWord(llValue As Variant) As Long
    Call CopyMemory(GetLoDWord, ByVal VarPtr(llValue) + 8, 4)
End Function

Public Function GetHiDWord(llValue As Variant) As Long
    Call CopyMemory(GetHiDWord, ByVal VarPtr(llValue) + 12, 4)
End Function

Public Function RollingHash(ByVal lPtr As Long, ByVal lSize As Long) As Long
    Dim lIdx            As Long
    
    If lPtr = 0 Then
        Call CopyMemory(ByVal ArrPtr(m_aPeekBuffer), 0&, 4)
        m_uPeekArray.cDims = 0
        Exit Function
    ElseIf m_uPeekArray.cDims = 0 Then
        With m_uPeekArray
            .cDims = 1
            .cbElements = 2
        End With
        Call CopyMemory(ByVal ArrPtr(m_aPeekBuffer), VarPtr(m_uPeekArray), 4)
    End If
    m_uPeekArray.pvData = lPtr
    m_uPeekArray.cElements = lSize
    For lIdx = 0 To lSize - 1
        RollingHash = (RollingHash * 263 + m_aPeekBuffer(lIdx)) And &H3FFFFF
    Next
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    '--- note: VariantCopy ne raboti za VT_BYREF wyw vDest
'    Call VariantCopy(vDest, vSrc)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function ToUtf8(sText As String) As String
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ToUtf8 = String(lSize, 0)
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal ToUtf8, lSize, 0, 0)
    End If
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = ApiEmptyByteArray
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8(sText As String) As String
    Dim lSize           As Long
    
    FromUtf8 = String$(4 * Len(sText), 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, ByVal sText, Len(sText), StrPtr(FromUtf8), Len(FromUtf8))
    FromUtf8 = Left$(FromUtf8, lSize)
End Function


Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    FromUtf8Array = String$(2 * UBound(baText), 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
    FromUtf8Array = Left$(FromUtf8Array, lSize)
End Function
