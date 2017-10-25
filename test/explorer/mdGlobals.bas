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

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long

'=========================================================================
' Functions
'=========================================================================

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
    PopRaiseError FUNC_NAME, MODULE_NAME, PushError
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

Public Function FormatXmlIndent(vDomOrString As Variant, sResult As String) As Boolean
    Dim oWriter         As Object ' MSXML2.MXXMLWriter

    On Error GoTo QH
    Set oWriter = CreateObject("MSXML2.MXXMLWriter")
    oWriter.omitXMLDeclaration = True
    oWriter.Indent = True
    With CreateObject("MSXML2.SAXXMLReader")
        Set .contentHandler = oWriter
        '--- keep CDATA elements
        .putProperty "http://xml.org/sax/properties/lexical-handler", oWriter
        .parse vDomOrString
    End With
    sResult = oWriter.Output
    '--- success
    FormatXmlIndent = True
    Exit Function
QH:
End Function
