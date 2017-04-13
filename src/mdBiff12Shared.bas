Attribute VB_Name = "mdBiff12Shared"
'=========================================================================
'
' Biff12Writer (c) 2017 by wqweto@gmail.com
'
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "mdBiff12Shared"

#Const ImplUseShared = BIFF12_USESHARED

#If ImplUseShared Then

'=========================================================================
' Public types
'=========================================================================

Public Enum UcsHorAlignmentEnum
    ucsHalLeft = 0
    ucsHalRight = 1
    ucsHalCenter = 2
End Enum

Public Type UcsBiff12BrtColorType
    m_xColorType        As Byte
    m_index             As Byte
    m_nTintAndShade     As Integer
    m_bRed              As Byte
    m_bGreen            As Byte
    m_bBlue             As Byte
    m_bAlpha            As Byte
End Type

Public Type UcsBiff12BrtFontType
    m_dyHeight          As Integer
    m_grbit             As Integer
    m_bls               As Integer
    m_sss               As Integer
    m_uls               As Byte
    m_bFamily           As Byte
    m_bCharSet          As Byte
    '--- padding 1 bytes
    m_brtColor          As UcsBiff12BrtColorType
    m_bFontScheme       As Byte
    '--- padding 3 bytes
    m_name              As String
End Type

Public Type UcsGradientStopType
    brtColor            As UcsBiff12BrtColorType
    xnumPosition        As Double
End Type

Public Type UcsBiff12BrtFillType
    m_fls               As Long
    m_brtColorFore      As UcsBiff12BrtColorType
    m_brtColorBack      As UcsBiff12BrtColorType
    m_iGradientType     As Long
    m_xnumDegree        As Double
    m_xnumFillToLeft    As Double
    m_xnumFillToRight   As Double
    m_xnumFillToTop     As Double
    m_xnumFillToBottom  As Double
    m_cNumStop          As Long
    m_xfillGradientStop() As UcsGradientStopType
End Type

Public Type UcsBiff12BrtBlxfType
    m_dg                As Integer
    '--- padding 2 bytes
    m_brtColor          As UcsBiff12BrtColorType
End Type

Public Type UcsBiff12BrtBorderType
    m_flags             As Long
    m_blxfTop           As UcsBiff12BrtBlxfType
    m_blxfBottom        As UcsBiff12BrtBlxfType
    m_blxfLeft          As UcsBiff12BrtBlxfType
    m_blxfRight         As UcsBiff12BrtBlxfType
    m_blxfDiag          As UcsBiff12BrtBlxfType
End Type

Public Type UcsBiff12BrtXfType
    m_ixfeParent        As Integer
    m_iFmt              As Integer
    m_iFont             As Integer
    m_iFill             As Integer
    m_ixBorder          As Integer
    m_trot              As Byte
    m_indent            As Byte
    m_flags             As Integer
    m_xfGrbitAtr        As Byte
End Type

Public Type UcsBiff12BrtStyleType
    m_ixf               As Long
    m_grbitObj1         As Integer
    m_iStyBuiltIn       As Byte
    m_iLevel            As Byte
    m_stName            As String
End Type

Public Type UcsBiff12BrtWbPropType
    m_flags             As Long
    m_dwThemeVersion    As Long
    m_strName           As String
End Type

Public Type UcsBiff12BrtBookViewType
    m_xWn               As Long
    m_yWn               As Long
    m_dxWn              As Long
    m_dyWn              As Long
    m_iTabRatio         As Long
    m_itabFirst         As Long
    m_itabCur           As Long
    m_flags             As Integer
End Type

Public Type UcsBiff12BrtBundleShType
    m_hsState           As Long
    m_iTabID            As Long
    m_strRelID          As String
    m_strName           As String
End Type

Public Type UcsBiff12BrtWsPropType
    m_flags             As Long
    m_brtcolorTab       As UcsBiff12BrtColorType
    m_rwSync            As Long
    m_colSync           As Long
    m_strName           As String
End Type

Public Type UcsBiff12BrtColInfoType
    m_colFirst          As Long
    m_colLast           As Long
    m_colDx             As Long
    m_ixfe              As Long
    m_flags             As Integer
End Type

Public Type UcsBiff12BrtColSpanType
    m_colMic            As Long
    m_colLast           As Long
End Type

Public Type UcsBiff12BrtRowHdrType
    m_rw                As Long
    m_ixfe              As Long
    m_miyRw             As Integer
    '--- padding
    m_flags             As Long '-- 3 bytes
    m_ccolspan          As Long
    m_rgBrtColspan()    As UcsBiff12BrtColSpanType
End Type

Public Type UcsBiff12BrtFmtType
    m_iFmt              As Integer
    m_stFmtCode         As String
End Type

Public Type UcsBiff12UncheckedRfXType
    m_rwFirst           As Long
    m_rwLast            As Long
    m_colFirst          As Long
    m_colLast           As Long
End Type

Public Type UcsBiff12BrtFileVersionType
    m_guidCodeName      As String
    m_stAppName         As String
    m_stLastEdited      As String
    m_stLowestEdited    As String
    m_stRupBuild        As String
End Type

'=========================================================================
' API
'=========================================================================

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
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

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    FromUtf8Array = String$(2 * UBound(baText), 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
    FromUtf8Array = Left$(FromUtf8Array, lSize)
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
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
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

#End If ' ImplUseShared
