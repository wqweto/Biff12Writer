VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12CellStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Biff12Writer (c) 2017 by wqweto@gmail.com
'
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cBiff12CellStyle"

#Const ImplUseShared = BIFF12_USESHARED

'=========================================================================
' Public enums
'=========================================================================

#If ImplUseShared = 0 Then

Public Enum UcsHorAlignmentEnum
    ucsHalLeft = 0
    ucsHalRight = 1
    ucsHalCenter = 2
End Enum

#End If

'=========================================================================
' API
'=========================================================================

Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal Length As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lIndex                    As Long
Private m_sFontName                 As String
Private m_sFormat                   As String
Private m_uState                    As UcsCellStyleType
Private m_lFmtIndex                 As Long
Private m_lFontIndex                As Long
Private m_lFillIndex                As Long
Private m_lBorderIndex              As Long

Private Type UcsCellStyleType
    FontName            As String * 32
    FontSize            As Double
    Bold                As Boolean
    Italic              As Boolean
    ForeColor           As Long
    BackColor           As Long
    HorAlign            As UcsHorAlignmentEnum
    BorderLeftColor     As Long
    BorderTopColor      As Long
    BorderRightColor    As Long
    BorderBottomColor   As Long
    WrapText            As Boolean
    Indent              As Long
    Format              As String * 64
End Type

'=========================================================================
' Properties
'=========================================================================

Property Get FontName() As String
    FontName = m_sFontName
End Property

Property Let FontName(sValue As String)
    If m_sFontName <> sValue Then
        m_sFontName = sValue
        m_uState.FontName = sValue
        m_lIndex = -1
    End If
End Property

Property Get FontSize() As Double
    FontSize = m_uState.FontSize
End Property

Property Let FontSize(dblValue As Double)
    If m_uState.FontSize <> dblValue Then
        m_uState.FontSize = dblValue
        m_lIndex = -1
    End If
End Property

Property Get Bold() As Boolean
    Bold = m_uState.Bold
End Property

Property Let Bold(ByVal bValue As Boolean)
    If m_uState.Bold <> bValue Then
        m_uState.Bold = bValue
        m_lIndex = -1
    End If
End Property

Property Get Italic() As Boolean
    Italic = m_uState.Italic
End Property

Property Let Italic(ByVal bValue As Boolean)
    If m_uState.Italic <> bValue Then
        m_uState.Italic = bValue
        m_lIndex = -1
    End If
End Property

Property Get ForeColor() As Long
    ForeColor = m_uState.ForeColor
End Property

Property Let ForeColor(ByVal lValue As Long)
    If m_uState.ForeColor <> lValue Then
        m_uState.ForeColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get BackColor() As Long
    BackColor = m_uState.BackColor
End Property

Property Let BackColor(ByVal lValue As Long)
    If m_uState.BackColor <> lValue Then
        m_uState.BackColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get HorAlign() As UcsHorAlignmentEnum
    HorAlign = m_uState.HorAlign
End Property

Property Let HorAlign(ByVal eValue As UcsHorAlignmentEnum)
    If m_uState.HorAlign <> eValue Then
        m_uState.HorAlign = eValue
        m_lIndex = -1
    End If
End Property

Property Get BorderLeftColor() As Long
    BorderLeftColor = m_uState.BorderLeftColor
End Property

Property Let BorderLeftColor(ByVal lValue As Long)
    If m_uState.BorderLeftColor <> lValue Then
        m_uState.BorderLeftColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get BorderTopColor() As Long
    BorderTopColor = m_uState.BorderTopColor
End Property

Property Let BorderTopColor(ByVal lValue As Long)
    If m_uState.BorderTopColor <> lValue Then
        m_uState.BorderTopColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get BorderRightColor() As Long
    BorderRightColor = m_uState.BorderRightColor
End Property

Property Let BorderRightColor(ByVal lValue As Long)
    If m_uState.BorderRightColor <> lValue Then
        m_uState.BorderRightColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get BorderBottomColor() As Long
    BorderBottomColor = m_uState.BorderBottomColor
End Property

Property Let BorderBottomColor(ByVal lValue As Long)
    If m_uState.BorderBottomColor <> lValue Then
        m_uState.BorderBottomColor = lValue
        m_lIndex = -1
    End If
End Property

Property Get WrapText() As Boolean
    WrapText = m_uState.WrapText
End Property

Property Let WrapText(ByVal bValue As Boolean)
    If m_uState.WrapText <> bValue Then
        m_uState.WrapText = bValue
        m_lIndex = -1
    End If
End Property

Property Get LeftIndent() As Long
    LeftIndent = m_uState.Indent
End Property

Property Let LeftIndent(ByVal lValue As Long)
    If m_uState.Indent <> lValue Then
        m_uState.Indent = lValue
        m_lIndex = -1
    End If
End Property

Property Get Format() As String
    Format = m_sFormat
End Property

Property Let Format(sValue As String)
    If m_sFormat <> sValue Then
        m_sFormat = sValue
        m_uState.Format = sValue
        m_lIndex = -1
    End If
End Property

'= friend ================================================================

Friend Property Get Index() As Long
    Index = m_lIndex
End Property

Friend Property Let Index(ByVal lValue As Long)
    m_lIndex = lValue
End Property

Friend Property Get Contents() As String
    Call PutMem4(ByVal VarPtr(Contents), ByVal SysAllocStringLen(ByVal VarPtr(m_uState), LenB(m_uState) \ 2))
End Property

Friend Property Let Contents(sValue As String)
    Call CopyMemory(ByVal VarPtr(m_uState), ByVal StrPtr(sValue), LenB(m_uState))
    m_sFontName = RTrim$(m_uState.FontName)
    m_sFormat = RTrim$(m_uState.Format)
End Property

Friend Property Get FmtIndex() As Long
    FmtIndex = m_lFmtIndex
End Property

Friend Property Let FmtIndex(ByVal lValue As Long)
    m_lFmtIndex = lValue
End Property

Friend Property Get FontIndex() As Long
    FontIndex = m_lFontIndex
End Property

Friend Property Let FontIndex(ByVal lValue As Long)
    m_lFontIndex = lValue
End Property

Friend Property Get FillIndex() As Long
    FillIndex = m_lFillIndex
End Property

Friend Property Let FillIndex(ByVal lValue As Long)
    m_lFillIndex = lValue
End Property

Friend Property Get BorderIndex() As Long
    BorderIndex = m_lBorderIndex
End Property

Friend Property Let BorderIndex(ByVal lValue As Long)
    m_lBorderIndex = lValue
End Property

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    '--- prevent fixed-length strings filled with Chr(0)
    m_uState.FontName = vbNullString
    m_uState.Format = vbNullString
    With New StdFont
        FontName = .Name
        FontSize = .Size
    End With
    ForeColor = vbBlack
    '--- use -1 for transparent
    BackColor = -1
    BorderLeftColor = -1
    BorderTopColor = -1
    BorderRightColor = -1
    BorderBottomColor = -1
End Sub
