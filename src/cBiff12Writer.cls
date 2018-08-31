VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12Writer"
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
Private Const STR_MODULE_NAME As String = "cBiff12Writer"

#Const ImplUseShared = BIFF12_USESHARED

'=========================================================================
' API
'=========================================================================

'Private Declare Function ApiEmptyLongArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbLong, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Long()

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lColCount             As Long
Private m_lRowCount             As Long
Private m_lCompressionLevel     As Long
Private m_oContainer            As cBiff12Container
Private m_oStringsPart          As cBiff12Part
Private m_oSheetPart            As cBiff12Part
Private m_aColWidth()           As Long
Private m_lWsDimPosition        As Long
Private m_uRowHdr               As UcsBiff12BrtRowHdrType
Private m_aMergeCells()         As UcsBiff12UncheckedRfXType
Private m_lMergeCellsCount      As Long
Private m_oStylesCollection     As cBiff12Part
Private m_bFlushed              As Boolean
'--- debug
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error management
'=========================================================================

Friend Property Get frInstanceName() As String
    frInstanceName = m_sInstanceName
End Property

Private Property Get MODULE_NAME() As String
    #If DebugMode Then
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName, m_sDebugID)
    #Else
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName)
    #End If
End Property

Private Function PrintError(sFunction As String) As VbMsgBoxResult
#If ImplUseShared Then
    PopPrintError sFunction, MODULE_NAME, PushError
#Else
    Debug.Print MODULE_NAME & "." & sFunction & ": " & Err.Description
#End If ' ImplUseShared
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get ColRowFactor() As Long
    ColRowFactor = 260
End Property

Property Get Container() As cBiff12Container
    Set Container = m_oContainer
End Property

Property Get ColCount() As Long
    ColCount = m_lColCount
End Property

Property Get ColWidth(ByVal Index As Long) As Long
    ColWidth = m_aColWidth(Index)
End Property

Property Let ColWidth(ByVal Index As Long, ByVal lValue As Long)
    m_aColWidth(Index) = lValue
End Property

Property Get RowCount() As Long
    RowCount = m_lRowCount
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function init( _
            ByVal lColCount As Long, _
            Optional ByVal UseSST As Boolean, _
            Optional ByVal CompressionLevel As Long = 6, _
            Optional ByVal SheetName As String = "Sheet1") As Boolean
    Set m_oContainer = New cBiff12Container
    m_lColCount = lColCount
    m_lCompressionLevel = CompressionLevel
    ReDim m_aColWidth(0 To m_lColCount) As Long
    If UseSST Then
        Set m_oStringsPart = m_oContainer.StringsPart
    End If
    Set m_oSheetPart = m_oContainer.SheetPart
    m_oSheetPart.SheetName = SheetName
    Set m_oStylesCollection = New cBiff12Part
    pvBeginStyles m_oContainer.StylesPart
    If Not m_oStringsPart Is Nothing Then
        pvBeginStrings m_oStringsPart
    End If
    '--- prepare current row struct
    m_uRowHdr.m_ccolspan = 1
    ReDim m_uRowHdr.m_rgBrtColspan(0 To 0) As UcsBiff12BrtColSpanType
    m_uRowHdr.m_rgBrtColspan(0).m_colLast = m_lColCount - 1
    '--- success
    init = True
End Function

Public Function AddSheet(Optional ByVal SheetName As String = "", Optional ByVal lColCount As Long = 1) As Boolean
    pvEndSheet m_oSheetPart
    Set m_oSheetPart = m_oContainer.SheetPart(m_oContainer.SheetPartCount + 1)
    m_oSheetPart.SheetName = IIf(SheetName <> "", SheetName, "Sheet" & Format(m_oContainer.SheetPartCount, "0"))
    '--- reset previous variables
    m_lRowCount = 0
    m_lColCount = lColCount
    ReDim m_aColWidth(0 To m_lColCount) As Long
    '--- prepare current row struct
    m_uRowHdr.m_ccolspan = 1
    ReDim m_uRowHdr.m_rgBrtColspan(0 To 0) As UcsBiff12BrtColSpanType
    m_uRowHdr.m_rgBrtColspan(0).m_colLast = m_lColCount - 1
    AddSheet = True
End Function

Public Function AddRow(Optional ByVal RowIndex As Long = -1, Optional RowHeight As Long) As Boolean
    Const FUNC_NAME     As String = "AddRow"
    
    On Error GoTo EH
    If m_lColCount = 0 Then
        Err.Raise vbObjectError, , "Not initialized"
    End If
    If RowIndex < 0 Then
        RowIndex = m_lRowCount
    End If
    If m_oSheetPart.Size = 0 Then
        pvBeginSheet m_oSheetPart
    End If
    m_uRowHdr.m_rw = RowIndex
    If RowHeight > 0 Then
        m_uRowHdr.m_miyRw = RowHeight
        Const fUnsynced As Long = 2 ^ 13        '--- 2.4.726 BrtRowHdr: F - fUnsynced (1 bit): A bit that specifies whether the row height has been manually specified.
        m_uRowHdr.m_flags = fUnsynced
    Else
        m_uRowHdr.m_miyRw = 288                 '--- note: this height is ignored (row auto-size)
        m_uRowHdr.m_flags = 0
    End If
    m_oSheetPart.OutputBrtRowHdr m_uRowHdr
    If m_lRowCount < RowIndex + 1 Then
        m_lRowCount = RowIndex + 1
    End If
    '--- success
    AddRow = True
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Public Function AddStringCell(ByVal lColIndex As Long, sText As String, oStyle As cBiff12CellStyle) As Boolean
    Const FUNC_NAME     As String = "AddStringCell"
    
    On Error GoTo EH
    If oStyle.Index < 0 Then
        oStyle.Index = pvGetStyleIndex(oStyle)
    End If
    If LenB(sText) = 0 Then
        m_oSheetPart.OutputCellBlank lColIndex, oStyle.Index
    ElseIf m_oStringsPart Is Nothing Then
        m_oSheetPart.OutputCellRString lColIndex, oStyle.Index, sText
    Else
        m_oSheetPart.OutputCellIsst lColIndex, oStyle.Index, m_oStringsPart.SstGetIndex(sText)
    End If
    '--- success
    AddStringCell = True
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Public Function AddNumberCell(ByVal lColIndex As Long, dblNum As Double, oStyle As cBiff12CellStyle) As Boolean
    Const FUNC_NAME     As String = "AddNumberCell"
    Dim lRkNumber       As Long
    
    On Error GoTo EH
    If oStyle.Index < 0 Then
        oStyle.Index = pvGetStyleIndex(oStyle)
    End If
    If m_oSheetPart.ToRkNumber(dblNum, RetVal:=lRkNumber) Then
        m_oSheetPart.OutputCellRk lColIndex, oStyle.Index, lRkNumber
    Else
        m_oSheetPart.OutputCellReal lColIndex, oStyle.Index, dblNum
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Public Function MergeCells(ByVal lColIndex As Long, ByVal lMergeDown As Long, ByVal lMergeAcross As Long)
    Const FUNC_NAME     As String = "MergeCells"
    
    On Error GoTo EH
    If m_lMergeCellsCount = 0 Then
        ReDim m_aMergeCells(0 To 7) As UcsBiff12UncheckedRfXType
    ElseIf m_lMergeCellsCount > UBound(m_aMergeCells) Then
        ReDim Preserve m_aMergeCells(0 To 2 * m_lMergeCellsCount) As UcsBiff12UncheckedRfXType
    End If
    With m_aMergeCells(m_lMergeCellsCount)
        .m_rwFirst = m_uRowHdr.m_rw
        .m_rwLast = .m_rwFirst + lMergeDown - 1
        .m_colFirst = lColIndex
        .m_colLast = lColIndex + lMergeAcross - 1
    End With
    m_lMergeCellsCount = m_lMergeCellsCount + 1
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Public Function AddImage(ByVal lColIndex As Long, baData() As Byte, ByVal lLeft As Long, ByVal lTOp As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    Const STR_NS3       As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    Const STR_NS4       As String = "http://schemas.microsoft.com/office/drawing/2010/main"
    Dim oDrawingPart    As cBiff12Part
    Dim oImagePart      As cBiff12Part
    Dim oDoc            As Object
    Dim oAttr           As Object
    Dim baTemp()        As Byte
    
    Set oImagePart = m_oContainer.AddImageToPart(m_oSheetPart, DrawingPart:=oDrawingPart)
    baTemp = baData
    oImagePart.Contents = baTemp
    Set oDoc = oDrawingPart.XmlDocument
    With oDoc.DocumentElement.appendChild(oDoc.createElement("xdr:twoCellAnchor"))
        .setAttribute "editAs", "oneCell"
        With .appendChild(oDoc.createElement("xdr:from"))
            .appendChild(oDoc.createElement("xdr:col")).nodeTypedValue = lColIndex
            .appendChild(oDoc.createElement("xdr:colOff")).nodeTypedValue = lLeft
            .appendChild(oDoc.createElement("xdr:row")).nodeTypedValue = m_uRowHdr.m_rw
            .appendChild(oDoc.createElement("xdr:rowOff")).nodeTypedValue = lTOp
        End With
        With .appendChild(oDoc.createElement("xdr:to"))
            .appendChild(oDoc.createElement("xdr:col")).nodeTypedValue = lColIndex
            .appendChild(oDoc.createElement("xdr:colOff")).nodeTypedValue = lLeft + lWidth
            .appendChild(oDoc.createElement("xdr:row")).nodeTypedValue = m_uRowHdr.m_rw
            .appendChild(oDoc.createElement("xdr:rowOff")).nodeTypedValue = lTOp + lHeight
        End With
        With .appendChild(oDoc.createElement("xdr:pic"))
            With .appendChild(oDoc.createElement("xdr:nvPicPr"))
                With .appendChild(oDoc.createElement("xdr:cNvPr"))
                    .setAttribute "id", m_oContainer.ImagePartsCount
                    .setAttribute "name", "Picture " & m_oContainer.ImagePartsCount
                End With
                With .appendChild(oDoc.createElement("xdr:cNvPicPr"))
                    With .appendChild(oDoc.createElement("a:picLocks"))
                        .setAttribute "noChangeAspect", 1
                    End With
                End With
            End With
            With .appendChild(oDoc.createElement("xdr:blipFill"))
                With .appendChild(oDoc.createElement("a:blip"))
                    Set oAttr = oDoc.createNode(2, "r:embed", STR_NS3)
                    oAttr.NodeValue = m_oContainer.GetRelID(oDrawingPart, oImagePart)
                    .setAttributeNode oAttr
                    .setAttribute "cstate", "print"
                    With .appendChild(oDoc.createElement("a:extLst"))
                        With .appendChild(oDoc.createElement("a:ext"))
                            .setAttribute "uri", "{28A0092B-C50C-407E-A947-70E740481C1C}"
                            With .appendChild(oDoc.createNode(1, "a14:useLocalDpi", STR_NS4))
                                .setAttribute "val", 0
                            End With
                        End With
                    End With
                End With
                With .appendChild(oDoc.createElement("a:stretch"))
                    .appendChild oDoc.createElement("a:fillRect")
                End With
            End With
            With .appendChild(oDoc.createElement("xdr:spPr"))
                With .appendChild(oDoc.createElement("a:xfrm"))
                    With .appendChild(oDoc.createElement("a:off"))
                        .setAttribute "x", lLeft
                        .setAttribute "y", lTOp
                    End With
                    With .appendChild(oDoc.createElement("a:ext"))
                        .setAttribute "cx", lWidth
                        .setAttribute "cy", lHeight
                    End With
                End With
                With .appendChild(oDoc.createElement("a:prstGeom"))
                    .setAttribute "prst", "rect"
                    .appendChild oDoc.createElement("a:avLst")
                End With
            End With
        End With
        .appendChild oDoc.createElement("xdr:clientData")
    End With
End Function

Public Sub Flush()
    If Not m_bFlushed Then
        pvEndSheet m_oSheetPart
        If Not m_oStringsPart Is Nothing Then
            pvEndStrings m_oStringsPart
        End If
        pvEndStyles m_oContainer.StylesPart
        pvEndWorkbook m_oContainer.WorkbookPart
        m_bFlushed = True
    End If
End Sub

Public Function SaveToFile(OutputFile As Variant) As Boolean
    Const FUNC_NAME     As String = "SaveToFile"
    
    On Error GoTo EH
    Flush
    If Not m_oContainer.SaveToFile(OutputFile, Level:=m_lCompressionLevel) Then
        GoTo QH
    End If
    '--- success
    SaveToFile = True
QH:
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

'= private ===============================================================

Private Sub pvBeginStrings(oPart As cBiff12Part)
    With oPart
        .OutputCount2 ucsBrtBeginSst, 0, 0
    End With
End Sub

Private Sub pvEndStrings(oPart As cBiff12Part)
    With oPart
        .Output ucsBrtEndSst
    End With
End Sub

Private Sub pvBeginSheet(oPart As cBiff12Part)
    Dim uWsProp         As UcsBiff12BrtWsPropType
    Dim uWsDim          As UcsBiff12UncheckedRfXType
    Dim uColInfo        As UcsBiff12BrtColInfoType
    Dim lIdx            As Long
    
    With oPart
        '--- WORKSHEET = BrtBeginSheet [BrtWsProp] [BrtWsDim] [WSVIEWS2] [WSFMTINFO] *COLINFOS CELLTABLE [BrtSheetCalcProp] [[BrtSheetProtectionIso] BrtSheetProtection]
        '---    *([BrtRangeProtectionIso] BrtRangeProtection) [SCENMAN] [AUTOFILTER] [SORTSTATE] [DCON] [USERSHVIEWS] [MERGECELLS] [BrtPhoneticInfo] *CONDITIONALFORMATTING
        '---    [DVALS] *([ACUID] BrtHLink) [BrtPrintOptions] [BrtMargins] [BrtPageSetup] [HEADERFOOTER] [RWBRK] [COLBRK] *BrtBigName [CELLWATCHES] [IGNOREECS] [SMARTTAGS]
        '---    [BrtDrawing] [BrtLegacyDrawing] [BrtLegacyDrawingHF] [BrtBkHim] [OLEOBJECTS] [ACTIVEXCONTROLS] [WEBPUBITEMS] [LISTPARTS] FRTWORKSHEET [ACUID] BrtEndSheet
        .Output ucsBrtBeginSheet
#If False Then
        '--- worksheet props
        uWsProp.m_flags = &H204C9
        uWsProp.m_brtcolorTab.m_index = 64
        uWsProp.m_rwSync = -1
        uWsProp.m_colSync = -1
        .OutputBrtWsProp uWsProp
#End If
        '--- worksheet dimensions
        uWsDim.m_colLast = m_lColCount - 1
        m_lWsDimPosition = .OutputBrtWsDim(uWsDim)
        '--- COLINFOS = BrtBeginColInfos 1*BrtColInfo BrtEndColInfos
        For lIdx = 0 To m_lColCount - 1
            If m_aColWidth(lIdx) <> 0 Then
                Exit For
            End If
        Next
        If lIdx < m_lColCount Then
            .Output ucsBrtBeginColInfos
            For lIdx = lIdx To m_lColCount - 1
                If m_aColWidth(lIdx) <> 0 Then
                    uColInfo.m_colFirst = lIdx
                    uColInfo.m_colLast = lIdx
                    uColInfo.m_colDx = m_aColWidth(lIdx)
                    .OutputBrtColInfo uColInfo
                End If
            Next
            .Output ucsBrtEndColInfos
        End If
        .OutputCount3 ucsBrtWsFmtInfo, -1, &H1200008, 0
        '--- CELLTABLE = BrtBeginSheetData *1048576([ACCELLTABLE] BrtRowHdr *16384CELL *FRT) BrtEndSheetData
        .Output ucsBrtBeginSheetData
    End With
End Sub

Private Sub pvEndSheet(oPart As cBiff12Part)
    Dim lIdx            As Long
    Dim oDrawingPart    As cBiff12Part
    
    With oPart
        .Output ucsBrtEndSheetData
        '--- MERGECELLS
        If m_lMergeCellsCount > 0 Then
            .OutputCount ucsBrtBeginMergeCells, m_lMergeCellsCount
            For lIdx = 0 To m_lMergeCellsCount - 1
                .OutputBrtMergeCell m_aMergeCells(lIdx)
            Next
            .Output ucsBrtEndMergeCells
        End If
        '--- BrtDrawing
        If m_oContainer.GetImagesCount(m_oSheetPart, SheetDrawingPart:=oDrawingPart) > 0 Then
            .OutputBrtDrawing m_oContainer.GetRelID(m_oSheetPart, oDrawingPart)
        End If
        .Output ucsBrtEndSheet
        '--- fix WsDim record
        If m_lRowCount > 0 Then
            .Position = m_lWsDimPosition + 4
            .WriteDWord m_lRowCount - 1
        End If
    End With
End Sub

Private Sub pvBeginStyles(oPart As cBiff12Part)
    Dim oStyle          As cBiff12CellStyle
    
    With oPart
        Set oStyle = New cBiff12CellStyle
        '--- ToDo: This becomes default style for cells outside specified worksheet dimensions
        pvGetStyleIndex oStyle
    End With
End Sub

Private Sub pvEndStyles(oPart As cBiff12Part)
    Dim uFmt            As UcsBiff12BrtFmtType
    Dim uFont           As UcsBiff12BrtFontType
    Dim uFill           As UcsBiff12BrtFillType
    Dim uBorder         As UcsBiff12BrtBorderType
    Dim uXf             As UcsBiff12BrtXfType
    Dim oStyle          As cBiff12CellStyle
    Dim lIdx            As Long
    Dim vStyles         As Variant
    Dim vElem           As Variant
    Dim cUnique         As Collection
    Dim sKey            As String
    Dim oItem           As cBiff12CellStyle
    Dim lAlc            As Long
    Dim lAlcv           As Long
    
    '--- re-hydrate cell styles from shared strings hash-table
    vStyles = m_oStylesCollection.SstDumpTable()
    For lIdx = 0 To UBound(vStyles)
        Set oStyle = New cBiff12CellStyle
        oStyle.Contents = vStyles(lIdx)
        Set vStyles(lIdx) = oStyle
    Next
    ' 2.1.7.50 Styles: STYLESHEET = BrtBeginStyleSheet [FMTS] [FONTS] [FILLS] [BORDERS] CELLSTYLEXFS CELLXFS STYLES DXFS TABLESTYLES [COLORPALETTE] FRTSTYLESHEET BrtEndStyleSheet
    With oPart
        .Output ucsBrtBeginStyleSheet
        '--- collect unique custom formats
        Set cUnique = New Collection
        For Each vElem In vStyles
            Set oStyle = vElem
            If LenB(oStyle.Format) <> 0 Then
                sKey = oStyle.Format & "#"
                If SearchCollection(cUnique, sKey, RetVal:=oItem) Then
                    oStyle.FmtIndex = oItem.FmtIndex
                Else
                    '--- 2.4.657 BrtFmt: ifmt - custom number format 164 to 382
                    oStyle.FmtIndex = 164 + cUnique.Count
                    cUnique.Add oStyle, sKey
                End If
            End If
        Next
        '--- 2.1.7.50 Styles: FMTS = BrtBeginFmts 1*(2ACFMT / BrtFmt) BrtEndFmts
        If cUnique.Count > 0 Then
            .OutputCount ucsBrtBeginFmts, cUnique.Count
            For Each oStyle In cUnique
                uFmt.m_iFmt = oStyle.FmtIndex
                uFmt.m_stFmtCode = oStyle.Format
                .OutputBrtFmt uFmt
            Next
            .Output ucsBrtEndFmts
        End If
        '--- collect unique fonts (idx=0 - font of default style)
        Set cUnique = New Collection
        For Each vElem In vStyles
            Set oStyle = vElem
            With oStyle
                sKey = .FontName & "#" & .FontSize & "#" & -.Bold & "#" & -.Italic & "#" & .ForeColor
            End With
            If SearchCollection(cUnique, sKey, RetVal:=oItem) Then
                oStyle.FontIndex = oItem.FontIndex
            Else
                oStyle.FontIndex = cUnique.Count
                cUnique.Add oStyle, sKey
            End If
        Next
        '--- 2.1.7.50 Styles: FONTS = BrtBeginFonts 1*65491 BrtFont [ACFONTS] BrtEndFonts
        If cUnique.Count > 0 Then
            .OutputCount ucsBrtBeginFonts, cUnique.Count
            For Each oStyle In cUnique
                uFont.m_dyHeight = oStyle.FontSize * 20                     '--- 20 * Points = Twips (https://support.microsoft.com/en-us/help/76388)
                uFont.m_grbit = -oStyle.Italic * 2                          '--- 2.5.52 FontFlags: B - fItalic (1 bit): A bit that specifies whether the text is italic.
                uFont.m_bls = IIf(oStyle.Bold, 700, 400)
                uFont.m_bFamily = 0                                         '--- 2.4.659 BrtFont: bFamily=0 - Not applicable
                uFont.m_bCharSet = 1                                        '--- 2.4.659 BrtFont: bCharSet=1 - DEFAULT_CHARSET
                uFont.m_brtColor = pvToBrtColor(oStyle.ForeColor)
                uFont.m_bFontScheme = 0                                     '--- 2.4.659 BrtFont: bFontScheme=0 - None
                uFont.m_name = oStyle.FontName
                .OutputBrtFont uFont
            Next
            .Output ucsBrtEndFonts
        End If
        '--- collect unique fills (idx=0 - fill of default style)
        Set cUnique = New Collection
        For Each vElem In vStyles
            Set oStyle = vElem
            sKey = oStyle.BackColor & "#"
            If SearchCollection(cUnique, sKey, RetVal:=oItem) Then
                oStyle.FillIndex = oItem.FillIndex
            Else
                oStyle.FillIndex = cUnique.Count + 2
                cUnique.Add oStyle, sKey
            End If
        Next
        '--- 2.1.7.50 Styles: FILLS = BrtBeginFills 1*65431 BrtFill BrtEndFills
        If cUnique.Count > 0 Then
            .OutputCount ucsBrtBeginFills, cUnique.Count
            '--- system iFill=0
            uFill.m_fls = 0                                 ' none (no fill)
            uFill.m_brtColorFore = pvToBrtColor(vbBlack)
            uFill.m_brtColorBack = pvToBrtColor(vbWhite)
            .OutputBrtFill uFill
            '--- system iFill=1
            uFill.m_fls = &H11                              ' grayscale of 0.125 (1/8) value
            uFill.m_brtColorFore = pvToBrtColor(vbBlack)
            uFill.m_brtColorBack = pvToBrtColor(vbWhite)
            .OutputBrtFill uFill
            For Each oStyle In cUnique
                uFill.m_fls = IIf(oStyle.BackColor = -1, 0, 1)              '--- 2.4.650 BrtFill: 0 - The fill pattern is none (no fill), 1 - The fill pattern is solid
                uFill.m_brtColorFore = pvToBrtColor(oStyle.BackColor)       '--- 2.4.650 BrtFill: When solid is specified, brtColorFore is the only color rendered, even when brtColorBack is also specified
                uFill.m_brtColorBack = pvToBrtColor(vbWhite, 0)
                .OutputBrtFill uFill
            Next
            .Output ucsBrtEndFills
        End If
        '--- collect unique borders (idx=0 - borders of default style)
        Set cUnique = New Collection
        For Each vElem In vStyles
            Set oStyle = vElem
            With oStyle
                sKey = .BorderLeftColor & "#" & .BorderTopColor & "#" & .BorderRightColor & "#" & .BorderBottomColor
            End With
            If SearchCollection(cUnique, sKey, RetVal:=oItem) Then
                oStyle.BorderIndex = oItem.BorderIndex
            Else
                oStyle.BorderIndex = cUnique.Count
                cUnique.Add oStyle, sKey
            End If
        Next
        '--- 2.1.7.50 Styles: BORDERS = BrtBeginBorders 1*65430 BrtBorder BrtEndBorders
        If cUnique.Count > 0 Then
            .OutputCount ucsBrtBeginBorders, cUnique.Count
            For Each vElem In cUnique
                Set oStyle = vElem
                uBorder.m_flags = 0
                With uBorder.m_blxfTop
                    .m_dg = IIf(oStyle.BorderTopColor = -1, 0, 1)           '--- 2.5.4 Blxf: 0 - None, 1 - Thin
                    .m_brtColor = pvToBrtColor(oStyle.BorderTopColor)
                End With
                With uBorder.m_blxfBottom
                    .m_dg = IIf(oStyle.BorderBottomColor = -1, 0, 1)        '--- 2.5.4 Blxf: 0 - None, 1 - Thin
                    .m_brtColor = pvToBrtColor(oStyle.BorderBottomColor)
                End With
                With uBorder.m_blxfLeft
                    .m_dg = IIf(oStyle.BorderLeftColor = -1, 0, 1)          '--- 2.5.4 Blxf: 0 - None, 1 - Thin
                    .m_brtColor = pvToBrtColor(oStyle.BorderLeftColor)
                End With
                With uBorder.m_blxfRight
                    .m_dg = IIf(oStyle.BorderRightColor = -1, 0, 1)         '--- 2.5.4 Blxf: 0 - None, 1 - Thin
                    .m_brtColor = pvToBrtColor(oStyle.BorderRightColor)
                End With
                With uBorder.m_blxfDiag
                    .m_dg = 0                                               '--- 2.5.4 Blxf: 0 - None
                    .m_brtColor = pvToBrtColor(-1)
                End With
                .OutputBrtBorder uBorder
            Next
            .Output ucsBrtEndBorders
        End If
        '--- 2.1.7.50 Styles: CELLSTYLEXFS = BrtBeginCellStyleXFs 1*65430( BrtXF *FRT) BrtEndCellStyleXFs
        .OutputCount ucsBrtBeginCellStyleXFs, 1
            uXf.m_ixfeParent = -1               '--- 2.4.824 BrtXF: If this record specifies a cell style XF, the value MUST be 0xFFFF
            lAlc = 0                            '--- 2.4.824 BrtXF: alc=0 - General alignment
            lAlcv = 2                           '--- 2.4.824 BrtXF: alcv=2 - Bottom alignment
            Const fWrap As Long = 2 ^ 6         '--- 2.4.824 BrtXF: A - fWrap (1 bit): A bit that specifies whether the text in a cell is line-wrapped within the cell.
            Const fShrinkToFit  As Long = 2 ^ 8 '--- 2.4.824 BrtXF: C - fShrinkToFit (1 bit): A bit that specifies whether the displayed text in the cell is shrink to fit.
            Const fLocked As Long = 2 ^ 12      '--- 2.4.824 BrtXF: F - fLocked (1 bit): A bit that specifies whether the locked protection property is set to true.
            uXf.m_flags = lAlc Or (lAlcv * 8) Or (-oStyle.WrapText * fWrap) Or (-oStyle.ShrinkText * fShrinkToFit) Or fLocked
            .OutputBrtXf uXf
        .Output ucsBrtEndCellStyleXFs
        '--- 2.1.7.50 Styles: CELLXFS = BrtBeginCellXFs 1*65430(BrtXF *FRT) BrtEndCellXFs
        .OutputCount ucsBrtBeginCellXFs, UBound(vStyles) + 1
        For Each vElem In vStyles
            Set oStyle = vElem
            uXf.m_ixfeParent = 0                '--- 2.4.824 BrtXF: If the record specifies a cell XF, the value is a zero-based index to another BrtXF record in the collection of all records directly following BrtBeginCellStyleXFs (section 2.4.18) which is a cell style XF.
            uXf.m_iFmt = oStyle.FmtIndex
            uXf.m_iFont = oStyle.FontIndex
            uXf.m_iFill = oStyle.FillIndex
            uXf.m_ixBorder = oStyle.BorderIndex
            uXf.m_trot = 0                      '--- 2.4.824 BrtXF: 0-90 - Text rotated counterclockwise 0 to 90 degrees
            uXf.m_indent = oStyle.LeftIndent    '--- 2.4.824 BrtXF: An unsigned integer that specifies the indentation level for text in a cell.
            Select Case oStyle.HorAlign         '--- 2.4.824 BrtXF: alc=1 - Left alignment, alc=2 - Center Alignment, alc=3 - Right alignment
            Case ucsHalCenter
                lAlc = 2
            Case ucsHalRight
                lAlc = 3
            Case Else
                lAlc = 0
            End Select
            Select Case oStyle.VertAlign        '--- 2.4.824 BrtXF: alcv=0 - Top alignment, alcv=1 - Center alignment, alcv=2 - Bottom alignment
            Case ucsValMiddle
                lAlcv = 1
            Case ucsValBottom
                lAlcv = 2
            Case Else
                lAlcv = 0
            End Select
            uXf.m_flags = lAlc Or (lAlcv * 8) Or (-oStyle.WrapText * fWrap) Or (-oStyle.ShrinkText * fShrinkToFit) Or fLocked
            uXf.m_xfGrbitAtr = 0
            .OutputBrtXf uXf
        Next
        .Output ucsBrtEndCellXFs
        '--- 2.1.7.50 Styles: DXFS = BrtBeginDXFs *2147483647DXF BrtEndDXFs
        .OutputCount ucsBrtBeginDXFs, 0
        .Output ucsBrtEndDXFs
        '--- 2.1.7.50 Styles: TABLESTYLES = BrtBeginTableStyles *TABLESTYLE BrtEndTableStyles
        .OutputCount3 ucsBrtBeginTableStyles, 0, 0, 0
        .Output ucsBrtEndTableStyles
        .Output ucsBrtEndStyleSheet
    End With
End Sub

Private Sub pvEndWorkbook(oPart As cBiff12Part)
    Dim uFileVersion    As UcsBiff12BrtFileVersionType
    Dim uBundle         As UcsBiff12BrtBundleShType
    Dim uOleSize        As UcsBiff12UncheckedRfXType
    Dim i As Long
    
    With oPart
        .Output ucsBrtBeginBook
#If False Then
            uFileVersion.m_stAppName = "xl"
            uFileVersion.m_stLastEdited = "6"
            uFileVersion.m_stLowestEdited = "6"
            uFileVersion.m_stRupBuild = "14420"
            .OutputBrtFileVersion uFileVersion
#End If
            .Output ucsBrtBeginBundleShs
            If m_oContainer.SheetPartCount > 0 Then
              For i = 1 To m_oContainer.SheetPartCount
                Set m_oSheetPart = m_oContainer.SheetPart(i)
                uBundle.m_iTabID = i
                uBundle.m_strRelID = m_oContainer.GetRelID(oPart, m_oSheetPart)
                uBundle.m_strName = m_oSheetPart.SheetName
                .OutputBrtBundleSh uBundle
              Next
            End If
            .Output ucsBrtEndBundleShs
            uOleSize.m_rwLast = m_lRowCount - 1
            uOleSize.m_colLast = m_lColCount - 1
            .OutputBrtOleSize uOleSize
        .Output ucsBrtEndBook
    End With
End Sub

Private Function pvGetStyleIndex(oStyle As cBiff12CellStyle) As Long
    oStyle.Index = m_oStylesCollection.SstGetIndex(oStyle.Contents)
    pvGetStyleIndex = oStyle.Index
End Function

Private Function pvToBrtColor(ByVal clrValue As Long, Optional ByVal lAlpha As Long = 255) As UcsBiff12BrtColorType
    Const fValidRGB     As Long = 1
    
    With pvToBrtColor
        Select Case clrValue
        Case vbBlack
            .m_xColorType = 1 * 2 + fValidRGB   '--- 2.4.324 BrtColor: xColorType=1 - color from a color palette and is specified by index + fValidRGB=1 - The color specified by index matches the color specified by RGBA.
            .m_index = 64
        Case vbWhite
            .m_xColorType = 1 * 2 + fValidRGB   '--- 2.4.324 BrtColor: xColorType=1 - color from a color palette and is specified by index + fValidRGB=1 - The color specified by index matches the color specified by RGBA.
            .m_index = 65
        Case -1
            .m_xColorType = 0 * 2 + fValidRGB   '--- 2.4.324 BrtColor: xColorType=0 - Color information is automatically determined by the application + fValidRGB=1 - The color specified by index matches the color specified by RGBA.
            .m_index = 0
        Case Else
            .m_xColorType = 2 * 2 + fValidRGB   '--- 2.4.324 BrtColor: xColorType=2 - ARGB color + fValidRGB=0 - The color specified by index does not match the color specified by RGBA
            .m_index = 255                      '--- 2.4.324 BrtColor: xColorType=2 - index is undefined and MUST be ignored
        End Select
        .m_nTintAndShade = 0                    '--- 2.4.324 BrtColor: nTintAndShade=0 - means no change
        If clrValue >= 0 Then
            .m_bRed = clrValue And &HFF&
            .m_bGreen = (clrValue And &HFF00&) \ &H100&
            .m_bBlue = (clrValue And &HFF0000) \ &H10000
            .m_bAlpha = lAlpha
        End If
    End With
End Function

#If ImplUseShared = 0 Then

Private Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

#End If ' ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit STR_MODULE_NAME, m_sDebugID, Me
    End Sub

    Private Sub Class_Terminate()
        DebugInstanceTerm STR_MODULE_NAME, m_sDebugID
    End Sub
#End If
