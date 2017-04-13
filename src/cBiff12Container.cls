VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12Container"
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
Private Const MODULE_NAME As String = "cBiff12Container"

#Const ImplUseShared = BIFF12_USESHARED

'=========================================================================
' API
'=========================================================================

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function ApiEmptyBiff12PartArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbObject, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As cBiff12Part()
#If ImplUseShared = 0 Then
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
#End If ' ImplUseShared

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ROOT                  As String = "/[Content_Types].xml"
Private Const STR_XML_PROLOG            As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf

Private m_uState                    As UcsStateType

Private Type UcsRelationType
    Id                  As String
    RelType             As String
    RelTarget           As String
End Type

Private Type UcsPartType
    PartName            As String
    PartType            As String
    Contents            As Object
    ContentType         As String
    Relations()         As UcsRelationType
    RelationsIndex      As Collection
    ImagesParts()       As cBiff12Part
End Type

Private Type UcsStateType
    Parts()             As UcsPartType
    PartsIndex          As Collection
    CorePropsPart       As cBiff12Part
    AppPropsPart        As cBiff12Part
    WorkbookPart        As cBiff12Part
    StylesPart          As cBiff12Part
    ThemeParts()        As cBiff12Part
    StringsPart         As cBiff12Part
    SheetParts()        As cBiff12Part
    DrawingParts()      As cBiff12Part
    ImagesParts()       As cBiff12Part
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
#If ImplUseShared Then
    PopRaiseError sFunction, MODULE_NAME, PushError
#Else
    Debug.Print MODULE_NAME & "." & sFunction & ": " & Err.Description
#End If ' ImplUseShared
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get CorePropsPart() As cBiff12Part
    Set CorePropsPart = m_uState.CorePropsPart
    If CorePropsPart Is Nothing Then
        Set CorePropsPart = AddXmlPart("/docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml")
        Set m_uState.CorePropsPart = CorePropsPart
        pvAddRelation STR_ROOT, pvPathDifference(STR_ROOT, CorePropsPart.Name), CorePropsPart.Type_
    End If
End Property

Property Get AppPropsPart() As cBiff12Part
    Set AppPropsPart = m_uState.AppPropsPart
    If AppPropsPart Is Nothing Then
        Set AppPropsPart = AddXmlPart("/docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
        Set m_uState.AppPropsPart = AppPropsPart
        pvAddRelation STR_ROOT, pvPathDifference(STR_ROOT, AppPropsPart.Name), AppPropsPart.Type_
    End If
End Property

Property Get WorkbookPart() As cBiff12Part
    Dim lIdx            As Long
    
    Set WorkbookPart = m_uState.WorkbookPart
    If WorkbookPart Is Nothing Then
        Set WorkbookPart = AddBinPart("/xl/workbook.bin", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.ms-excel.workbook")
        Set m_uState.WorkbookPart = WorkbookPart
        pvAddRelation STR_ROOT, pvPathDifference(STR_ROOT, WorkbookPart.Name), WorkbookPart.Type_
        AddRelation WorkbookPart, m_uState.StylesPart
        For lIdx = 0 To UBound(m_uState.ThemeParts)
            AddRelation WorkbookPart, m_uState.ThemeParts(lIdx)
        Next
        AddRelation WorkbookPart, m_uState.StringsPart
        For lIdx = 0 To UBound(m_uState.SheetParts)
            AddRelation WorkbookPart, m_uState.SheetParts(lIdx)
        Next
    End If
End Property

Property Get StylesPart() As cBiff12Part
    Set StylesPart = m_uState.StylesPart
    If StylesPart Is Nothing Then
        Set StylesPart = AddBinPart("/xl/styles.bin", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.ms-excel.styles")
        Set m_uState.StylesPart = StylesPart
        AddRelation m_uState.WorkbookPart, StylesPart
    End If
End Property

Property Get ThemePart(Optional ByVal Index As Long = 1) As cBiff12Part
    If UBound(m_uState.ThemeParts) >= Index Then
        Set ThemePart = m_uState.ThemeParts(Index)
    End If
    If ThemePart Is Nothing Then
        Set ThemePart = AddXmlPart(Replace("/xl/theme/theme{0}.xml", "{0}", Index), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "application/vnd.openxmlformats-officedocument.theme+xml")
        If UBound(m_uState.ThemeParts) < Index Then
            ReDim Preserve m_uState.ThemeParts(0 To Index) As cBiff12Part
        End If
        Set m_uState.ThemeParts(Index) = ThemePart
        AddRelation m_uState.WorkbookPart, ThemePart
    End If
End Property

Property Get StringsPart() As cBiff12Part
    Set StringsPart = m_uState.StringsPart
    If StringsPart Is Nothing Then
        Set StringsPart = AddBinPart("/xl/sharedStrings.bin", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "application/vnd.ms-excel.sharedStrings")
        Set m_uState.StringsPart = StringsPart
        AddRelation m_uState.WorkbookPart, StringsPart
    End If
End Property

Property Get SheetPart(Optional ByVal Index As Long = 1) As cBiff12Part
    If UBound(m_uState.SheetParts) >= Index Then
        Set SheetPart = m_uState.SheetParts(Index)
    End If
    If SheetPart Is Nothing Then
        Set SheetPart = AddBinPart(Replace("/xl/worksheets/sheet{0}.bin", "{0}", Index), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "application/vnd.ms-excel.worksheet")
        If UBound(m_uState.SheetParts) < Index Then
            ReDim Preserve m_uState.SheetParts(0 To Index) As cBiff12Part
        End If
        Set m_uState.SheetParts(Index) = SheetPart
        AddRelation m_uState.WorkbookPart, SheetPart
    End If
End Property

Property Get DrawingPart(Optional ByVal Index As Long = 1) As cBiff12Part
    Const STR_NS1       As String = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    Const STR_NS2       As String = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
    If UBound(m_uState.DrawingParts) >= Index Then
        Set DrawingPart = m_uState.DrawingParts(Index)
    End If
    If DrawingPart Is Nothing Then
        Set DrawingPart = AddXmlPart(Replace("/xl/drawings/drawing{0}.xml", "{0}", Index), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "application/vnd.openxmlformats-officedocument.drawing+xml")
        If UBound(m_uState.DrawingParts) < Index Then
            ReDim Preserve m_uState.DrawingParts(0 To Index) As cBiff12Part
        End If
        Set m_uState.DrawingParts(Index) = DrawingPart
        AddRelation SheetPart(Index), DrawingPart
        With DrawingPart.XmlDocument
            With .appendChild(.createNode(1, "xdr:wsDr", STR_NS1))
                .setAttribute "xmlns:a", STR_NS2
            End With
        End With
    End If
End Property

Property Get ImagePartsCount() As Long
    ImagePartsCount = UBound(m_uState.ImagesParts) + 1
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AddBinPart(sPartName As String, sPartType As String, sContentType As String) As cBiff12Part
    Set AddBinPart = New cBiff12Part
    AddBinPart.Name = sPartName
    AddBinPart.Type_ = sPartType
    AddPart sPartName, sPartType, AddBinPart, sContentType
End Function

Public Function AddXmlPart(sPartName As String, sPartType As String, sContentType As String) As cBiff12Part
    Set AddXmlPart = New cBiff12Part
    AddXmlPart.Name = sPartName
    AddXmlPart.Type_ = sPartType
    Set AddXmlPart.XmlDocument = CreateObject("MSXML2.DOMDocument")
    AddPart sPartName, sPartType, AddXmlPart, sContentType
End Function

Public Function AddPart(sPartName As String, sPartType As String, vContents As Variant, sContentType As String) As Long
    Dim lIdx            As Long
    
    With m_uState
        If Peek(ArrPtr(.Parts)) = 0 Then
            ReDim .Parts(0 To 0) As UcsPartType
            Set .PartsIndex = New Collection
        Else
            ReDim Preserve .Parts(0 To UBound(.Parts) + 1) As UcsPartType
        End If
        lIdx = UBound(.Parts)
        .PartsIndex.Add lIdx, sPartName
        With .Parts(lIdx)
            .PartName = sPartName
            .PartType = sPartType
            AssignVariant .Contents, vContents
            .ContentType = sContentType
        End With
        '--- retval
        AddPart = lIdx
    End With
End Function

Public Function AddImageToSheet(oSheetPart As cBiff12Part, Optional SheetDrawingPart As cBiff12Part) As cBiff12Part
    Dim lIdx            As Long
    Dim lCount          As Long
    
    lIdx = m_uState.PartsIndex.Item(oSheetPart.Name)
    If SheetDrawingPart Is Nothing Then
        Set SheetDrawingPart = DrawingPart(lIdx)
    End If
    lCount = UBound(m_uState.ImagesParts) + 1
    Set AddImageToSheet = AddBinPart(Replace("/xl/media/image{0}.png", "{0}", lCount + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "image/png")
    ReDim Preserve m_uState.ImagesParts(0 To lCount)
    Set m_uState.ImagesParts(lCount) = AddImageToSheet
    AddRelation SheetDrawingPart, AddImageToSheet
    '--- append to sheet images
    With m_uState.Parts(lIdx)
        If Peek(ArrPtr(.ImagesParts)) = 0 Then
            lCount = 0
            ReDim .ImagesParts(0 To lCount) As cBiff12Part
        Else
            lCount = UBound(.ImagesParts) + 1
            ReDim Preserve .ImagesParts(0 To lCount) As cBiff12Part
        End If
        Set .ImagesParts(lCount) = AddImageToSheet
    End With
End Function

Public Function AddRelation(oPart As cBiff12Part, oTarget As cBiff12Part) As String
    If Not oPart Is Nothing And Not oTarget Is Nothing Then
        AddRelation = pvAddRelation(oPart.Name, pvPathDifference(oPart.Name, oTarget.Name), oTarget.Type_)
    End If
End Function

Public Function GetRelID(oPart As cBiff12Part, oTarget As cBiff12Part) As String
    Dim lIdx            As Long
    Dim sPartName       As String
    Dim sRelTarget      As String
    
    sPartName = oPart.Name
    sRelTarget = pvPathDifference(sPartName, oTarget.Name)
    With m_uState.Parts(m_uState.PartsIndex.Item(sPartName))
        If SearchCollection(.RelationsIndex, sRelTarget, RetVal:=lIdx) Then
            GetRelID = .Relations(lIdx).Id
        End If
    End With
End Function

Public Function GetImagesCount(oSheetPart As cBiff12Part, Optional SheetDrawingPart As cBiff12Part) As Long
    Dim lIdx            As Long
    
    lIdx = m_uState.PartsIndex.Item(oSheetPart.Name)
    With m_uState.Parts(lIdx)
        If Peek(ArrPtr(.ImagesParts)) <> 0 Then
            GetImagesCount = UBound(.ImagesParts) + 1
            Set SheetDrawingPart = DrawingPart(lIdx)
        End If
    End With
End Function

Public Function SaveToFile(OutputFile As Variant, Optional ByVal Level As Long) As Boolean
    Const STR_NS        As String = "http://schemas.openxmlformats.org/package/2006/content-types"
    Dim oDoc            As Object
    Dim oRoot           As Object
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim cDefaults       As Collection
    Dim cOverwrite      As Collection
    Dim sExt            As String
    Dim oArchive        As cZipArchive
    
    '--- collect defaults & overwrites for root
    Set cDefaults = New Collection
    cDefaults.Add "bin=application/vnd.ms-excel.sheet.binary.macroEnabled.main", "bin"
    cDefaults.Add "rels=application/vnd.openxmlformats-package.relationships+xml", "rels"
    cDefaults.Add "xml=application/xml", "xml"
    cDefaults.Add "png=image/png", "png"
    Set cOverwrite = New Collection
    For lIdx = 1 To UBound(m_uState.Parts)
        With m_uState.Parts(lIdx)
            If InStrRev(.PartName, ".") > InStrRev(.PartName, "/") Then
                sExt = Mid$(.PartName, InStrRev(.PartName, ".") + 1)
                If Not SearchCollection(cDefaults, sExt, RetVal:=vElem) Then
                    cDefaults.Add sExt & "=" & .ContentType, sExt
                Else
                    vElem = Split(vElem, "=")
                    If vElem(1) <> .ContentType And .ContentType <> "application/vnd.ms-excel.workbook" Then
                        cOverwrite.Add .PartName & "=" & .ContentType
                    End If
                End If
            End If
        End With
    Next
    '--- construct root [Content_Types].xml
    Set oDoc = m_uState.Parts(0).Contents.XmlDocument
    oDoc.loadXml vbNullString
    Set oRoot = oDoc.appendChild(oDoc.createNode(1, "Types", STR_NS))
    For Each vElem In cDefaults
        vElem = Split(vElem, "=")
        With oRoot.appendChild(oDoc.createNode(1, "Default", STR_NS))
            .setAttribute "Extension", vElem(0)
            .setAttribute "ContentType", vElem(1)
        End With
    Next
    For Each vElem In cOverwrite
        vElem = Split(vElem, "=")
        With oRoot.appendChild(oDoc.createNode(1, "Override", STR_NS))
            .setAttribute "PartName", vElem(0)
            .setAttribute "ContentType", vElem(1)
        End With
    Next
    '--- output all parts to archive
    Set oArchive = New cZipArchive
    For lIdx = 0 To UBound(m_uState.Parts)
        With m_uState.Parts(lIdx)
            oArchive.AddFile .Contents, Mid$(Replace(.PartName, "/", "\"), 2)
            pvWriteRelations oArchive, IIf(lIdx = 0, "\", Replace(.PartName, "/", "\")), .Relations
        End With
    Next
    If oArchive.CompressArchive(OutputFile, Level:=Level) Then
        SaveToFile = True
    Else
        Err.Raise vbObjectError, , oArchive.LastError
    End If
End Function

'= private ===============================================================

Private Function pvAddRelation(sPartName As String, sRelTarget As String, sRelType As String) As String
    Dim lIdx            As Long
    
    With m_uState.Parts(m_uState.PartsIndex.Item(sPartName))
        If Not SearchCollection(.RelationsIndex, sRelTarget, RetVal:=lIdx) Then
            If Peek(ArrPtr(.Relations)) = 0 Then
                ReDim .Relations(0 To 0) As UcsRelationType
                Set .RelationsIndex = New Collection
            Else
                ReDim Preserve .Relations(0 To UBound(.Relations) + 1) As UcsRelationType
            End If
            lIdx = UBound(.Relations)
            .RelationsIndex.Add lIdx, sRelTarget
            With .Relations(lIdx)
                .Id = "rId" & (lIdx + 1)
                .RelType = sRelType
                .RelTarget = sRelTarget
            End With
        End If
        '--- retval
        pvAddRelation = .Relations(lIdx).Id
    End With
End Function

Private Sub pvWriteRelations(oArchive As cZipArchive, sFile As String, uRels() As UcsRelationType)
    Const FUNC_NAME     As String = "pvWriteRelations"
    Const STR_NS        As String = "http://schemas.openxmlformats.org/package/2006/relationships"
    Dim lIdx            As Long
    Dim oDoc            As Object
    Dim oRoot           As Object
    Dim oPart           As cBiff12Part
    
    On Error GoTo EH
    If Peek(ArrPtr(uRels)) <> 0 Then
        Set oDoc = CreateObject("MSXML2.DOMDocument")
        Set oRoot = oDoc.appendChild(oDoc.createNode(1, "Relationships", STR_NS))
        For lIdx = 0 To UBound(uRels)
            With oRoot.appendChild(oDoc.createNode(1, "Relationship", STR_NS))
                .setAttribute "Id", uRels(lIdx).Id
                .setAttribute "Type", uRels(lIdx).RelType
                .setAttribute "Target", uRels(lIdx).RelTarget
            End With
        Next
        Set oPart = New cBiff12Part
        oPart.Contents = ToUtf8Array(STR_XML_PROLOG & oDoc.xml)
        lIdx = InStrRev(sFile, "\")
        oArchive.AddFile oPart, Mid$(Left$(sFile, lIdx) & "_rels\" & Mid$(sFile, lIdx + 1) & ".rels", 2)
    End If
    Exit Sub
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Sub

Private Function pvPathDifference(sSource As String, sTarget As String) As String
    Const FUNC_NAME     As String = "pvPathDifference"
    Dim vBase           As Variant
    Dim vTarget         As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    On Error GoTo EH
    If InStrRev(sSource, "/") = 0 Then
        pvPathDifference = sTarget
    Else
        vBase = Split(Left$(sSource, InStrRev(sSource, "/") - 1), "/")
        vTarget = Split(sTarget, "/")
        For lIdx = 0 To UBound(vTarget)
            If lIdx <= UBound(vBase) Then
                If LCase$(vBase(lIdx)) <> LCase$(vTarget(lIdx)) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        For lJdx = lIdx To UBound(vBase)
            pvPathDifference = pvPathDifference & IIf(LenB(pvPathDifference) <> 0, "/", vbNullString) & ".."
        Next
        For lJdx = lIdx To UBound(vTarget)
            pvPathDifference = pvPathDifference & IIf(LenB(pvPathDifference) <> 0, "/", vbNullString) & vTarget(lJdx)
        Next
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

#If ImplUseShared = 0 Then

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Private Function ToUtf8Array(sText As String) As Byte()
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

Private Sub Class_Initialize()
    m_uState.ThemeParts = ApiEmptyBiff12PartArray()
    m_uState.SheetParts = ApiEmptyBiff12PartArray()
    m_uState.DrawingParts = ApiEmptyBiff12PartArray()
    m_uState.ImagesParts = ApiEmptyBiff12PartArray()
    AddXmlPart STR_ROOT, "package", vbNullString
End Sub

