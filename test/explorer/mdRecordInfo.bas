Attribute VB_Name = "mdRecordInfo"
'=========================================================================
'
' Biff12Writer (c) 2017 by wqweto@gmail.com
'
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_BRT_LOOKUP As String = "RowHdr|0|CellBlank|1|CellRk|2|CellError|3|CellBool|4|CellReal|5|CellSt|6|CellIsst|7|FmlaString|8|FmlaNum|9|FmlaBool|A|FmlaError|B|FRTArchID|10|SSTItem|13|PCDIMissing|14|PCDINumber|15|PCDIBoolean|16|PCDIError|17|PCDIString|18|PCDIDatetime|19|PCDIIndex|1A|PCDIAMissing|1B|PCDIANumber|1C|PCDIABoolean|1D|PCDIAError|1E|PCDIAString|1F|PCDIADatetime|20|PCRRecord|21|PCRRecordDt|22|FRTBegin|23|FRTEnd|24|ACBegin|25|ACEnd|26|Name|27|IndexRowBlock|28|IndexBlock|2A|Font|2B|Fmt|2C|Fill|2D|Border|2E|XF|2F|Style|30|CellMeta|31|ValueMeta|32|Mdb|33|BeginFmd|34|EndFmd|35|BeginMdx|36|EndMdx|37|BeginMdxTuple|38|EndMdxTuple|39|MdxMbrIstr|3A|" & _
        "Str|3B|ColInfo|3C|CellRString|3E|CalcChainItem|3F|DVal|40|SxvcellNum|41|SxvcellStr|42|SxvcellBool|43|SxvcellErr|44|SxvcellDate|45|SxvcellNil|46|FileVersion|80|BeginSheet|81|EndSheet|82|BeginBook|83|EndBook|84|BeginWsViews|85|EndWsViews|86|BeginBookViews|87|EndBookViews|88|BeginWsView|89|EndWsView|8A|BeginCsViews|8B|EndCsViews|8C|BeginCsView|8D|EndCsView|8E|BeginBundleShs|8F|EndBundleShs|90|BeginSheetData|91|EndSheetData|92|WsProp|93|WsDim|94|Pane|97|Sel|98|WbProp|99|WbFactoid|9A|FileRecover|9B|BundleSh|9C|CalcProp|9D|BookView|9E|BeginSst|9F|EndSst|A0|BeginAFilter|A1|EndAFilter|A2|BeginFilterColumn|A3|EndFilterColumn|A4|BeginFilters|A5|EndFilters|A6|Filter|A7|" & _
        "ColorFilter|A8|IconFilter|A9|Top10Filter|AA|DynamicFilter|AB|BeginCustomFilters|AC|EndCustomFilters|AD|CustomFilter|AE|AFilterDateGroupItem|AF|MergeCell|B0|BeginMergeCells|B1|EndMergeCells|B2|BeginPivotCacheDef|B3|EndPivotCacheDef|B4|BeginPCDFields|B5|EndPCDFields|B6|BeginPCDField|B7|EndPCDField|B8|BeginPCDSource|B9|EndPCDSource|BA|BeginPCDSRange|BB|EndPCDSRange|BC|BeginPCDFAtbl|BD|EndPCDFAtbl|BE|BeginPCDIRun|BF|EndPCDIRun|C0|BeginPivotCacheRecords|C1|EndPivotCacheRecords|C2|BeginPCDHierarchies|C3|EndPCDHierarchies|C4|BeginPCDHierarchy|C5|EndPCDHierarchy|C6|BeginPCDHFieldsUsage|C7|EndPCDHFieldsUsage|C8|BeginExtConnection|C9|EndExtConnection|CA|BeginECDbProps|CB|EndECDbProps|CC|BeginECOlapProps|CD|EndECOlapProps|CE|" & _
        "BeginPCDSConsol|CF|EndPCDSConsol|D0|BeginPCDSCPages|D1|EndPCDSCPages|D2|BeginPCDSCPage|D3|EndPCDSCPage|D4|BeginPCDSCPItem|D5|EndPCDSCPItem|D6|BeginPCDSCSets|D7|EndPCDSCSets|D8|BeginPCDSCSet|D9|EndPCDSCSet|DA|BeginPCDFGroup|DB|EndPCDFGroup|DC|BeginPCDFGItems|DD|EndPCDFGItems|DE|BeginPCDFGRange|DF|EndPCDFGRange|E0|BeginPCDFGDiscrete|E1|EndPCDFGDiscrete|E2|BeginPCDSDTupleCache|E3|EndPCDSDTupleCache|E4|BeginPCDSDTCEntries|E5|EndPCDSDTCEntries|E6|BeginPCDSDTCEMembers|E7|EndPCDSDTCEMembers|E8|BeginPCDSDTCEMember|E9|EndPCDSDTCEMember|EA|BeginPCDSDTCQueries|EB|EndPCDSDTCQueries|EC|BeginPCDSDTCQuery|ED|EndPCDSDTCQuery|EE|BeginPCDSDTCSets|EF|EndPCDSDTCSets|F0|BeginPCDSDTCSet|F1|EndPCDSDTCSet|F2|BeginPCDCalcItems|F3|" & _
        "EndPCDCalcItems|F4|BeginPCDCalcItem|F5|EndPCDCalcItem|F6|BeginPRule|F7|EndPRule|F8|BeginPRFilters|F9|EndPRFilters|FA|BeginPRFilter|FB|EndPRFilter|FC|BeginPNames|FD|EndPNames|FE|BeginPName|FF|EndPName|100|BeginPNPairs|101|EndPNPairs|102|BeginPNPair|103|EndPNPair|104|BeginECWebProps|105|EndECWebProps|106|BeginEcWpTables|107|EndECWPTables|108|BeginECParams|109|EndECParams|10A|BeginECParam|10B|EndECParam|10C|BeginPCDKPIs|10D|EndPCDKPIs|10E|BeginPCDKPI|10F|EndPCDKPI|110|BeginDims|111|EndDims|112|BeginDim|113|EndDim|114|IndexPartEnd|115|BeginStyleSheet|116|EndStyleSheet|117|BeginSXView|118|EndSXVI|119|BeginSXVI|11A|BeginSXVIs|11B|EndSXVIs|11C|BeginSXVD|11D|EndSXVD|11E|BeginSXVDs|11F|" & _
        "EndSXVDs|120|BeginSXPI|121|EndSXPI|122|BeginSXPIs|123|EndSXPIs|124|BeginSXDI|125|EndSXDI|126|BeginSXDIs|127|EndSXDIs|128|BeginSXLI|129|EndSXLI|12A|BeginSXLIRws|12B|EndSXLIRws|12C|BeginSXLICols|12D|EndSXLICols|12E|BeginSXFormat|12F|EndSXFormat|130|BeginSXFormats|131|EndSxFormats|132|BeginSxSelect|133|EndSxSelect|134|BeginISXVDRws|135|EndISXVDRws|136|BeginISXVDCols|137|EndISXVDCols|138|EndSXLocation|139|BeginSXLocation|13A|EndSXView|13B|BeginSXTHs|13C|EndSXTHs|13D|BeginSXTH|13E|EndSXTH|13F|BeginISXTHRws|140|EndISXTHRws|141|BeginISXTHCols|142|EndISXTHCols|143|BeginSXTDMPS|144|EndSXTDMPs|145|BeginSXTDMP|146|EndSXTDMP|147|BeginSXTHItems|148|EndSXTHItems|149|BeginSXTHItem|14A|EndSXTHItem|14B|" & _
        "BeginMetadata|14C|EndMetadata|14D|BeginEsmdtinfo|14E|Mdtinfo|14F|EndEsmdtinfo|150|BeginEsmdb|151|EndEsmdb|152|BeginEsfmd|153|EndEsfmd|154|BeginSingleCells|155|EndSingleCells|156|BeginList|157|EndList|158|BeginListCols|159|EndListCols|15A|BeginListCol|15B|EndListCol|15C|BeginListXmlCPr|15D|EndListXmlCPr|15E|ListCCFmla|15F|ListTrFmla|160|BeginExternals|161|EndExternals|162|SupBookSrc|163|SupSelf|165|SupSame|166|SupTabs|167|BeginSupBook|168|PlaceholderName|169|ExternSheet|16A|ExternTableStart|16B|ExternTableEnd|16C|ExternRowHdr|16E|ExternCellBlank|16F|ExternCellReal|170|ExternCellBool|171|ExternCellError|172|ExternCellString|173|BeginEsmdx|174|EndEsmdx|175|BeginMdxSet|176|EndMdxSet|177|" & _
        "BeginMdxMbrProp|178|EndMdxMbrProp|179|BeginMdxKPI|17A|EndMdxKPI|17B|BeginEsstr|17C|EndEsstr|17D|BeginPRFItem|17E|EndPRFItem|17F|BeginPivotCacheIDs|180|EndPivotCacheIDs|181|BeginPivotCacheID|182|EndPivotCacheID|183|BeginISXVIs|184|EndISXVIs|185|BeginColInfos|186|EndColInfos|187|BeginRwBrk|188|EndRwBrk|189|BeginColBrk|18A|EndColBrk|18B|Brk|18C|UserBookView|18D|Info|18E|CUsr|18F|Usr|190|BeginUsers|191|EOF|193|UCR|194|RRInsDel|195|RREndInsDel|196|RRMove|197|RREndMove|198|RRChgCell|199|RREndChgCell|19A|RRHeader|19B|RRUserView|19C|RRRenSheet|19D|RRInsertSh|19E|RRDefName|19F|RRNote|1A0|RRConflict|1A1|RRTQSIF|1A2|RRFormat|1A3|RREndFormat|1A4|RRAutoFmt|1A5|BeginUserShViews|1A6|" & _
        "BeginUserShView|1A7|EndUserShView|1A8|EndUserShViews|1A9|ArrFmla|1AA|ShrFmla|1AB|Table|1AC|BeginExtConnections|1AD|EndExtConnections|1AE|BeginPCDCalcMems|1AF|EndPCDCalcMems|1B0|BeginPCDCalcMem|1B1|EndPCDCalcMem|1B2|BeginPCDHGLevels|1B3|EndPCDHGLevels|1B4|BeginPCDHGLevel|1B5|EndPCDHGLevel|1B6|BeginPCDHGLGroups|1B7|EndPCDHGLGroups|1B8|BeginPCDHGLGroup|1B9|EndPCDHGLGroup|1BA|BeginPCDHGLGMembers|1BB|EndPCDHGLGMembers|1BC|BeginPCDHGLGMember|1BD|EndPCDHGLGMember|1BE|BeginQSI|1BF|EndQSI|1C0|BeginQSIR|1C1|EndQSIR|1C2|BeginDeletedNames|1C3|EndDeletedNames|1C4|BeginDeletedName|1C5|EndDeletedName|1C6|BeginQSIFs|1C7|EndQSIFs|1C8|BeginQSIF|1C9|EndQSIF|1CA|BeginAutoSortScope|1CB|EndAutoSortScope|1CC|BeginConditionalFormatting|1CD|" & _
        "EndConditionalFormatting|1CE|BeginCFRule|1CF|EndCFRule|1D0|BeginIconSet|1D1|EndIconSet|1D2|BeginDatabar|1D3|EndDatabar|1D4|BeginColorScale|1D5|EndColorScale|1D6|CFVO|1D7|ExternValueMeta|1D8|BeginColorPalette|1D9|EndColorPalette|1DA|IndexedColor|1DB|Margins|1DC|PrintOptions|1DD|PageSetup|1DE|BeginHeaderFooter|1DF|EndHeaderFooter|1E0|BeginSXCrtFormat|1E1|EndSXCrtFormat|1E2|BeginSXCrtFormats|1E3|EndSXCrtFormats|1E4|WsFmtInfo|1E5|BeginMgs|1E6|EndMGs|1E7|BeginMGMaps|1E8|EndMGMaps|1E9|BeginMG|1EA|EndMG|1EB|BeginMap|1EC|EndMap|1ED|HLink|1EE|BeginDCon|1EF|EndDCon|1F0|BeginDRefs|1F1|EndDRefs|1F2|DRef|1F3|BeginScenMan|1F4|EndScenMan|1F5|BeginSct|1F6|EndSct|1F7|Slc|1F8|BeginDXFs|1F9|EndDXFs|1FA|" & _
        "DXF|1FB|BeginTableStyles|1FC|EndTableStyles|1FD|BeginTableStyle|1FE|EndTableStyle|1FF|TableStyleElement|200|TableStyleClient|201|BeginVolDeps|202|EndVolDeps|203|BeginVolType|204|EndVolType|205|BeginVolMain|206|EndVolMain|207|BeginVolTopic|208|EndVolTopic|209|VolSubtopic|20A|VolRef|20B|VolNum|20C|VolErr|20D|VolStr|20E|VolBool|20F|BeginCalcChain|210|EndCalcChain|211|BeginSortState|212|EndSortState|213|BeginSortCond|214|EndSortCond|215|BookProtection|216|SheetProtection|217|RangeProtection|218|PhoneticInfo|219|BeginECTxtWiz|21A|EndECTxtWiz|21B|BeginECTWFldInfoLst|21C|EndECTWFldInfoLst|21D|BeginECTwFldInfo|21E|FileSharing|224|OleSize|225|Drawing|226|LegacyDrawing|227|LegacyDrawingHF|228|WebOpt|229|" & _
        "BeginWebPubItems|22A|EndWebPubItems|22B|BeginWebPubItem|22C|EndWebPubItem|22D|BeginSXCondFmt|22E|EndSXCondFmt|22F|BeginSXCondFmts|230|EndSXCondFmts|231|BkHim|232|Color|234|BeginIndexedColors|235|EndIndexedColors|236|BeginMRUColors|239|EndMRUColors|23A|MRUColor|23C|BeginDVals|23D|EndDVals|23E|SupNameStart|241|SupNameValueStart|242|SupNameValueEnd|243|SupNameNum|244|SupNameErr|245|SupNameSt|246|SupNameNil|247|SupNameBool|248|SupNameFmla|249|SupNameBits|24A|SupNameEnd|24B|EndSupBook|24C|CellSmartTagProperty|24D|BeginCellSmartTag|24E|EndCellSmartTag|24F|BeginCellSmartTags|250|EndCellSmartTags|251|BeginSmartTags|252|EndSmartTags|253|SmartTagType|254|BeginSmartTagTypes|255|EndSmartTagTypes|256|BeginSXFilters|257|" & _
        "EndSXFilters|258|BeginSXFILTER|259|EndSXFilter|25A|BeginFills|25B|EndFills|25C|BeginCellWatches|25D|EndCellWatches|25E|CellWatch|25F|BeginCRErrs|260|EndCRErrs|261|CrashRecErr|262|BeginFonts|263|EndFonts|264|BeginBorders|265|EndBorders|266|BeginFmts|267|EndFmts|268|BeginCellXFs|269|EndCellXFs|26A|BeginStyles|26B|EndStyles|26C|BigName|271|BeginCellStyleXFs|272|EndCellStyleXFs|273|BeginComments|274|EndComments|275|BeginCommentAuthors|276|EndCommentAuthors|277|CommentAuthor|278|BeginCommentList|279|EndCommentList|27A|BeginComment|27B|EndComment|27C|CommentText|27D|BeginOleObjects|27E|OleObject|27F|EndOleObjects|280|BeginSxrules|281|EndSxRules|282|BeginActiveXControls|283|ActiveX|284|EndActiveXControls|285|" & _
        "BeginPCDSDTCEMembersSortBy|286|BeginCellIgnoreECs|288|CellIgnoreEC|289|EndCellIgnoreECs|28A|CsProp|28B|CsPageSetup|28C|BeginUserCsViews|28D|EndUserCsViews|28E|BeginUserCsView|28F|EndUserCsView|290|BeginPcdSFCIEntries|291|EndPCDSFCIEntries|292|PCDSFCIEntry|293|BeginListParts|294|ListPart|295|EndListParts|296|SheetCalcProp|297|BeginFnGroup|298|FnGroup|299|EndFnGroup|29A|SupAddin|29B|SXTDMPOrder|29C|CsProtection|29D|BeginWsSortMap|29F|EndWsSortMap|2A0|BeginRRSort|2A1|EndRRSort|2A2|RRSortItem|2A3|FileSharingIso|2A4|BookProtectionIso|2A5|SheetProtectionIso|2A6|CsProtectionIso|2A7|RangeProtectionIso|2A8|RwDescent|400|KnownFonts|401|BeginSXTupleSet|402|EndSXTupleSet|403|BeginSXTupleSetHeader|404|EndSXTupleSetHeader|405|" & _
        "SXTupleSetHeaderItem|406|BeginSXTupleSetData|407|EndSXTupleSetData|408|BeginSXTupleSetRow|409|EndSXTupleSetRow|40A|SXTupleSetRowItem|40B|NameExt|40C|PCDH14|40D|BeginPCDCalcMem14|40E|EndPCDCalcMem14|40F|SXTH14|410|BeginSparklineGroup|411|EndSparklineGroup|412|Sparkline|413|SXDI14|414|WsFmtInfoEx14|415|BeginConditionalFormatting14|416|EndConditionalFormatting14|417|BeginCFRule14|418|EndCFRule14|419|CFVO14|41A|BeginDatabar14|41B|BeginIconSet14|41C|DVal14|41D|BeginDVals14|41E|Color14|41F|BeginSparklines|420|EndSparklines|421|BeginSparklineGroups|422|EndSparklineGroups|423|SXVD14|425|BeginSxview14|426|EndSxview14|427|BeginPCD14|42A|EndPCD14|42B|BeginExtConn14|42C|EndExtConn14|42D|BeginSlicerCacheIDs|42E|EndSlicerCacheIDs|42F|"
Private Const STR_BRT_LOOKUP2 As String = "BeginSlicerCacheID|430|EndSlicerCacheID|431|BeginSlicerCache|433|EndSlicerCache|434|BeginSlicerCacheDef|435|EndSlicerCacheDef|436|BeginSlicersEx|437|EndSlicersEx|438|BeginSlicerEx|439|EndSlicerEx|43A|BeginSlicer|43B|EndSlicer|43C|SlicerCachePivotTables|43D|BeginSlicerCacheOlapImpl|43E|EndSlicerCacheOlapImpl|43F|BeginSlicerCacheLevelsData|440|EndSlicerCacheLevelsData|441|BeginSlicerCacheLevelData|442|EndSlicerCacheLevelData|443|BeginSlicerCacheSiRanges|444|EndSlicerCacheSiRanges|445|BeginSlicerCacheSiRange|446|EndSlicerCacheSiRange|447|SlicerCacheOlapItem|448|BeginSlicerCacheSelections|449|SlicerCacheSelection|44A|EndSlicerCacheSelections|44B|BeginSlicerCacheNative|44C|EndSlicerCacheNative|44D|SlicerCacheNativeItem|44E|RangeProtection14|44F|" & _
        "RangeProtectionIso14|450|CellIgnoreEC14|451|List14|457|CFIcon|458|BeginSlicerCachesPivotCacheIDs|459|EndSlicerCachesPivotCacheIDs|45A|BeginSlicers|45B|EndSlicers|45C|WbProp14|45D|BeginSXEdit|45E|EndSXEdit|45F|BeginSXEdits|460|EndSXEdits|461|BeginSXChange|462|EndSXChange|463|BeginSXChanges|464|EndSXChanges|465|SXTupleItems|466|BeginSlicerStyle|468|EndSlicerStyle|469|SlicerStyleElement|46A|BeginStyleSheetExt14|46B|EndStyleSheetExt14|46C|BeginSlicerCachesPivotCacheID|46D|EndSlicerCachesPivotCacheID|46E|BeginConditionalFormattings|46F|EndConditionalFormattings|470|BeginPCDCalcMemExt|471|EndPCDCalcMemExt|472|BeginPCDCalcMemsExt|473|EndPCDCalcMemsExt|474|PCDField14|475|BeginSlicerStyles|476|EndSlicerStyles|477|BeginSlicerStyleElements|478|" & _
        "EndSlicerStyleElements|479|CFRuleExt|47A|BeginSXCondFmt14|47B|EndSXCondFmt14|47C|BeginSXCondFmts14|47D|EndSXCondFmts14|47E|BeginSortCond14|480|EndSortCond14|481|EndDVals14|482|EndIconSet14|483|EndDatabar14|484|BeginColorScale14|485|EndColorScale14|486|BeginSxrules14|487|EndSxrules14|488|BeginPRule14|489|EndPRule14|48A|BeginPRFilters14|48B|EndPRFilters14|48C|BeginPRFilter14|48D|EndPRFilter14|48E|BeginPRFItem14|48F|EndPRFItem14|490|BeginCellIgnoreECs14|491|EndCellIgnoreECs14|492|Dxf14|493|BeginDxF14s|494|EndDxf14s|495|Filter14|499|BeginCustomFilters14|49A|CustomFilter14|49C|IconFilter14|49D|PivotCacheConnectionName|49E|BeginDecoupledPivotCacheIDs|800|EndDecoupledPivotCacheIDs|801|DecoupledPivotCacheID|802|BeginPivotTableRefs|803|" & _
        "EndPivotTableRefs|804|PivotTableRef|805|SlicerCacheBookPivotTables|806|BeginSxvcells|807|EndSxvcells|808|BeginSxRow|809|EndSxRow|80A|PcdCalcMem15|80C|Qsi15|813|BeginWebExtensions|814|EndWebExtensions|815|WebExtension|816|AbsPath15|817|BeginPivotTableUISettings|818|EndPivotTableUISettings|819|TableSlicerCacheIDs|81B|TableSlicerCacheID|81C|BeginTableSlicerCache|81D|EndTableSlicerCache|81E|SxFilter15|81F|BeginTimelineCachePivotCacheIDs|820|EndTimelineCachePivotCacheIDs|821|TimelineCachePivotCacheID|822|BeginTimelineCacheIDs|823|EndTimelineCacheIDs|824|BeginTimelineCacheID|825|EndTimelineCacheID|826|BeginTimelinesEx|827|EndTimelinesEx|828|BeginTimelineEx|829|EndTimelineEx|82A|WorkBookPr15|82B|PCDH15|82C|BeginTimelineStyle|82D|EndTimelineStyle|82E|" & _
        "TimelineStyleElement|82F|BeginTimelineStylesheetExt15|830|EndTimelineStylesheetExt15|831|BeginTimelineStyles|832|EndTimelineStyles|833|BeginTimelineStyleElements|834|EndTimelineStyleElements|835|Dxf15|836|BeginDxfs15|837|EndDxfs15|838|SlicerCacheHideItemsWithNoData|839|BeginItemUniqueNames|83A|EndItemUniqueNames|83B|ItemUniqueName|83C|BeginExtConn15|83D|EndExtConn15|83E|BeginOledbPr15|83F|EndOledbPr15|840|BeginDataFeedPr15|841|EndDataFeedPr15|842|TextPr15|843|RangePr15|844|DbCommand15|845|BeginDbTables15|846|EndDbTables15|847|DbTable15|848|BeginDataModel|849|EndDataModel|84A|BeginModelTables|84B|EndModelTables|84C|ModelTable|84D|BeginModelRelationships|84E|EndModelRelationships|84F|ModelRelationship|850|BeginECTxtWiz15|851|" & _
        "EndECTxtWiz15|852|BeginECTWFldInfoLst15|853|EndECTWFldInfoLst15|854|BeginECTWFldInfo15|855|FieldListActiveItem|856|PivotCacheIdVersion|857|SXDI15|858"


Private m_aBrtName()        As String
Private m_bInitBrtName      As Boolean

'=========================================================================
' Function
'=========================================================================

Public Function GetBrtName(ByVal eRecID As UcsBiff12RecortTypeEnum) As String
    Dim vSplit          As Variant
    Dim lIdx            As Long

    If Not m_bInitBrtName Then
        m_bInitBrtName = True
        ReDim m_aBrtName(0 To &HFFF) As String
        vSplit = Split(STR_BRT_LOOKUP & STR_BRT_LOOKUP2, "|")
        For lIdx = 0 To UBound(vSplit) Step 2
            If Left$(vSplit(lIdx), 3) <> "" Then
                vSplit(lIdx) = vSplit(lIdx)
            End If
            m_aBrtName(CLng("&H" & vSplit(lIdx + 1))) = vSplit(lIdx)
        Next
    End If
    GetBrtName = m_aBrtName(eRecID)
    If LenB(GetBrtName) = 0 Then
        GetBrtName = "0x" & Hex$(eRecID)
    End If
End Function

Public Function GetBrtData(ByVal eRecID As UcsBiff12RecortTypeEnum, ByVal lRecSize As Long, oBin As cBiff12Part) As Variant
    Dim cRetVal         As Collection
    Dim lPtr            As Long
    Dim lFlags          As Long
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim sText           As String

    Set cRetVal = New Collection
    lPtr = oBin.Ptr(oBin.Position)
    Select Case eRecID
    Case ucsBrtBeginSst
        cRetVal.Add "cstTotal=" & oBin.ReadDWord()
        cRetVal.Add "cstUnique=" & oBin.ReadDWord()
    Case ucsBrtSSTItem
        pvDumpRichString vbNullString, oBin, cRetVal
    Case ucsBrtBeginFmts
        cRetVal.Add "cfmts=" & oBin.ReadDWord()
    Case ucsBrtFmt
        cRetVal.Add "ifmt=" & oBin.ReadWord()
        cRetVal.Add "stFmtCode=" & oBin.ReadString()
    Case ucsBrtBeginFonts
        cRetVal.Add "cfonts=" & oBin.ReadDWord()
    Case ucsBrtFont
        cRetVal.Add "dyHeight=" & oBin.ReadWord()
        cRetVal.Add "grbit=" & pvFormatFlags(oBin.ReadWord(), "|fItalic||fStrikeout|fOutline|fShadow|fCondense|fExtend")
        cRetVal.Add "bls=" & oBin.ReadWord()
        cRetVal.Add "sss=" & oBin.ReadWord()
        cRetVal.Add "uls=" & oBin.ReadByte()
        cRetVal.Add "bFamily=" & oBin.ReadByte()
        cRetVal.Add "bCharSet=" & oBin.ReadByte()
        oBin.ReadByte '--- unused
        pvDumpColor vbNullString, oBin, cRetVal
        cRetVal.Add "bFontScheme=" & oBin.ReadByte()
        cRetVal.Add "name=" & oBin.ReadString()
    Case ucsBrtBeginFills
        cRetVal.Add "cfills=" & oBin.ReadDWord()
    Case ucsBrtFill
        cRetVal.Add "fls=0x" & Hex$(oBin.ReadDWord())
        pvDumpColor "brtColorFore.", oBin, cRetVal
        pvDumpColor "brtColorBack.", oBin, cRetVal
        cRetVal.Add "iGradientType=" & oBin.ReadDWord()
        cRetVal.Add "xnumDegree=" & oBin.ReadDouble()
        cRetVal.Add "xnumFillToLeft=" & oBin.ReadDouble()
        cRetVal.Add "xnumFillToRight=" & oBin.ReadDouble()
        cRetVal.Add "xnumFillToTop=" & oBin.ReadDouble()
        cRetVal.Add "xnumFillToBottom=" & oBin.ReadDouble()
        lIdx = oBin.ReadDWord()
        cRetVal.Add "cNumStop=" & lIdx
        For lIdx = 0 To lIdx - 1
            pvDumpColor "xfillGradientStop[" & lIdx & "].brtColor.", oBin, cRetVal
            cRetVal.Add "xfillGradientStop[" & lIdx & "].xnumPosition=" & oBin.ReadDouble()
        Next
    Case ucsBrtBeginBorders
        cRetVal.Add "cborders=" & oBin.ReadDWord()
    Case ucsBrtBorder
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadByte(), "fBdrDiagDown|fBdrDiagUp")
        pvDumpBorder "blxfTop.", oBin, cRetVal
        pvDumpBorder "blxfBottom.", oBin, cRetVal
        pvDumpBorder "blxfLeft.", oBin, cRetVal
        pvDumpBorder "blxfRight.", oBin, cRetVal
        pvDumpBorder "blxfDiag.", oBin, cRetVal
    Case ucsBrtBeginCellStyleXFs
        cRetVal.Add "cxfs=" & oBin.ReadDWord()
    Case ucsBrtBeginCellXFs
        cRetVal.Add "cxfs=" & oBin.ReadDWord()
    Case ucsBrtXF
        cRetVal.Add "ixfeParent=" & oBin.ReadWord()
        cRetVal.Add "iFmt=" & oBin.ReadWord()
        cRetVal.Add "iFont=" & oBin.ReadWord()
        cRetVal.Add "iFill=" & oBin.ReadWord()
        cRetVal.Add "ixBorder=" & oBin.ReadWord()
        cRetVal.Add "trot=" & oBin.ReadByte()
        cRetVal.Add "indent=" & oBin.ReadByte()
        lFlags = oBin.Read3Bytes()
        cRetVal.Add "alc=" & (lFlags And 7)
        cRetVal.Add "alcv=" & (lFlags \ 8 And 7)
        cRetVal.Add "iReadingOrder=" & (lFlags \ &H200 And 3)
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "||||||fWrap|fJustLast|fShrinkToFit|fMergeCell|||fLocked|fHidden|fSxButton|f123Prefix")
        cRetVal.Add "xfGrbitAtr=" & (oBin.ReadByte() And &H3F)
    Case ucsBrtBeginStyles
        cRetVal.Add "cstyles=" & oBin.ReadDWord()
    Case ucsBrtStyle
        cRetVal.Add "ixf=" & oBin.ReadDWord()
        cRetVal.Add "grbitObj1=" & oBin.ReadWord()
        cRetVal.Add "iStyBuiltIn=" & oBin.ReadByte()
        cRetVal.Add "iLevel=" & oBin.ReadByte()
        cRetVal.Add "stName=" & oBin.ReadString()
    Case ucsBrtFileVersion
        cRetVal.Add "guidCodeName=" & oBin.ReadGuid()
        cRetVal.Add "stAppName=" & oBin.ReadString()
        cRetVal.Add "stLastEdited=" & oBin.ReadString()
        cRetVal.Add "stLowestEdited=" & oBin.ReadString()
        cRetVal.Add "stRupBuild=" & oBin.ReadString()
    Case ucsBrtWbProp
        lFlags = oBin.ReadDWord()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "f1904||fHideBorderUnselLists|fFilterPrivacy|fBuggedUserAboutSolution|fShowInkAnnotation|fBackup|fNoSaveSup|||fHidePivotTableFList|fPublishedBookItems|fCheckCompat||fShowPivotChartFilter|fAutoCompressPictures||fRefreshAll")
        cRetVal.Add "grbitUpdateLinks=" & (lFlags \ &H80 And 3)
        cRetVal.Add "mdDspObj=" & (lFlags \ &H400 And 3)
        cRetVal.Add "dwThemeVersion=0x" & Hex$(oBin.ReadDWord())
        cRetVal.Add "strName=" & oBin.ReadString()
    Case ucsBrtBookView
        cRetVal.Add "xWn=" & oBin.ReadDWord()
        cRetVal.Add "yWn=" & oBin.ReadDWord()
        cRetVal.Add "dxWn=" & oBin.ReadDWord()
        cRetVal.Add "dyWn=" & oBin.ReadDWord()
        cRetVal.Add "iTabRatio=" & oBin.ReadDWord()
        cRetVal.Add "itabFirst=" & oBin.ReadDWord()
        cRetVal.Add "itabCur=" & oBin.ReadDWord()
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadByte(), "fHidden|fVeryHidden|fIconic|fDspHScroll|fDspVScroll|fBotAdornment|fAFDateGroup")
    Case ucsBrtBundleSh
        cRetVal.Add "hsState=" & oBin.ReadDWord()
        cRetVal.Add "iTabID=" & oBin.ReadDWord()
        cRetVal.Add "strRelID=" & oBin.ReadString()
        cRetVal.Add "strName=" & oBin.ReadString()
    Case ucsBrtCalcProp
        cRetVal.Add "recalcID=" & oBin.ReadDWord()
        cRetVal.Add "fAutoRecalc=" & oBin.ReadDWord()
        cRetVal.Add "cCalcCount=" & oBin.ReadDWord()
        cRetVal.Add "xnumDelta=" & oBin.ReadDouble()
        cRetVal.Add "cUserThreadCount=" & oBin.ReadDWord()
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fFullCalcOnLoad|fRefA1|fIter|fFullPrec|fSomeUncalced|fSaveRecalc|fMTREnabled|fUserSetThreadCount|fNoDeps")
    Case ucsBrtOleSize
        pvDumpUncheckedRfx "rfx.", oBin, cRetVal
    Case ucsBrtWsProp
        lFlags = oBin.Read3Bytes()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "fShowAutoBreaks|||fPublish|fDialog|fApplyStyles|fRowSumsBelow|fColSumsRight|fFitToPage||fShowOutlineSymbols||fSyncHoriz|fSyncVert|fAltExprEval|fAltFormulaEntry|fFilterMode|fCondFmtCalc")
        pvDumpColor "brtcolorTab.", oBin, cRetVal
        cRetVal.Add "rwSync=" & oBin.ReadDWord()
        cRetVal.Add "colSync=" & oBin.ReadDWord()
        cRetVal.Add "strName=" & oBin.ReadString()
    Case ucsBrtWsDim
        pvDumpUncheckedRfx "rfx.", oBin, cRetVal
    Case ucsBrtACBegin
        lIdx = oBin.ReadWord()
        cRetVal.Add "cver=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add "RgACVer[" & lIdx & "].fileVersion=" & oBin.ReadWord()
            lFlags = oBin.ReadWord()
            cRetVal.Add "RgACVer[" & lIdx & "].fileProduct=" & (lFlags And &H7FFF)
            cRetVal.Add "RgACVer[" & lIdx & "].flags=" & pvFormatFlags(lFlags \ &H8000, "fileExtension")
        Next
    Case ucsBrtRwDescent
        cRetVal.Add "dyDescent=" & oBin.ReadWord()
    Case ucsBrtRowHdr
        cRetVal.Add "rw=" & oBin.ReadDWord()
        cRetVal.Add "ixfe=" & oBin.ReadDWord()
        cRetVal.Add "miyRw=" & oBin.ReadWord()
        lFlags = oBin.Read3Bytes()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "fExtraAsc|fExtraDsc|||||||||fCollapsed|fDyZero|fUnsynced|fGhostDirty||fPhShow")
        cRetVal.Add "iOutLevel=" & (lFlags \ &H100 And 7)
        lIdx = oBin.ReadDWord()
        cRetVal.Add "ccolspan=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add "rgBrtColspan[" & lIdx & "].colMic=" & oBin.ReadDWord()
            cRetVal.Add "rgBrtColspan[" & lIdx & "].colLast=" & oBin.ReadDWord()
        Next
    Case ucsBrtCellRString
        pvDumpCell "cell.", oBin, cRetVal
        pvDumpRichString "value.", oBin, cRetVal
    Case ucsBrtCellBlank
        pvDumpCell "cell.", oBin, cRetVal
    Case ucsBrtCellRk
        pvDumpCell "cell.", oBin, cRetVal
        lFlags = oBin.ReadDWord()
        cRetVal.Add "value.flags=" & pvFormatFlags(lFlags, "fx100|fInt")
        cRetVal.Add "value.num=" & oBin.FromRkNumber(lFlags)
    Case ucsBrtCellReal
        pvDumpCell "cell.", oBin, cRetVal
        cRetVal.Add "xnum=" & oBin.ReadDouble()
    Case ucsBrtCellIsst
        pvDumpCell "cell.", oBin, cRetVal
        cRetVal.Add "isst=" & oBin.ReadDWord()
    Case ucsBrtFmlaNum
        pvDumpCell "cell.", oBin, cRetVal
        cRetVal.Add "xnum=" & oBin.ReadDouble()
        cRetVal.Add "grbitFlags=" & pvFormatFlags(oBin.ReadWord(), "|fAlwaysCalc")
        lIdx = oBin.ReadDWord()
        cRetVal.Add "formula.cce=" & lIdx
        cRetVal.Add "formula.rgce=" & pvFormatBlob(oBin.ReadBlob(lIdx))
        lIdx = oBin.ReadDWord()
        cRetVal.Add "formula.cb=" & lIdx
        cRetVal.Add "formula.rgcb=" & pvFormatBlob(oBin.ReadBlob(lIdx))
    Case ucsBrtBeginTableStyles
        cRetVal.Add "cts=" & oBin.ReadDWord()
        cRetVal.Add "strDefList=" & oBin.ReadString()
        cRetVal.Add "strDefPivot=" & oBin.ReadString()
    Case ucsBrtBeginTableStyle
        lFlags = oBin.ReadWord()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "|fIsPivot|fIsTable")
        cRetVal.Add "ctse=" & oBin.ReadDWord()
        cRetVal.Add "strName=" & oBin.ReadString()
    Case ucsBrtTableStyleElement
        cRetVal.Add "tseType=" & oBin.ReadDWord()
        cRetVal.Add "size=" & oBin.ReadDWord()
        cRetVal.Add "dxfId=" & oBin.ReadDWord()
    Case ucsBrtFRTBegin
        cRetVal.Add "productVersion=" & oBin.ReadDWord()
    Case ucsBrtBeginDXFs
        cRetVal.Add "cdxfs=" & oBin.ReadDWord()
    Case ucsBrtDXF
        lFlags = oBin.ReadWord()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags \ &H8000, "fNewBorder")
        oBin.ReadWord
        lIdx = oBin.ReadWord()
        cRetVal.Add "cprops=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add "xfPropArray[" & lIdx & "].xfPropType=", oBin.ReadWord()
            lFlags = oBin.ReadWord()
            cRetVal.Add "xfPropArray[" & lIdx & "].cb=", lFlags
            cRetVal.Add "xfPropArray[" & lIdx & "].xfPropDataBlob=", pvFormatBlob(oBin.ReadBlob(lFlags))
        Next
    Case ucsBrtAbsPath15
        cRetVal.Add "stAbsPath=" & oBin.ReadString()
    Case ucsBrtFileRecover
        cRetVal.Add "falgs=" & pvFormatFlags(oBin.ReadByte(), "fDontAutoRecover|fSavedDuringRecovery|fCreatedViaMinimalSave|fOpenedViaDataRecovery|fOpenedViaSafeLoad")
    Case ucsBrtBeginWsView
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fWnProt|fDspFmla|fDspGrid|fDspRwCol|fDspZeros|fRightToLeft|fSelected|fDspRuler|fDspGuts|fDefaultHdr|fWhitespaceHidden")
        cRetVal.Add "xlView=" & oBin.ReadDWord()
        cRetVal.Add "rwTop=" & oBin.ReadDWord()
        cRetVal.Add "colLeft=" & oBin.ReadDWord()
        cRetVal.Add "icvHdr=" & oBin.ReadByte()
        oBin.ReadByte
        oBin.ReadWord
        cRetVal.Add "wScaleNormal=" & oBin.ReadWord()
        cRetVal.Add "wScaleSLV=" & oBin.ReadWord()
        cRetVal.Add "wScalePLV=" & oBin.ReadWord()
        cRetVal.Add "iWbkView=" & oBin.ReadDWord()
    Case ucsBrtSel
        cRetVal.Add "pnn=" & oBin.ReadDWord()
        cRetVal.Add "rwAct=" & oBin.ReadDWord()
        cRetVal.Add "colAct=" & oBin.ReadDWord()
        cRetVal.Add "dwRfxAct=" & oBin.ReadDWord()
        lIdx = oBin.ReadDWord()
        cRetVal.Add "sqrfx.crfx=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add "sqrfx.rgrfx[" & lIdx & "]=" & oBin.ReadDWord()
        Next
    Case ucsBrtWsFmtInfo
        cRetVal.Add "dxGCol=" & oBin.ReadDWord()
        cRetVal.Add "cchDefColWidth=" & oBin.ReadWord()
        cRetVal.Add "miyDefRwHeight=" & oBin.ReadWord()
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fUnsynced|fDyZero|fExAsc|fExDesc")
        cRetVal.Add "iOutLevelRw=" & oBin.ReadByte()
        cRetVal.Add "iOutLevelCol=" & oBin.ReadByte()
    Case ucsBrtWsFmtInfoEx14
        cRetVal.Add "dyDescent=" & oBin.ReadWord()
    Case ucsBrtColInfo
        cRetVal.Add "colFirst=" & oBin.ReadDWord()
        cRetVal.Add "colLast=" & oBin.ReadDWord()
        cRetVal.Add "coldx=" & oBin.ReadDWord()
        cRetVal.Add "ixfe=" & oBin.ReadDWord()
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fHidden|fUserSet|fBestFit|fPhonetic|||||iOutLevel||fCollapsed")
    Case ucsBrtBeginMergeCells
        cRetVal.Add "cmcs=" & oBin.ReadDWord()
    Case ucsBrtMergeCell
        cRetVal.Add "rwFirst=" & oBin.ReadDWord()
        cRetVal.Add "rwLast=" & oBin.ReadDWord()
        cRetVal.Add "colFirst=" & oBin.ReadDWord()
        cRetVal.Add "colLast=" & oBin.ReadDWord()
    Case ucsBrtPrintOptions
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fHCenter|fVCenter|fPrintHeaders|fPrintGrid")
    Case ucsBrtMargins
        cRetVal.Add "xnumLeft=" & oBin.ReadDouble()
        cRetVal.Add "xnumRight=" & oBin.ReadDouble()
        cRetVal.Add "xnumTop=" & oBin.ReadDouble()
        cRetVal.Add "xnumBottom=" & oBin.ReadDouble()
        cRetVal.Add "xnumHeader=" & oBin.ReadDouble()
        cRetVal.Add "xnumFooter=" & oBin.ReadDouble()
    Case ucsBrtPageSetup
        cRetVal.Add "iPaperSize=" & oBin.ReadDWord()
        cRetVal.Add "iScale=" & oBin.ReadDWord()
        cRetVal.Add "iRes=" & oBin.ReadDWord()
        cRetVal.Add "iVRes=" & oBin.ReadDWord()
        cRetVal.Add "iCopies=" & oBin.ReadDWord()
        cRetVal.Add "iPageStart=" & oBin.ReadDWord()
        cRetVal.Add "iFitWidth=" & oBin.ReadDWord()
        cRetVal.Add "iFitHeight=" & oBin.ReadDWord()
        lFlags = oBin.ReadWord()
        cRetVal.Add "flags=" & pvFormatFlags(lFlags, "fLeftToRight|fLandscape||fNoColor|fDraft|fNotes|fNoOrient|fUsePage|fEndNotes|")
        cRetVal.Add "iErrors=" & (lFlags \ &H800 And 3)
        cRetVal.Add "szRelID=" & oBin.ReadString()
    Case ucsBrtBeginHeaderFooter
        cRetVal.Add "flags=" & pvFormatFlags(oBin.ReadWord(), "fHFDiffOddEven|fHFDiffFirst|fHFScaleWithDoc|fHFAlignMargins")
        cRetVal.Add "stHeader=" & oBin.ReadString()
        cRetVal.Add "stFooter=" & oBin.ReadString()
        cRetVal.Add "stHeaderEven=" & oBin.ReadString()
        cRetVal.Add "stFooterEven=" & oBin.ReadString()
        cRetVal.Add "stHeaderFirst=" & oBin.ReadString()
        cRetVal.Add "stFooterFirst=" & oBin.ReadString()
    Case ucsBrtDrawing
        cRetVal.Add "stRelId=" & oBin.ReadString()
    Case ucsBrtSheetProtection
        cRetVal.Add "protpwd=" & oBin.ReadWord()
        cRetVal.Add "fLocked=" & oBin.ReadDWord()
        cRetVal.Add "fObjects=" & oBin.ReadDWord()
        cRetVal.Add "fScenarios=" & oBin.ReadDWord()
        cRetVal.Add "fFormatCells=" & oBin.ReadDWord()
        cRetVal.Add "fFormatColumns=" & oBin.ReadDWord()
        cRetVal.Add "fFormatRows=" & oBin.ReadDWord()
        cRetVal.Add "fInsertColumns=" & oBin.ReadDWord()
        cRetVal.Add "fInsertRows=" & oBin.ReadDWord()
        cRetVal.Add "fInsertHyperlinks=" & oBin.ReadDWord()
        cRetVal.Add "fDeleteColumns=" & oBin.ReadDWord()
        cRetVal.Add "fDeleteRows=" & oBin.ReadDWord()
        cRetVal.Add "fSelLockedCells=" & oBin.ReadDWord()
        cRetVal.Add "fSort=" & oBin.ReadDWord()
        cRetVal.Add "fAutoFilter=" & oBin.ReadDWord()
        cRetVal.Add "fPivotTables=" & oBin.ReadDWord()
        cRetVal.Add "fSelUnlockedCells=" & oBin.ReadDWord()
    Case ucsBrtBeginSlicerStyles
        oBin.ReadDWord
        cRetVal.Add "stDefSlicer=" & oBin.ReadString()
    Case ucsBrtBeginTimelineStyles
        oBin.ReadDWord
        cRetVal.Add "stDefTimelineStyle=" & oBin.ReadString()
    Case ucsBrtIndexBlock
        cRetVal.Add "rwMic=" & oBin.ReadDWord()
        cRetVal.Add "rwMac=" & oBin.ReadDWord()
    Case ucsBrtIndexRowBlock
        lIdx = oBin.ReadDWord()
        cRetVal.Add "grbitRowMask=" & pvHexPad(lIdx, 8)
        cRetVal.Add "ibBaseOffset=" & oBin.ReadQWord()
        For lIdx = 0 To PopCount(lIdx) - 1
            lFlags = oBin.ReadWord()
            sText = sText & " " & pvHexPad(lFlags, 4)
            lCount = lCount + PopCount(lFlags)
        Next
        cRetVal.Add "arrayColbitMask=" & sText
        For lIdx = 0 To lCount - 1
            cRetVal.Add "arraySubBaseOffset[" & lIdx & "]=" & pvHexPad(oBin.ReadDWord(), 8)
        Next
    Case ucsBrtMRUColor
        For lIdx = 0 To lRecSize \ 8 - 1
            pvDumpColor "colorMRU[" & lIdx & "].", oBin, cRetVal
        Next
    Case Else
        If lRecSize > 0 Then
            GetBrtData = DesignDumpMemory(lPtr, lRecSize, AddrPadding:=4)
        End If
        Exit Function
    End Select
    If lRecSize > 0 Then
        cRetVal.Add DesignDumpMemory(lPtr, lRecSize, AddrPadding:=4)
    End If
    GetBrtData = ToArray(cRetVal)
End Function

Private Sub pvDumpRichString(sPrefix As String, oBin As cBiff12Part, cRetVal As Collection)
    Dim lFlags          As Long
    Dim lIdx            As Long

    lFlags = oBin.ReadByte()
    cRetVal.Add sPrefix & "flags=" & pvFormatFlags(lFlags, "fRichStr|fExtStr")
    cRetVal.Add sPrefix & "str=" & oBin.ReadString()
    If (lFlags And 1) <> 0 Then
        lIdx = oBin.ReadDWord()
        cRetVal.Add sPrefix & "dwSizeStrRun=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add sPrefix & "rgsStrRun[" & lIdx & "].ich=" & oBin.ReadWord()
            cRetVal.Add sPrefix & "rgsStrRun[" & lIdx & "].ifnt=" & oBin.ReadWord()
        Next
    End If
    If (lFlags And 2) <> 0 Then
        cRetVal.Add sPrefix & "phoneticStr=" & oBin.ReadString()
        lIdx = oBin.ReadDWord()
        cRetVal.Add sPrefix & "dwPhoneticRun=" & lIdx
        For lIdx = 0 To lIdx - 1
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].ichFirst=" & oBin.ReadWord()
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].ichMom=" & oBin.ReadWord()
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].cchMom=" & oBin.ReadWord()
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].ifnt=" & oBin.ReadWord()
            lFlags = oBin.ReadDWord()
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].ifnt=" & (lFlags And 3)
            cRetVal.Add sPrefix & "rgsPhRun[" & lIdx & "].phType=" & (lFlags \ 4 And 3)
        Next
    End If
End Sub

Private Sub pvDumpCell(sPrefix As String, oBin As cBiff12Part, cRetVal As Collection)
    Dim lFlags          As Long

    cRetVal.Add sPrefix & "column=" & oBin.ReadDWord()
    lFlags = oBin.ReadDWord()
    cRetVal.Add sPrefix & "iStyleRef=" & (lFlags And &HFFFFFF)
    cRetVal.Add sPrefix & "flags=" & pvFormatFlags(lFlags \ &H1000000, "fPhShow")
End Sub

Private Sub pvDumpUncheckedRfx(sPrefix As String, oBin As cBiff12Part, cRetVal As Collection)
    cRetVal.Add sPrefix & "rwFirst=" & oBin.ReadDWord()
    cRetVal.Add sPrefix & "rwLast=" & oBin.ReadDWord()
    cRetVal.Add sPrefix & "colFirst=" & oBin.ReadDWord()
    cRetVal.Add sPrefix & "colLast=" & oBin.ReadDWord()
End Sub

Private Sub pvDumpBorder(sPrefix As String, oBin As cBiff12Part, cRetVal As Collection)
    cRetVal.Add sPrefix & "dg=" & oBin.ReadByte()
    oBin.ReadByte
    pvDumpColor sPrefix & "brtColor.", oBin, cRetVal
End Sub

Private Sub pvDumpColor(sPrefix As String, oBin As cBiff12Part, cRetVal As Collection)
    Dim lFlags          As Long

    lFlags = oBin.ReadByte()
    cRetVal.Add sPrefix & "flags=" & pvFormatFlags(lFlags, "fValidRGB")
    cRetVal.Add sPrefix & "xColorType=" & (lFlags \ 2)
    cRetVal.Add sPrefix & "index=" & oBin.ReadByte()
    cRetVal.Add sPrefix & "nTintAndShade=" & oBin.ReadWord()
    cRetVal.Add sPrefix & "bRed=" & oBin.ReadByte()
    cRetVal.Add sPrefix & "bGreen=" & oBin.ReadByte()
    cRetVal.Add sPrefix & "bBlue=" & oBin.ReadByte()
    cRetVal.Add sPrefix & "bAlpha=" & oBin.ReadByte()
End Sub

Private Function pvFormatFlags(ByVal lFlags As Long, sFlags As String) As String
    Dim vSplit          As Variant
    Dim lIdx            As Long

    vSplit = Split(sFlags, "|")
    For lIdx = 0 To UBound(vSplit)
        If LenB(vSplit(lIdx)) <> 0 Then
            If (lFlags And CLng(2 ^ lIdx)) <> 0 Then
                pvFormatFlags = IIf(LenB(pvFormatFlags) <> 0, pvFormatFlags & ", ", vbNullString) & vSplit(lIdx)
            End If
        End If
    Next
    pvFormatFlags = "[" & pvFormatFlags & "]"
End Function

Private Function pvFormatBlob(baBlob() As Byte) As String
    If UBound(baBlob) >= 0 Then
        pvFormatBlob = Replace(DesignDumpMemory(VarPtr(baBlob(0)), UBound(baBlob) + 1, 0), vbCrLf, " ")
    End If
    pvFormatBlob = "{" & pvFormatBlob & "}"
End Function

' HAKMEM algorithm
'     i = i - ((i >> 1) & 0x55555555);
'     i = (i & 0x33333333) + ((i >> 2) & 0x33333333);
'     return (((i + (i >> 4)) & 0x0F0F0F0F) * 0x01010101) >> 24;
Private Function PopCount(ByVal lValue As Long) As Long
    Dim llTemp          As Variant

    llTemp = ToLngLng(lValue, 0)
    llTemp = llTemp - (llTemp \ 2 And &H55555555)
    llTemp = (llTemp And &H33333333) + (llTemp \ 4 And &H33333333)
    PopCount = GetLoDWord((llTemp + llTemp \ 16 And &HF0F0F0F) * &H1010101) \ &H1000000
End Function

Private Function pvHexPad(ByVal lValue As Long, ByVal lSize As Long) As String
    pvHexPad = Right$(String$(lSize, "0") & Hex$(lValue), lSize)
End Function


