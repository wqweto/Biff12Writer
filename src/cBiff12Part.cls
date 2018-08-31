VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12Part"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private Const STR_MODULE_NAME As String = "cBiff12Part"

#Const ImplUseShared = BIFF12_USESHARED
#Const ImplPublicClasses = BIFF12_PUBLICCLASSES

'=========================================================================
' API
'=========================================================================

'--- for SetFilePointer
Private Const FILE_BEGIN                    As Long = 0
Private Const FILE_CURRENT                  As Long = 1
Private Const FILE_END                      As Long = 2
'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, pclsid As Any) As Long
#If ImplUseShared = 0 Then
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
#End If ' ImplUseShared

'=========================================================================
' Public enums and types
'=========================================================================

Public Enum UcsBiff12RecortTypeEnum
    ucsBrtRowHdr = &H0
    ucsBrtCellBlank = &H1
    ucsBrtCellRk = &H2
    ucsBrtCellError = &H3
    ucsBrtCellBool = &H4
    ucsBrtCellReal = &H5
    ucsBrtCellSt = &H6
    ucsBrtCellIsst = &H7
    ucsBrtFmlaString = &H8
    ucsBrtFmlaNum = &H9
    ucsBrtFmlaBool = &HA
    ucsBrtFmlaError = &HB
    ucsBrtFRTArchID = &H10
    ucsBrtSSTItem = &H13
    ucsBrtPCDIMissing = &H14
    ucsBrtPCDINumber = &H15
    ucsBrtPCDIBoolean = &H16
    ucsBrtPCDIError = &H17
    ucsBrtPCDIString = &H18
    ucsBrtPCDIDatetime = &H19
    ucsBrtPCDIIndex = &H1A
    ucsBrtPCDIAMissing = &H1B
    ucsBrtPCDIANumber = &H1C
    ucsBrtPCDIABoolean = &H1D
    ucsBrtPCDIAError = &H1E
    ucsBrtPCDIAString = &H1F
    ucsBrtPCDIADatetime = &H20
    ucsBrtPCRRecord = &H21
    ucsBrtPCRRecordDt = &H22
    ucsBrtFRTBegin = &H23
    ucsBrtFRTEnd = &H24
    ucsBrtACBegin = &H25
    ucsBrtACEnd = &H26
    ucsBrtName = &H27
    ucsBrtIndexRowBlock = &H28
    ucsBrtIndexBlock = &H2A
    ucsBrtFont = &H2B
    ucsBrtFmt = &H2C
    ucsBrtFill = &H2D
    ucsBrtBorder = &H2E
    ucsBrtXF = &H2F
    ucsBrtStyle = &H30
    ucsBrtCellMeta = &H31
    ucsBrtValueMeta = &H32
    ucsBrtMdb = &H33
    ucsBrtBeginFmd = &H34
    ucsBrtEndFmd = &H35
    ucsBrtBeginMdx = &H36
    ucsBrtEndMdx = &H37
    ucsBrtBeginMdxTuple = &H38
    ucsBrtEndMdxTuple = &H39
    ucsBrtMdxMbrIstr = &H3A
    ucsBrtStr = &H3B
    ucsBrtColInfo = &H3C
    ucsBrtCellRString = &H3E
    ucsBrtCalcChainItem = &H3F
    ucsBrtDVal = &H40
    ucsBrtSxvcellNum = &H41
    ucsBrtSxvcellStr = &H42
    ucsBrtSxvcellBool = &H43
    ucsBrtSxvcellErr = &H44
    ucsBrtSxvcellDate = &H45
    ucsBrtSxvcellNil = &H46
    ucsBrtFileVersion = &H80
    ucsBrtBeginSheet = &H81
    ucsBrtEndSheet = &H82
    ucsBrtBeginBook = &H83
    ucsBrtEndBook = &H84
    ucsBrtBeginWsViews = &H85
    ucsBrtEndWsViews = &H86
    ucsBrtBeginBookViews = &H87
    ucsBrtEndBookViews = &H88
    ucsBrtBeginWsView = &H89
    ucsBrtEndWsView = &H8A
    ucsBrtBeginCsViews = &H8B
    ucsBrtEndCsViews = &H8C
    ucsBrtBeginCsView = &H8D
    ucsBrtEndCsView = &H8E
    ucsBrtBeginBundleShs = &H8F
    ucsBrtEndBundleShs = &H90
    ucsBrtBeginSheetData = &H91
    ucsBrtEndSheetData = &H92
    ucsBrtWsProp = &H93
    ucsBrtWsDim = &H94
    ucsBrtPane = &H97
    ucsBrtSel = &H98
    ucsBrtWbProp = &H99
    ucsBrtWbFactoid = &H9A
    ucsBrtFileRecover = &H9B
    ucsBrtBundleSh = &H9C
    ucsBrtCalcProp = &H9D
    ucsBrtBookView = &H9E
    ucsBrtBeginSst = &H9F
    ucsBrtEndSst = &HA0
    ucsBrtBeginAFilter = &HA1
    ucsBrtEndAFilter = &HA2
    ucsBrtBeginFilterColumn = &HA3
    ucsBrtEndFilterColumn = &HA4
    ucsBrtBeginFilters = &HA5
    ucsBrtEndFilters = &HA6
    ucsBrtFilter = &HA7
    ucsBrtColorFilter = &HA8
    ucsBrtIconFilter = &HA9
    ucsBrtTop10Filter = &HAA
    ucsBrtDynamicFilter = &HAB
    ucsBrtBeginCustomFilters = &HAC
    ucsBrtEndCustomFilters = &HAD
    ucsBrtCustomFilter = &HAE
    ucsBrtAFilterDateGroupItem = &HAF
    ucsBrtMergeCell = &HB0
    ucsBrtBeginMergeCells = &HB1
    ucsBrtEndMergeCells = &HB2
    ucsBrtBeginPivotCacheDef = &HB3
    ucsBrtEndPivotCacheDef = &HB4
    ucsBrtBeginPCDFields = &HB5
    ucsBrtEndPCDFields = &HB6
    ucsBrtBeginPCDField = &HB7
    ucsBrtEndPCDField = &HB8
    ucsBrtBeginPCDSource = &HB9
    ucsBrtEndPCDSource = &HBA
    ucsBrtBeginPCDSRange = &HBB
    ucsBrtEndPCDSRange = &HBC
    ucsBrtBeginPCDFAtbl = &HBD
    ucsBrtEndPCDFAtbl = &HBE
    ucsBrtBeginPCDIRun = &HBF
    ucsBrtEndPCDIRun = &HC0
    ucsBrtBeginPivotCacheRecords = &HC1
    ucsBrtEndPivotCacheRecords = &HC2
    ucsBrtBeginPCDHierarchies = &HC3
    ucsBrtEndPCDHierarchies = &HC4
    ucsBrtBeginPCDHierarchy = &HC5
    ucsBrtEndPCDHierarchy = &HC6
    ucsBrtBeginPCDHFieldsUsage = &HC7
    ucsBrtEndPCDHFieldsUsage = &HC8
    ucsBrtBeginExtConnection = &HC9
    ucsBrtEndExtConnection = &HCA
    ucsBrtBeginECDbProps = &HCB
    ucsBrtEndECDbProps = &HCC
    ucsBrtBeginECOlapProps = &HCD
    ucsBrtEndECOlapProps = &HCE
    ucsBrtBeginPCDSConsol = &HCF
    ucsBrtEndPCDSConsol = &HD0
    ucsBrtBeginPCDSCPages = &HD1
    ucsBrtEndPCDSCPages = &HD2
    ucsBrtBeginPCDSCPage = &HD3
    ucsBrtEndPCDSCPage = &HD4
    ucsBrtBeginPCDSCPItem = &HD5
    ucsBrtEndPCDSCPItem = &HD6
    ucsBrtBeginPCDSCSets = &HD7
    ucsBrtEndPCDSCSets = &HD8
    ucsBrtBeginPCDSCSet = &HD9
    ucsBrtEndPCDSCSet = &HDA
    ucsBrtBeginPCDFGroup = &HDB
    ucsBrtEndPCDFGroup = &HDC
    ucsBrtBeginPCDFGItems = &HDD
    ucsBrtEndPCDFGItems = &HDE
    ucsBrtBeginPCDFGRange = &HDF
    ucsBrtEndPCDFGRange = &HE0
    ucsBrtBeginPCDFGDiscrete = &HE1
    ucsBrtEndPCDFGDiscrete = &HE2
    ucsBrtBeginPCDSDTupleCache = &HE3
    ucsBrtEndPCDSDTupleCache = &HE4
    ucsBrtBeginPCDSDTCEntries = &HE5
    ucsBrtEndPCDSDTCEntries = &HE6
    ucsBrtBeginPCDSDTCEMembers = &HE7
    ucsBrtEndPCDSDTCEMembers = &HE8
    ucsBrtBeginPCDSDTCEMember = &HE9
    ucsBrtEndPCDSDTCEMember = &HEA
    ucsBrtBeginPCDSDTCQueries = &HEB
    ucsBrtEndPCDSDTCQueries = &HEC
    ucsBrtBeginPCDSDTCQuery = &HED
    ucsBrtEndPCDSDTCQuery = &HEE
    ucsBrtBeginPCDSDTCSets = &HEF
    ucsBrtEndPCDSDTCSets = &HF0
    ucsBrtBeginPCDSDTCSet = &HF1
    ucsBrtEndPCDSDTCSet = &HF2
    ucsBrtBeginPCDCalcItems = &HF3
    ucsBrtEndPCDCalcItems = &HF4
    ucsBrtBeginPCDCalcItem = &HF5
    ucsBrtEndPCDCalcItem = &HF6
    ucsBrtBeginPRule = &HF7
    ucsBrtEndPRule = &HF8
    ucsBrtBeginPRFilters = &HF9
    ucsBrtEndPRFilters = &HFA
    ucsBrtBeginPRFilter = &HFB
    ucsBrtEndPRFilter = &HFC
    ucsBrtBeginPNames = &HFD
    ucsBrtEndPNames = &HFE
    ucsBrtBeginPName = &HFF
    ucsBrtEndPName = &H100
    ucsBrtBeginPNPairs = &H101
    ucsBrtEndPNPairs = &H102
    ucsBrtBeginPNPair = &H103
    ucsBrtEndPNPair = &H104
    ucsBrtBeginECWebProps = &H105
    ucsBrtEndECWebProps = &H106
    ucsBrtBeginEcWpTables = &H107
    ucsBrtEndECWPTables = &H108
    ucsBrtBeginECParams = &H109
    ucsBrtEndECParams = &H10A
    ucsBrtBeginECParam = &H10B
    ucsBrtEndECParam = &H10C
    ucsBrtBeginPCDKPIs = &H10D
    ucsBrtEndPCDKPIs = &H10E
    ucsBrtBeginPCDKPI = &H10F
    ucsBrtEndPCDKPI = &H110
    ucsBrtBeginDims = &H111
    ucsBrtEndDims = &H112
    ucsBrtBeginDim = &H113
    ucsBrtEndDim = &H114
    ucsBrtIndexPartEnd = &H115
    ucsBrtBeginStyleSheet = &H116
    ucsBrtEndStyleSheet = &H117
    ucsBrtBeginSXView = &H118
    ucsBrtEndSXVI = &H119
    ucsBrtBeginSXVI = &H11A
    ucsBrtBeginSXVIs = &H11B
    ucsBrtEndSXVIs = &H11C
    ucsBrtBeginSXVD = &H11D
    ucsBrtEndSXVD = &H11E
    ucsBrtBeginSXVDs = &H11F
    ucsBrtEndSXVDs = &H120
    ucsBrtBeginSXPI = &H121
    ucsBrtEndSXPI = &H122
    ucsBrtBeginSXPIs = &H123
    ucsBrtEndSXPIs = &H124
    ucsBrtBeginSXDI = &H125
    ucsBrtEndSXDI = &H126
    ucsBrtBeginSXDIs = &H127
    ucsBrtEndSXDIs = &H128
    ucsBrtBeginSXLI = &H129
    ucsBrtEndSXLI = &H12A
    ucsBrtBeginSXLIRws = &H12B
    ucsBrtEndSXLIRws = &H12C
    ucsBrtBeginSXLICols = &H12D
    ucsBrtEndSXLICols = &H12E
    ucsBrtBeginSXFormat = &H12F
    ucsBrtEndSXFormat = &H130
    ucsBrtBeginSXFormats = &H131
    ucsBrtEndSxFormats = &H132
    ucsBrtBeginSxSelect = &H133
    ucsBrtEndSxSelect = &H134
    ucsBrtBeginISXVDRws = &H135
    ucsBrtEndISXVDRws = &H136
    ucsBrtBeginISXVDCols = &H137
    ucsBrtEndISXVDCols = &H138
    ucsBrtEndSXLocation = &H139
    ucsBrtBeginSXLocation = &H13A
    ucsBrtEndSXView = &H13B
    ucsBrtBeginSXTHs = &H13C
    ucsBrtEndSXTHs = &H13D
    ucsBrtBeginSXTH = &H13E
    ucsBrtEndSXTH = &H13F
    ucsBrtBeginISXTHRws = &H140
    ucsBrtEndISXTHRws = &H141
    ucsBrtBeginISXTHCols = &H142
    ucsBrtEndISXTHCols = &H143
    ucsBrtBeginSXTDMPS = &H144
    ucsBrtEndSXTDMPs = &H145
    ucsBrtBeginSXTDMP = &H146
    ucsBrtEndSXTDMP = &H147
    ucsBrtBeginSXTHItems = &H148
    ucsBrtEndSXTHItems = &H149
    ucsBrtBeginSXTHItem = &H14A
    ucsBrtEndSXTHItem = &H14B
    ucsBrtBeginMetadata = &H14C
    ucsBrtEndMetadata = &H14D
    ucsBrtBeginEsmdtinfo = &H14E
    ucsBrtMdtinfo = &H14F
    ucsBrtEndEsmdtinfo = &H150
    ucsBrtBeginEsmdb = &H151
    ucsBrtEndEsmdb = &H152
    ucsBrtBeginEsfmd = &H153
    ucsBrtEndEsfmd = &H154
    ucsBrtBeginSingleCells = &H155
    ucsBrtEndSingleCells = &H156
    ucsBrtBeginList = &H157
    ucsBrtEndList = &H158
    ucsBrtBeginListCols = &H159
    ucsBrtEndListCols = &H15A
    ucsBrtBeginListCol = &H15B
    ucsBrtEndListCol = &H15C
    ucsBrtBeginListXmlCPr = &H15D
    ucsBrtEndListXmlCPr = &H15E
    ucsBrtListCCFmla = &H15F
    ucsBrtListTrFmla = &H160
    ucsBrtBeginExternals = &H161
    ucsBrtEndExternals = &H162
    ucsBrtSupBookSrc = &H163
    ucsBrtSupSelf = &H165
    ucsBrtSupSame = &H166
    ucsBrtSupTabs = &H167
    ucsBrtBeginSupBook = &H168
    ucsBrtPlaceholderName = &H169
    ucsBrtExternSheet = &H16A
    ucsBrtExternTableStart = &H16B
    ucsBrtExternTableEnd = &H16C
    ucsBrtExternRowHdr = &H16E
    ucsBrtExternCellBlank = &H16F
    ucsBrtExternCellReal = &H170
    ucsBrtExternCellBool = &H171
    ucsBrtExternCellError = &H172
    ucsBrtExternCellString = &H173
    ucsBrtBeginEsmdx = &H174
    ucsBrtEndEsmdx = &H175
    ucsBrtBeginMdxSet = &H176
    ucsBrtEndMdxSet = &H177
    ucsBrtBeginMdxMbrProp = &H178
    ucsBrtEndMdxMbrProp = &H179
    ucsBrtBeginMdxKPI = &H17A
    ucsBrtEndMdxKPI = &H17B
    ucsBrtBeginEsstr = &H17C
    ucsBrtEndEsstr = &H17D
    ucsBrtBeginPRFItem = &H17E
    ucsBrtEndPRFItem = &H17F
    ucsBrtBeginPivotCacheIDs = &H180
    ucsBrtEndPivotCacheIDs = &H181
    ucsBrtBeginPivotCacheID = &H182
    ucsBrtEndPivotCacheID = &H183
    ucsBrtBeginISXVIs = &H184
    ucsBrtEndISXVIs = &H185
    ucsBrtBeginColInfos = &H186
    ucsBrtEndColInfos = &H187
    ucsBrtBeginRwBrk = &H188
    ucsBrtEndRwBrk = &H189
    ucsBrtBeginColBrk = &H18A
    ucsBrtEndColBrk = &H18B
    ucsBrtBrk = &H18C
    ucsBrtUserBookView = &H18D
    ucsBrtInfo = &H18E
    ucsBrtCUsr = &H18F
    ucsBrtUsr = &H190
    ucsBrtBeginUsers = &H191
    ucsBrtEOF = &H193
    ucsBrtUCR = &H194
    ucsBrtRRInsDel = &H195
    ucsBrtRREndInsDel = &H196
    ucsBrtRRMove = &H197
    ucsBrtRREndMove = &H198
    ucsBrtRRChgCell = &H199
    ucsBrtRREndChgCell = &H19A
    ucsBrtRRHeader = &H19B
    ucsBrtRRUserView = &H19C
    ucsBrtRRRenSheet = &H19D
    ucsBrtRRInsertSh = &H19E
    ucsBrtRRDefName = &H19F
    ucsBrtRRNote = &H1A0
    ucsBrtRRConflict = &H1A1
    ucsBrtRRTQSIF = &H1A2
    ucsBrtRRFormat = &H1A3
    ucsBrtRREndFormat = &H1A4
    ucsBrtRRAutoFmt = &H1A5
    ucsBrtBeginUserShViews = &H1A6
    ucsBrtBeginUserShView = &H1A7
    ucsBrtEndUserShView = &H1A8
    ucsBrtEndUserShViews = &H1A9
    ucsBrtArrFmla = &H1AA
    ucsBrtShrFmla = &H1AB
    ucsBrtTable = &H1AC
    ucsBrtBeginExtConnections = &H1AD
    ucsBrtEndExtConnections = &H1AE
    ucsBrtBeginPCDCalcMems = &H1AF
    ucsBrtEndPCDCalcMems = &H1B0
    ucsBrtBeginPCDCalcMem = &H1B1
    ucsBrtEndPCDCalcMem = &H1B2
    ucsBrtBeginPCDHGLevels = &H1B3
    ucsBrtEndPCDHGLevels = &H1B4
    ucsBrtBeginPCDHGLevel = &H1B5
    ucsBrtEndPCDHGLevel = &H1B6
    ucsBrtBeginPCDHGLGroups = &H1B7
    ucsBrtEndPCDHGLGroups = &H1B8
    ucsBrtBeginPCDHGLGroup = &H1B9
    ucsBrtEndPCDHGLGroup = &H1BA
    ucsBrtBeginPCDHGLGMembers = &H1BB
    ucsBrtEndPCDHGLGMembers = &H1BC
    ucsBrtBeginPCDHGLGMember = &H1BD
    ucsBrtEndPCDHGLGMember = &H1BE
    ucsBrtBeginQSI = &H1BF
    ucsBrtEndQSI = &H1C0
    ucsBrtBeginQSIR = &H1C1
    ucsBrtEndQSIR = &H1C2
    ucsBrtBeginDeletedNames = &H1C3
    ucsBrtEndDeletedNames = &H1C4
    ucsBrtBeginDeletedName = &H1C5
    ucsBrtEndDeletedName = &H1C6
    ucsBrtBeginQSIFs = &H1C7
    ucsBrtEndQSIFs = &H1C8
    ucsBrtBeginQSIF = &H1C9
    ucsBrtEndQSIF = &H1CA
    ucsBrtBeginAutoSortScope = &H1CB
    ucsBrtEndAutoSortScope = &H1CC
    ucsBrtBeginConditionalFormatting = &H1CD
    ucsBrtEndConditionalFormatting = &H1CE
    ucsBrtBeginCFRule = &H1CF
    ucsBrtEndCFRule = &H1D0
    ucsBrtBeginIconSet = &H1D1
    ucsBrtEndIconSet = &H1D2
    ucsBrtBeginDatabar = &H1D3
    ucsBrtEndDatabar = &H1D4
    ucsBrtBeginColorScale = &H1D5
    ucsBrtEndColorScale = &H1D6
    ucsBrtCFVO = &H1D7
    ucsBrtExternValueMeta = &H1D8
    ucsBrtBeginColorPalette = &H1D9
    ucsBrtEndColorPalette = &H1DA
    ucsBrtIndexedColor = &H1DB
    ucsBrtMargins = &H1DC
    ucsBrtPrintOptions = &H1DD
    ucsBrtPageSetup = &H1DE
    ucsBrtBeginHeaderFooter = &H1DF
    ucsBrtEndHeaderFooter = &H1E0
    ucsBrtBeginSXCrtFormat = &H1E1
    ucsBrtEndSXCrtFormat = &H1E2
    ucsBrtBeginSXCrtFormats = &H1E3
    ucsBrtEndSXCrtFormats = &H1E4
    ucsBrtWsFmtInfo = &H1E5
    ucsBrtBeginMgs = &H1E6
    ucsBrtEndMGs = &H1E7
    ucsBrtBeginMGMaps = &H1E8
    ucsBrtEndMGMaps = &H1E9
    ucsBrtBeginMG = &H1EA
    ucsBrtEndMG = &H1EB
    ucsBrtBeginMap = &H1EC
    ucsBrtEndMap = &H1ED
    ucsBrtHLink = &H1EE
    ucsBrtBeginDCon = &H1EF
    ucsBrtEndDCon = &H1F0
    ucsBrtBeginDRefs = &H1F1
    ucsBrtEndDRefs = &H1F2
    ucsBrtDRef = &H1F3
    ucsBrtBeginScenMan = &H1F4
    ucsBrtEndScenMan = &H1F5
    ucsBrtBeginSct = &H1F6
    ucsBrtEndSct = &H1F7
    ucsBrtSlc = &H1F8
    ucsBrtBeginDXFs = &H1F9
    ucsBrtEndDXFs = &H1FA
    ucsBrtDXF = &H1FB
    ucsBrtBeginTableStyles = &H1FC
    ucsBrtEndTableStyles = &H1FD
    ucsBrtBeginTableStyle = &H1FE
    ucsBrtEndTableStyle = &H1FF
    ucsBrtTableStyleElement = &H200
    ucsBrtTableStyleClient = &H201
    ucsBrtBeginVolDeps = &H202
    ucsBrtEndVolDeps = &H203
    ucsBrtBeginVolType = &H204
    ucsBrtEndVolType = &H205
    ucsBrtBeginVolMain = &H206
    ucsBrtEndVolMain = &H207
    ucsBrtBeginVolTopic = &H208
    ucsBrtEndVolTopic = &H209
    ucsBrtVolSubtopic = &H20A
    ucsBrtVolRef = &H20B
    ucsBrtVolNum = &H20C
    ucsBrtVolErr = &H20D
    ucsBrtVolStr = &H20E
    ucsBrtVolBool = &H20F
    ucsBrtBeginCalcChain = &H210
    ucsBrtEndCalcChain = &H211
    ucsBrtBeginSortState = &H212
    ucsBrtEndSortState = &H213
    ucsBrtBeginSortCond = &H214
    ucsBrtEndSortCond = &H215
    ucsBrtBookProtection = &H216
    ucsBrtSheetProtection = &H217
    ucsBrtRangeProtection = &H218
    ucsBrtPhoneticInfo = &H219
    ucsBrtBeginECTxtWiz = &H21A
    ucsBrtEndECTxtWiz = &H21B
    ucsBrtBeginECTWFldInfoLst = &H21C
    ucsBrtEndECTWFldInfoLst = &H21D
    ucsBrtBeginECTwFldInfo = &H21E
    ucsBrtFileSharing = &H224
    ucsBrtOleSize = &H225
    ucsBrtDrawing = &H226
    ucsBrtLegacyDrawing = &H227
    ucsBrtLegacyDrawingHF = &H228
    ucsBrtWebOpt = &H229
    ucsBrtBeginWebPubItems = &H22A
    ucsBrtEndWebPubItems = &H22B
    ucsBrtBeginWebPubItem = &H22C
    ucsBrtEndWebPubItem = &H22D
    ucsBrtBeginSXCondFmt = &H22E
    ucsBrtEndSXCondFmt = &H22F
    ucsBrtBeginSXCondFmts = &H230
    ucsBrtEndSXCondFmts = &H231
    ucsBrtBkHim = &H232
    ucsbrtcolor = &H234
    ucsBrtBeginIndexedColors = &H235
    ucsBrtEndIndexedColors = &H236
    ucsBrtBeginMRUColors = &H239
    ucsBrtEndMRUColors = &H23A
    ucsBrtMRUColor = &H23C
    ucsBrtBeginDVals = &H23D
    ucsBrtEndDVals = &H23E
    ucsBrtSupNameStart = &H241
    ucsBrtSupNameValueStart = &H242
    ucsBrtSupNameValueEnd = &H243
    ucsBrtSupNameNum = &H244
    ucsBrtSupNameErr = &H245
    ucsBrtSupNameSt = &H246
    ucsBrtSupNameNil = &H247
    ucsBrtSupNameBool = &H248
    ucsBrtSupNameFmla = &H249
    ucsBrtSupNameBits = &H24A
    ucsBrtSupNameEnd = &H24B
    ucsBrtEndSupBook = &H24C
    ucsBrtCellSmartTagProperty = &H24D
    ucsBrtBeginCellSmartTag = &H24E
    ucsBrtEndCellSmartTag = &H24F
    ucsBrtBeginCellSmartTags = &H250
    ucsBrtEndCellSmartTags = &H251
    ucsBrtBeginSmartTags = &H252
    ucsBrtEndSmartTags = &H253
    ucsBrtSmartTagType = &H254
    ucsBrtBeginSmartTagTypes = &H255
    ucsBrtEndSmartTagTypes = &H256
    ucsBrtBeginSXFilters = &H257
    ucsBrtEndSXFilters = &H258
    ucsBrtBeginSXFILTER = &H259
    ucsBrtEndSXFilter = &H25A
    ucsBrtBeginFills = &H25B
    ucsBrtEndFills = &H25C
    ucsBrtBeginCellWatches = &H25D
    ucsBrtEndCellWatches = &H25E
    ucsBrtCellWatch = &H25F
    ucsBrtBeginCRErrs = &H260
    ucsBrtEndCRErrs = &H261
    ucsBrtCrashRecErr = &H262
    ucsBrtBeginFonts = &H263
    ucsBrtEndFonts = &H264
    ucsBrtBeginBorders = &H265
    ucsBrtEndBorders = &H266
    ucsBrtBeginFmts = &H267
    ucsBrtEndFmts = &H268
    ucsBrtBeginCellXFs = &H269
    ucsBrtEndCellXFs = &H26A
    ucsBrtBeginStyles = &H26B
    ucsBrtEndStyles = &H26C
    ucsBrtBigName = &H271
    ucsBrtBeginCellStyleXFs = &H272
    ucsBrtEndCellStyleXFs = &H273
    ucsBrtBeginComments = &H274
    ucsBrtEndComments = &H275
    ucsBrtBeginCommentAuthors = &H276
    ucsBrtEndCommentAuthors = &H277
    ucsBrtCommentAuthor = &H278
    ucsBrtBeginCommentList = &H279
    ucsBrtEndCommentList = &H27A
    ucsBrtBeginComment = &H27B
    ucsBrtEndComment = &H27C
    ucsBrtCommentText = &H27D
    ucsBrtBeginOleObjects = &H27E
    ucsBrtOleObject = &H27F
    ucsBrtEndOleObjects = &H280
    ucsBrtBeginSxrules = &H281
    ucsBrtEndSxRules = &H282
    ucsBrtBeginActiveXControls = &H283
    ucsBrtActiveX = &H284
    ucsBrtEndActiveXControls = &H285
    ucsBrtBeginPCDSDTCEMembersSortBy = &H286
    ucsBrtBeginCellIgnoreECs = &H288
    ucsBrtCellIgnoreEC = &H289
    ucsBrtEndCellIgnoreECs = &H28A
    ucsBrtCsProp = &H28B
    ucsBrtCsPageSetup = &H28C
    ucsBrtBeginUserCsViews = &H28D
    ucsBrtEndUserCsViews = &H28E
    ucsBrtBeginUserCsView = &H28F
    ucsBrtEndUserCsView = &H290
    ucsBrtBeginPcdSFCIEntries = &H291
    ucsBrtEndPCDSFCIEntries = &H292
    ucsBrtPCDSFCIEntry = &H293
    ucsBrtBeginListParts = &H294
    ucsBrtListPart = &H295
    ucsBrtEndListParts = &H296
    ucsBrtSheetCalcProp = &H297
    ucsBrtBeginFnGroup = &H298
    ucsBrtFnGroup = &H299
    ucsBrtEndFnGroup = &H29A
    ucsBrtSupAddin = &H29B
    ucsBrtSXTDMPOrder = &H29C
    ucsBrtCsProtection = &H29D
    ucsBrtBeginWsSortMap = &H29F
    ucsBrtEndWsSortMap = &H2A0
    ucsBrtBeginRRSort = &H2A1
    ucsBrtEndRRSort = &H2A2
    ucsBrtRRSortItem = &H2A3
    ucsBrtFileSharingIso = &H2A4
    ucsBrtBookProtectionIso = &H2A5
    ucsBrtSheetProtectionIso = &H2A6
    ucsBrtCsProtectionIso = &H2A7
    ucsBrtRangeProtectionIso = &H2A8
    ucsBrtRwDescent = &H400
    ucsBrtKnownFonts = &H401
    ucsBrtBeginSXTupleSet = &H402
    ucsBrtEndSXTupleSet = &H403
    ucsBrtBeginSXTupleSetHeader = &H404
    ucsBrtEndSXTupleSetHeader = &H405
    ucsBrtSXTupleSetHeaderItem = &H406
    ucsBrtBeginSXTupleSetData = &H407
    ucsBrtEndSXTupleSetData = &H408
    ucsBrtBeginSXTupleSetRow = &H409
    ucsBrtEndSXTupleSetRow = &H40A
    ucsBrtSXTupleSetRowItem = &H40B
    ucsBrtNameExt = &H40C
    ucsBrtPCDH14 = &H40D
    ucsBrtBeginPCDCalcMem14 = &H40E
    ucsBrtEndPCDCalcMem14 = &H40F
    ucsBrtSXTH14 = &H410
    ucsBrtBeginSparklineGroup = &H411
    ucsBrtEndSparklineGroup = &H412
    ucsBrtSparkline = &H413
    ucsBrtSXDI14 = &H414
    ucsBrtWsFmtInfoEx14 = &H415
    ucsBrtBeginConditionalFormatting14 = &H416
    ucsBrtEndConditionalFormatting14 = &H417
    ucsBrtBeginCFRule14 = &H418
    ucsBrtEndCFRule14 = &H419
    ucsBrtCFVO14 = &H41A
    ucsBrtBeginDatabar14 = &H41B
    ucsBrtBeginIconSet14 = &H41C
    ucsBrtDVal14 = &H41D
    ucsBrtBeginDVals14 = &H41E
    ucsBrtColor14 = &H41F
    ucsBrtBeginSparklines = &H420
    ucsBrtEndSparklines = &H421
    ucsBrtBeginSparklineGroups = &H422
    ucsBrtEndSparklineGroups = &H423
    ucsBrtSXVD14 = &H425
    ucsBrtBeginSxview14 = &H426
    ucsBrtEndSxview14 = &H427
    ucsBrtBeginPCD14 = &H42A
    ucsBrtEndPCD14 = &H42B
    ucsBrtBeginExtConn14 = &H42C
    ucsBrtEndExtConn14 = &H42D
    ucsBrtBeginSlicerCacheIDs = &H42E
    ucsBrtEndSlicerCacheIDs = &H42F
    ucsBrtBeginSlicerCacheID = &H430
    ucsBrtEndSlicerCacheID = &H431
    ucsBrtBeginSlicerCache = &H433
    ucsBrtEndSlicerCache = &H434
    ucsBrtBeginSlicerCacheDef = &H435
    ucsBrtEndSlicerCacheDef = &H436
    ucsBrtBeginSlicersEx = &H437
    ucsBrtEndSlicersEx = &H438
    ucsBrtBeginSlicerEx = &H439
    ucsBrtEndSlicerEx = &H43A
    ucsBrtBeginSlicer = &H43B
    ucsBrtEndSlicer = &H43C
    ucsBrtSlicerCachePivotTables = &H43D
    ucsBrtBeginSlicerCacheOlapImpl = &H43E
    ucsBrtEndSlicerCacheOlapImpl = &H43F
    ucsBrtBeginSlicerCacheLevelsData = &H440
    ucsBrtEndSlicerCacheLevelsData = &H441
    ucsBrtBeginSlicerCacheLevelData = &H442
    ucsBrtEndSlicerCacheLevelData = &H443
    ucsBrtBeginSlicerCacheSiRanges = &H444
    ucsBrtEndSlicerCacheSiRanges = &H445
    ucsBrtBeginSlicerCacheSiRange = &H446
    ucsBrtEndSlicerCacheSiRange = &H447
    ucsBrtSlicerCacheOlapItem = &H448
    ucsBrtBeginSlicerCacheSelections = &H449
    ucsBrtSlicerCacheSelection = &H44A
    ucsBrtEndSlicerCacheSelections = &H44B
    ucsBrtBeginSlicerCacheNative = &H44C
    ucsBrtEndSlicerCacheNative = &H44D
    ucsBrtSlicerCacheNativeItem = &H44E
    ucsBrtRangeProtection14 = &H44F
    ucsBrtRangeProtectionIso14 = &H450
    ucsBrtCellIgnoreEC14 = &H451
    ucsBrtList14 = &H457
    ucsBrtCFIcon = &H458
    ucsBrtBeginSlicerCachesPivotCacheIDs = &H459
    ucsBrtEndSlicerCachesPivotCacheIDs = &H45A
    ucsBrtBeginSlicers = &H45B
    ucsBrtEndSlicers = &H45C
    ucsBrtWbProp14 = &H45D
    ucsBrtBeginSXEdit = &H45E
    ucsBrtEndSXEdit = &H45F
    ucsBrtBeginSXEdits = &H460
    ucsBrtEndSXEdits = &H461
    ucsBrtBeginSXChange = &H462
    ucsBrtEndSXChange = &H463
    ucsBrtBeginSXChanges = &H464
    ucsBrtEndSXChanges = &H465
    ucsBrtSXTupleItems = &H466
    ucsBrtBeginSlicerStyle = &H468
    ucsBrtEndSlicerStyle = &H469
    ucsBrtSlicerStyleElement = &H46A
    ucsBrtBeginStyleSheetExt14 = &H46B
    ucsBrtEndStyleSheetExt14 = &H46C
    ucsBrtBeginSlicerCachesPivotCacheID = &H46D
    ucsBrtEndSlicerCachesPivotCacheID = &H46E
    ucsBrtBeginConditionalFormattings = &H46F
    ucsBrtEndConditionalFormattings = &H470
    ucsBrtBeginPCDCalcMemExt = &H471
    ucsBrtEndPCDCalcMemExt = &H472
    ucsBrtBeginPCDCalcMemsExt = &H473
    ucsBrtEndPCDCalcMemsExt = &H474
    ucsBrtPCDField14 = &H475
    ucsBrtBeginSlicerStyles = &H476
    ucsBrtEndSlicerStyles = &H477
    ucsBrtBeginSlicerStyleElements = &H478
    ucsBrtEndSlicerStyleElements = &H479
    ucsBrtCFRuleExt = &H47A
    ucsBrtBeginSXCondFmt14 = &H47B
    ucsBrtEndSXCondFmt14 = &H47C
    ucsBrtBeginSXCondFmts14 = &H47D
    ucsBrtEndSXCondFmts14 = &H47E
    ucsBrtBeginSortCond14 = &H480
    ucsBrtEndSortCond14 = &H481
    ucsBrtEndDVals14 = &H482
    ucsBrtEndIconSet14 = &H483
    ucsBrtEndDatabar14 = &H484
    ucsBrtBeginColorScale14 = &H485
    ucsBrtEndColorScale14 = &H486
    ucsBrtBeginSxrules14 = &H487
    ucsBrtEndSxrules14 = &H488
    ucsBrtBeginPRule14 = &H489
    ucsBrtEndPRule14 = &H48A
    ucsBrtBeginPRFilters14 = &H48B
    ucsBrtEndPRFilters14 = &H48C
    ucsBrtBeginPRFilter14 = &H48D
    ucsBrtEndPRFilter14 = &H48E
    ucsBrtBeginPRFItem14 = &H48F
    ucsBrtEndPRFItem14 = &H490
    ucsBrtBeginCellIgnoreECs14 = &H491
    ucsBrtEndCellIgnoreECs14 = &H492
    ucsBrtDxf14 = &H493
    ucsBrtBeginDxF14s = &H494
    ucsBrtEndDxf14s = &H495
    ucsBrtFilter14 = &H499
    ucsBrtBeginCustomFilters14 = &H49A
    ucsBrtCustomFilter14 = &H49C
    ucsBrtIconFilter14 = &H49D
    ucsBrtPivotCacheConnectionName = &H49E
    ucsBrtBeginDecoupledPivotCacheIDs = &H800
    ucsBrtEndDecoupledPivotCacheIDs = &H801
    ucsBrtDecoupledPivotCacheID = &H802
    ucsBrtBeginPivotTableRefs = &H803
    ucsBrtEndPivotTableRefs = &H804
    ucsBrtPivotTableRef = &H805
    ucsBrtSlicerCacheBookPivotTables = &H806
    ucsBrtBeginSxvcells = &H807
    ucsBrtEndSxvcells = &H808
    ucsBrtBeginSxRow = &H809
    ucsBrtEndSxRow = &H80A
    ucsBrtPcdCalcMem15 = &H80C
    ucsBrtQsi15 = &H813
    ucsBrtBeginWebExtensions = &H814
    ucsBrtEndWebExtensions = &H815
    ucsBrtWebExtension = &H816
    ucsBrtAbsPath15 = &H817
    ucsBrtBeginPivotTableUISettings = &H818
    ucsBrtEndPivotTableUISettings = &H819
    ucsBrtTableSlicerCacheIDs = &H81B
    ucsBrtTableSlicerCacheID = &H81C
    ucsBrtBeginTableSlicerCache = &H81D
    ucsBrtEndTableSlicerCache = &H81E
    ucsBrtSxFilter15 = &H81F
    ucsBrtBeginTimelineCachePivotCacheIDs = &H820
    ucsBrtEndTimelineCachePivotCacheIDs = &H821
    ucsBrtTimelineCachePivotCacheID = &H822
    ucsBrtBeginTimelineCacheIDs = &H823
    ucsBrtEndTimelineCacheIDs = &H824
    ucsBrtBeginTimelineCacheID = &H825
    ucsBrtEndTimelineCacheID = &H826
    ucsBrtBeginTimelinesEx = &H827
    ucsBrtEndTimelinesEx = &H828
    ucsBrtBeginTimelineEx = &H829
    ucsBrtEndTimelineEx = &H82A
    ucsBrtWorkBookPr15 = &H82B
    ucsBrtPCDH15 = &H82C
    ucsBrtBeginTimelineStyle = &H82D
    ucsBrtEndTimelineStyle = &H82E
    ucsBrtTimelineStyleElement = &H82F
    ucsBrtBeginTimelineStylesheetExt15 = &H830
    ucsBrtEndTimelineStylesheetExt15 = &H831
    ucsBrtBeginTimelineStyles = &H832
    ucsBrtEndTimelineStyles = &H833
    ucsBrtBeginTimelineStyleElements = &H834
    ucsBrtEndTimelineStyleElements = &H835
    ucsBrtDxf15 = &H836
    ucsBrtBeginDxfs15 = &H837
    ucsbrtEndDxfs15 = &H838
    ucsBrtSlicerCacheHideItemsWithNoData = &H839
    ucsBrtBeginItemUniqueNames = &H83A
    ucsBrtEndItemUniqueNames = &H83B
    ucsBrtItemUniqueName = &H83C
    ucsBrtBeginExtConn15 = &H83D
    ucsBrtEndExtConn15 = &H83E
    ucsBrtBeginOledbPr15 = &H83F
    ucsBrtEndOledbPr15 = &H840
    ucsBrtBeginDataFeedPr15 = &H841
    ucsBrtEndDataFeedPr15 = &H842
    ucsBrtTextPr15 = &H843
    ucsBrtRangePr15 = &H844
    ucsBrtDbCommand15 = &H845
    ucsBrtBeginDbTables15 = &H846
    ucsBrtEndDbTables15 = &H847
    ucsBrtDbTable15 = &H848
    ucsBrtBeginDataModel = &H849
    ucsBrtEndDataModel = &H84A
    ucsBrtBeginModelTables = &H84B
    ucsBrtEndModelTables = &H84C
    ucsBrtModelTable = &H84D
    ucsBrtBeginModelRelationships = &H84E
    ucsBrtEndModelRelationships = &H84F
    ucsBrtModelRelationship = &H850
    ucsBrtBeginECTxtWiz15 = &H851
    ucsBrtEndECTxtWiz15 = &H852
    ucsBrtBeginECTWFldInfoLst15 = &H853
    ucsBrtEndECTWFldInfoLst15 = &H854
    ucsBrtBeginECTWFldInfo15 = &H855
    ucsBrtFieldListActiveItem = &H856
    ucsBrtPivotCacheIdVersion = &H857
    ucsBrtSXDI15 = &H858
End Enum

#If ImplPublicClasses Then

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

#End If ' ImplPublicClasses

'=========================================================================
' API
'=========================================================================

Private Const VT_I8                                 As Long = 20

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Function VariantChangeType Lib "oleaut32" (dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_SST_HASH_SIZE     As Long = 2 ^ 13
Private Const LNG_RK_MAXINT         As Long = 2 ^ 29 - 1
Private Const LNG_RK_FX100          As Long = 1
Private Const LNG_RK_FINT           As Long = 2
Private Const LNG_RK_NUMOFFSET      As Long = 4

Private m_baBuffer()            As Byte
Private m_lSize                 As Long
Private m_lPosition             As Long
Private m_uaSstHashEntries()    As UcsSstHashEntryType
Private m_laSstKeyPositions()   As Long
Private m_lSstCount             As Long
Private m_lSstCapacity          As Long
Private m_sName                 As String
Private m_sSheetName            As String
Private m_sType                 As String
Private m_oXmlDocument          As Object
#If ImplUseShared = 0 Then
Private m_uPeekArray            As UcsSafeArraySingleDimension
Private m_aPeekBuffer()         As Integer
#End If ' ImplUseShared
'--- debug
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

Private Type VBGUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(0 To 7)       As Byte
End Type

Private Type UcsSstHashEntryType
    Index               As Long
    KeyPosition         As Long
    NextEntry           As Long
End Type

#If ImplUseShared = 0 Then
Private Type UcsSafeArraySingleDimension
    cDims       As Integer      '--- usually 1
    fFeatures   As Integer      '--- leave 0
    cbElements  As Long         '--- bytes per element (2-int, 4-long)
    cLocks      As Long         '--- leave 0
    pvData      As Long         '--- ptr to data
    cElements   As Long         '--- UBound + 1
    lLbound     As Long         '--- LBound
End Type
#End If ' ImplUseShared

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

Private Function RaiseError(sFunction As String) As VbMsgBoxResult
    Dim vErr            As Variant
    
    PushError vErr
'    RaiseError = GApp.HandleOutOfMemory(vErr)
    If RaiseError <> vbRetry Then
        PopRaiseError sFunction, MODULE_NAME, vErr
    End If
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get Position() As Long
    Position = m_lPosition
End Property

Property Let Position(ByVal lValue As Long)
    m_lPosition = lValue
End Property

Property Get Ptr(Optional ByVal lPosition As Long) As Long
    If lPosition >= 0 And lPosition < m_lSize Then
        Ptr = VarPtr(m_baBuffer(lPosition))
    End If
End Property

Property Get Size() As Long
    Size = m_lSize
End Property

Property Get Contents() As Byte()
    If Not m_oXmlDocument Is Nothing Then
        Contents = ToUtf8Array(m_oXmlDocument.XML)
    ElseIf m_lSize > 0 Then
        ReDim Preserve m_baBuffer(0 To m_lSize - 1) As Byte
        Contents = m_baBuffer
    Else
        Contents = ApiEmptyByteArray()
    End If
End Property

Property Let Contents(baValue() As Byte)
    If Not m_oXmlDocument Is Nothing Then
        m_oXmlDocument.XML = FromUtf8Array(baValue)
    Else
        m_baBuffer = baValue
        m_lPosition = 0
        If Peek(ArrPtr(baValue)) <> 0 Then
            m_lSize = UBound(baValue) + 1
        Else
            m_lSize = 0
        End If
    End If
End Property

Property Get Name() As String
    Name = m_sName
End Property

Property Let Name(sValue As String)
    m_sName = sValue
End Property

Property Get SheetName() As String
    SheetName = m_sSheetName
End Property

Property Let SheetName(sValue As String)
    m_sSheetName = sValue
End Property

Property Get Type_() As String
    Type_ = m_sType
End Property

Property Let Type_(sValue As String)
    m_sType = sValue
End Property

Property Get XmlDocument() As Object
    Set XmlDocument = m_oXmlDocument
End Property

Property Set XmlDocument(oValue As Object)
    Set m_oXmlDocument = oValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function ReadByte() As Long
    If m_lPosition >= 0 And m_lPosition < m_lSize - 1 Then
        ReadByte = m_baBuffer(m_lPosition)
        m_lPosition = m_lPosition + 1
    Else
        m_lPosition = -1
    End If
End Function

Public Function ReadWord() As Long
    Dim nRetVal         As Integer
    
    If m_lPosition >= 0 And m_lPosition < m_lSize - 1 Then
        Call CopyMemory(nRetVal, m_baBuffer(m_lPosition), 2)
        m_lPosition = m_lPosition + 2
    Else
        m_lPosition = -1
    End If
    ReadWord = nRetVal
End Function

Public Function Read3Bytes() As Long
    If m_lPosition >= 0 And m_lPosition < m_lSize - 2 Then
        Call CopyMemory(Read3Bytes, m_baBuffer(m_lPosition), 3)
        m_lPosition = m_lPosition + 3
    Else
        m_lPosition = -1
    End If
End Function

Public Function ReadDWord() As Long
    If m_lPosition >= 0 And m_lPosition < m_lSize - 3 Then
        Call CopyMemory(ReadDWord, m_baBuffer(m_lPosition), 4)
        m_lPosition = m_lPosition + 4
    Else
        m_lPosition = -1
    End If
End Function

Public Function ReadQWord() As Variant '-- vbLongLong
    Dim vRetVal         As Variant
    
    Call VariantChangeType(vRetVal, vRetVal, 0, VT_I8)
    If m_lPosition >= 0 And m_lPosition < m_lSize - 7 Then
        Call CopyMemory(ByVal VarPtr(vRetVal) + 8, m_baBuffer(m_lPosition), 8)
        m_lPosition = m_lPosition + 8
    Else
        m_lPosition = -1
    End If
    ReadQWord = vRetVal
End Function

Public Function ReadString() As String
    Dim lSize           As Long
    Dim sRetVal         As String
    
    lSize = ReadDWord() * 2
    If lSize > 0 Then
        If m_lPosition >= 0 And m_lPosition < m_lSize + 1 - lSize Then
            sRetVal = String$(lSize \ 2, 0)
            Call CopyMemory(ByVal StrPtr(sRetVal), m_baBuffer(m_lPosition), lSize)
            m_lPosition = m_lPosition + lSize
        Else
            m_lPosition = -1
        End If
    End If
    ReadString = sRetVal
End Function

Public Function ReadVarDWord() As Long
    Dim lIdx            As Long
    
    If m_lPosition >= 0 And m_lPosition < m_lSize Then
        For lIdx = m_lPosition To m_lPosition + 3
            If Not (lIdx >= 0 And lIdx < m_lSize) Then
                Exit For
            ElseIf (m_baBuffer(lIdx) And &H80) = 0 Then
                Exit For
            End If
        Next
        Call CopyMemory(ReadVarDWord, m_baBuffer(m_lPosition), lIdx - m_lPosition + 1)
        ReadVarDWord = (ReadVarDWord And &H7F&) Or (ReadVarDWord And &H7F00&) \ 2 Or (ReadVarDWord And &H7F0000) \ 4 Or (ReadVarDWord And &H7F000000) \ 8
        If lIdx >= 0 And lIdx < m_lSize - 1 Then
            m_lPosition = lIdx + 1
        Else
            m_lPosition = -1
        End If
    Else
        ReadVarDWord = -1
        m_lPosition = -1
    End If
End Function

Public Function ReadDouble() As Double
    If m_lPosition >= 0 And m_lPosition < m_lSize - 7 Then
        Call CopyMemory(ReadDouble, m_baBuffer(m_lPosition), 8)
        m_lPosition = m_lPosition + 8
    Else
        m_lPosition = -1
    End If
End Function

Public Function ReadGuid() As String
    Dim uGuid       As VBGUID
    
    If m_lPosition >= 0 And m_lPosition < m_lSize - 15 Then
        Call CopyMemory(uGuid, m_baBuffer(m_lPosition), 16)
        With uGuid
            ReadGuid = "{" & pvHexPad(.Data1, 8) & _
                "-" & pvHexPad(.Data2, 4) & _
                "-" & pvHexPad(.Data3, 4) & _
                "-" & pvHexPad(.Data4(0), 2) & pvHexPad(.Data4(1), 2) & _
                "-" & pvHexPad(.Data4(2), 2) & pvHexPad(.Data4(3), 2) & pvHexPad(.Data4(4), 2) & _
                        pvHexPad(.Data4(5), 2) & pvHexPad(.Data4(6), 2) & pvHexPad(.Data4(7), 2) & "}"
        End With
        m_lPosition = m_lPosition + 16
    Else
        m_lPosition = -1
    End If
End Function

Public Function ReadBlob(lSize As Long) As Byte()
    Dim baRetVal()          As Byte
    
    If lSize > 0 Then
        If m_lPosition >= 0 And m_lPosition < m_lSize - (lSize - 1) Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            Call CopyMemory(baRetVal(0), m_baBuffer(m_lPosition), lSize)
            m_lPosition = m_lPosition + lSize
        Else
            baRetVal = ApiEmptyByteArray()
            m_lPosition = -1
        End If
    Else
        baRetVal = ApiEmptyByteArray()
    End If
    ReadBlob = baRetVal
End Function

'--- 2.5.122 RkNumber: The RkNumber structure specifies a numeric value.
Public Function ToRkNumber(dblValue As Double, Optional RetVal As Long) As Boolean
    Dim dblTemp         As Double
    Dim lFlags          As Long
    
    If Round(dblValue, 0) = dblValue Then
        If Abs(dblValue) <= LNG_RK_MAXINT Then
            RetVal = CLng(dblValue) * LNG_RK_NUMOFFSET
            lFlags = LNG_RK_FINT
        Else
            Call CopyMemory(RetVal, ByVal VarPtr(dblValue) + 4, 4)
        End If
    ElseIf Round(dblValue, 2) = dblValue Then
        dblTemp = dblValue * 100
        lFlags = LNG_RK_FX100
        If Abs(dblTemp) <= LNG_RK_MAXINT Then
            RetVal = CLng(dblTemp) * LNG_RK_NUMOFFSET
            lFlags = lFlags Or LNG_RK_FINT
        Else
            Call CopyMemory(RetVal, ByVal VarPtr(dblTemp) + 4, 4)
        End If
    Else
        Call CopyMemory(RetVal, ByVal VarPtr(dblValue) + 4, 4)
    End If
    RetVal = (RetVal And -LNG_RK_NUMOFFSET) Or lFlags
    ToRkNumber = FromRkNumber(RetVal) = dblValue
End Function

Public Function FromRkNumber(ByVal lValue As Long) As Double
    Dim lIdx            As Long
    
    If lValue <> 0 Then
        If (lValue And LNG_RK_FINT) <> 0 Then
            FromRkNumber = lValue \ LNG_RK_NUMOFFSET
        Else
            lIdx = lValue And -LNG_RK_NUMOFFSET
            Call CopyMemory(ByVal VarPtr(FromRkNumber) + 4, lIdx, 4)
        End If
        If (lValue And LNG_RK_FX100) <> 0 Then
            FromRkNumber = FromRkNumber / 100#
        End If
    End If
End Function

'= write =================================================================

Public Function WriteByte(ByVal lValue As Long) As Long
    WriteByte = m_lPosition
    pvEnsureCapacity m_lPosition + 1
    m_baBuffer(m_lPosition) = lValue And &HFF&
    m_lPosition = m_lPosition + 1
End Function

Public Function WriteWord(ByVal lValue As Long) As Long
    WriteWord = m_lPosition
    pvEnsureCapacity m_lPosition + 2
    Call CopyMemory(m_baBuffer(m_lPosition), lValue, 2)
    m_lPosition = m_lPosition + 2
End Function

Public Function WriteDWord(ByVal lValue As Long) As Long
    WriteDWord = m_lPosition
    pvEnsureCapacity m_lPosition + 4
    Call CopyMemory(m_baBuffer(m_lPosition), lValue, 4)
    m_lPosition = m_lPosition + 4
End Function

Public Function WriteString(sValue As String) As Long
    WriteString = m_lPosition
    pvEnsureCapacity m_lPosition + 4 + LenB(sValue)
    Call CopyMemory(m_baBuffer(m_lPosition), Len(sValue), 4)
    Call CopyMemory(m_baBuffer(m_lPosition + 4), ByVal StrPtr(sValue), LenB(sValue))
    m_lPosition = m_lPosition + 4 + LenB(sValue)
End Function

Public Function WriteVarDWord(ByVal lValue As Long) As Long
    Dim lEnc            As Long
    Dim lSize           As Long
    
    WriteVarDWord = m_lPosition
    lEnc = (lValue And &H7F&) Or (lValue * 2 And &H7F00&) Or (lValue * 4 And &H7F0000) Or (lValue * 8 And &H7F000000)
    If lEnc < &HFF& Then
        lSize = 1
    ElseIf lEnc < &HFFFF& Then
        lEnc = lEnc Or &H80&
        lSize = 2
    ElseIf lEnc < &HFFFFFF Then
        lEnc = lEnc Or &H8080&
        lSize = 3
    Else
        lEnc = lEnc Or &H808080
        lSize = 4
    End If
    pvEnsureCapacity m_lPosition + lSize
    Call CopyMemory(m_baBuffer(m_lPosition), lEnc, lSize)
    m_lPosition = m_lPosition + lSize
End Function

Public Function WriteDouble(dblValue As Double) As Long
    WriteDouble = m_lPosition
    pvEnsureCapacity m_lPosition + 8
    Call CopyMemory(m_baBuffer(m_lPosition), dblValue, 8)
    m_lPosition = m_lPosition + 8
End Function

Public Function WriteGuid(sValue As String) As Long
    Dim uGuid           As VBGUID
    
    If LenB(sValue) <> 0 Then
        Call CLSIDFromString(StrPtr(sValue), uGuid)
    End If
    WriteGuid = m_lPosition
    pvEnsureCapacity m_lPosition + 16
    Call CopyMemory(m_baBuffer(m_lPosition), uGuid.Data1, 16)
    m_lPosition = m_lPosition + 16
End Function

Public Function WriteBlob(ByVal lPtr As Long, ByVal lSize As Long) As Long
    WriteBlob = m_lPosition
    pvEnsureCapacity m_lPosition + lSize
    Call CopyMemory(m_baBuffer(m_lPosition), ByVal lPtr, lSize)
    m_lPosition = m_lPosition + lSize
End Function

Public Function WriteRecord(ByVal eRecID As UcsBiff12RecortTypeEnum, Optional ByVal lSize As Long) As Long
    WriteVarDWord eRecID
    WriteVarDWord lSize
    WriteRecord = m_lPosition
End Function

'= output generic ========================================================

Public Function Output(ByVal eRecID As UcsBiff12RecortTypeEnum) As Long
    Output = WriteRecord(eRecID, 0)
End Function

Public Function OutputCount(ByVal eRecID As UcsBiff12RecortTypeEnum, ByVal lCount As Long) As Long
    Const FIXED_SIZE    As Long = 4
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCount = WriteRecord(eRecID, FIXED_SIZE)
    WriteDWord lCount
    Debug.Assert OutputCount + lSize = m_lPosition
End Function

Public Function OutputCount2(ByVal eRecID As UcsBiff12RecortTypeEnum, ByVal lCount As Long, ByVal lUnique As Long) As Long
    Const FIXED_SIZE    As Long = 8
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCount2 = WriteRecord(eRecID, FIXED_SIZE)
    WriteDWord lCount
    WriteDWord lUnique
    Debug.Assert OutputCount2 + lSize = m_lPosition
End Function

Public Function OutputCount3(ByVal eRecID As UcsBiff12RecortTypeEnum, ByVal lCount As Long, ByVal lSecond As Long, ByVal lThird As Long) As Long
    Const FIXED_SIZE    As Long = 12
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCount3 = WriteRecord(eRecID, FIXED_SIZE)
    WriteDWord lCount
    WriteDWord lSecond
    WriteDWord lThird
    Debug.Assert OutputCount3 + lSize = m_lPosition
End Function

'= output cells ==========================================================

Public Function OutputCellBlank(ByVal lColIndex As Long, ByVal lStyleRefIdx As Long) As Long
    Const FIXED_SIZE    As Long = 8
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCellBlank = WriteRecord(ucsBrtCellBlank, FIXED_SIZE)
    WriteDWord lColIndex
    WriteDWord lStyleRefIdx
    Debug.Assert OutputCellBlank + lSize = m_lPosition
End Function

Public Function OutputCellRString(ByVal lColIndex As Long, ByVal lStyleRefIdx As Long, sText As String) As Long
    Const FIXED_SIZE    As Long = 9
    Dim lSize           As Long
    
    lSize = FIXED_SIZE + 4 + LenB(sText)
    OutputCellRString = WriteRecord(ucsBrtCellRString, lSize)
    WriteDWord lColIndex
    WriteDWord lStyleRefIdx
    WriteByte 0
    WriteString sText
    Debug.Assert OutputCellRString + lSize = m_lPosition
End Function

Public Function OutputCellIsst(ByVal lColIndex As Long, ByVal lStyleRefIdx As Long, ByVal lSstIdx As Long) As Long
    Const FIXED_SIZE    As Long = 12
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCellIsst = WriteRecord(ucsBrtCellIsst, lSize)
    WriteDWord lColIndex
    WriteDWord lStyleRefIdx
    WriteDWord lSstIdx
    Debug.Assert OutputCellIsst + lSize = m_lPosition
End Function

Public Function OutputCellRk(ByVal lColIndex As Long, ByVal lStyleRefIdx As Long, ByVal lRkNumber As Long) As Long
    Const FIXED_SIZE    As Long = 12
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCellRk = WriteRecord(ucsBrtCellRk, lSize)
    WriteDWord lColIndex
    WriteDWord lStyleRefIdx
    WriteDWord lRkNumber
    Debug.Assert OutputCellRk + lSize = m_lPosition
End Function

Public Function OutputCellReal(ByVal lColIndex As Long, ByVal lStyleRefIdx As Long, dblNum As Double) As Long
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long
    
    lSize = FIXED_SIZE
    OutputCellReal = WriteRecord(ucsBrtCellReal, lSize)
    WriteDWord lColIndex
    WriteDWord lStyleRefIdx
    WriteDouble dblNum
    Debug.Assert OutputCellReal + lSize = m_lPosition
End Function

'= output brt structs ====================================================

#If ImplPublicClasses Then
Public Function OutputBrtFmt(uFmt As UcsBiff12BrtFmtType) As Long
#Else
Friend Function OutputBrtFmt(uFmt As UcsBiff12BrtFmtType) As Long
#End If
    Const FIXED_SIZE    As Long = 2
    Dim lSize           As Long
    
    lSize = FIXED_SIZE + 4 + LenB(uFmt.m_stFmtCode)
    OutputBrtFmt = WriteRecord(ucsBrtFmt, lSize)
    Debug.Assert VarPtr(uFmt.m_iFmt) + 2 - VarPtr(uFmt) = FIXED_SIZE
    WriteBlob VarPtr(uFmt), FIXED_SIZE
    WriteString uFmt.m_stFmtCode
    Debug.Assert OutputBrtFmt + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtFont(uFont As UcsBiff12BrtFontType) As Long
#Else
Friend Function OutputBrtFont(uFont As UcsBiff12BrtFontType) As Long
#End If
    Const FIXED_SIZE    As Long = 21
    Dim lSize           As Long
    
    lSize = FIXED_SIZE + 4 + LenB(uFont.m_name)
    OutputBrtFont = WriteRecord(ucsBrtFont, lSize)
    Debug.Assert VarPtr(uFont.m_bFontScheme) + 1 - VarPtr(uFont) = FIXED_SIZE
    WriteBlob VarPtr(uFont), FIXED_SIZE
    WriteString uFont.m_name
    Debug.Assert OutputBrtFont + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtFill(uFill As UcsBiff12BrtFillType) As Long
#Else
Friend Function OutputBrtFill(uFill As UcsBiff12BrtFillType) As Long
#End If
    Const FIXED_SIZE    As Long = 68
    Const GRSTOP_SIZE   As Long = 16
    Dim lSize           As Long
    
    lSize = FIXED_SIZE + uFill.m_cNumStop * GRSTOP_SIZE
    OutputBrtFill = WriteRecord(ucsBrtFill, lSize)
    Debug.Assert VarPtr(uFill.m_cNumStop) + 4 - VarPtr(uFill) = FIXED_SIZE
    WriteBlob VarPtr(uFill), FIXED_SIZE
    If uFill.m_cNumStop > 0 Then
        Debug.Assert LenB(uFill.m_xfillGradientStop(0)) = GRSTOP_SIZE
        WriteBlob VarPtr(uFill.m_xfillGradientStop(0)), uFill.m_cNumStop * GRSTOP_SIZE
    End If
    Debug.Assert OutputBrtFill + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtBorder(uBorder As UcsBiff12BrtBorderType) As Long
#Else
Friend Function OutputBrtBorder(uBorder As UcsBiff12BrtBorderType) As Long
#End If
    Const FIXED_SIZE    As Long = 50
    Dim lSize           As Long

    lSize = FIXED_SIZE + 1
    OutputBrtBorder = WriteRecord(ucsBrtBorder, lSize)
    WriteByte uBorder.m_flags
    Debug.Assert VarPtr(uBorder.m_blxfDiag) + 10 - VarPtr(uBorder.m_blxfTop) = FIXED_SIZE
    WriteBlob VarPtr(uBorder.m_blxfTop), FIXED_SIZE
    Debug.Assert OutputBrtBorder + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtXf(uXf As UcsBiff12BrtXfType) As Long
#Else
Friend Function OutputBrtXf(uXf As UcsBiff12BrtXfType) As Long
#End If
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtXf = WriteRecord(ucsBrtXF, lSize)
    Debug.Assert VarPtr(uXf.m_xfGrbitAtr) + 2 - VarPtr(uXf) = FIXED_SIZE
    WriteBlob VarPtr(uXf), FIXED_SIZE
    Debug.Assert OutputBrtXf + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtStyle(uStyle As UcsBiff12BrtStyleType) As Long
#Else
Friend Function OutputBrtStyle(uStyle As UcsBiff12BrtStyleType) As Long
#End If
    Const FIXED_SIZE    As Long = 8
    Dim lSize           As Long

    lSize = FIXED_SIZE + 4 + LenB(uStyle.m_stName)
    OutputBrtStyle = WriteRecord(ucsBrtStyle, lSize)
    Debug.Assert VarPtr(uStyle.m_iLevel) + 1 - VarPtr(uStyle) = FIXED_SIZE
    WriteBlob VarPtr(uStyle), FIXED_SIZE
    WriteString uStyle.m_stName
    Debug.Assert OutputBrtStyle + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtWbProp(uWbProp As UcsBiff12BrtWbPropType) As Long
#Else
Friend Function OutputBrtWbProp(uWbProp As UcsBiff12BrtWbPropType) As Long
#End If
    Const FIXED_SIZE    As Long = 8
    Dim lSize           As Long

    lSize = FIXED_SIZE + 4 + LenB(uWbProp.m_strName)
    OutputBrtWbProp = WriteRecord(ucsBrtWbProp, lSize)
    Debug.Assert VarPtr(uWbProp.m_dwThemeVersion) + 4 - VarPtr(uWbProp) = FIXED_SIZE
    WriteBlob VarPtr(uWbProp), FIXED_SIZE
    WriteString uWbProp.m_strName
    Debug.Assert OutputBrtWbProp + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtBookView(uBookView As UcsBiff12BrtBookViewType) As Long
#Else
Friend Function OutputBrtBookView(uBookView As UcsBiff12BrtBookViewType) As Long
#End If
    Const FIXED_SIZE    As Long = 29
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtBookView = WriteRecord(ucsBrtBookView, lSize)
    Debug.Assert VarPtr(uBookView.m_flags) + 1 - VarPtr(uBookView) = FIXED_SIZE
    WriteBlob VarPtr(uBookView), FIXED_SIZE
    Debug.Assert OutputBrtBookView + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtBundleSh(uBundle As UcsBiff12BrtBundleShType) As Long
#Else
Friend Function OutputBrtBundleSh(uBundle As UcsBiff12BrtBundleShType) As Long
#End If
    Const FIXED_SIZE    As Long = 8
    Dim lSize           As Long

    lSize = FIXED_SIZE + 4 + LenB(uBundle.m_strRelID) + 4 + LenB(uBundle.m_strName)
    OutputBrtBundleSh = WriteRecord(ucsBrtBundleSh, lSize)
    Debug.Assert VarPtr(uBundle.m_iTabID) + 4 - VarPtr(uBundle) = FIXED_SIZE
    WriteBlob VarPtr(uBundle), FIXED_SIZE
    WriteString uBundle.m_strRelID
    WriteString uBundle.m_strName
    Debug.Assert OutputBrtBundleSh + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtWsDim(uWsDim As UcsBiff12UncheckedRfXType) As Long
#Else
Friend Function OutputBrtWsDim(uWsDim As UcsBiff12UncheckedRfXType) As Long
#End If
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtWsDim = WriteRecord(ucsBrtWsDim, lSize)
    Debug.Assert LenB(uWsDim) = FIXED_SIZE
    WriteBlob VarPtr(uWsDim), FIXED_SIZE
    Debug.Assert OutputBrtWsDim + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtWsProp(uWsProp As UcsBiff12BrtWsPropType) As Long
#Else
Friend Function OutputBrtWsProp(uWsProp As UcsBiff12BrtWsPropType) As Long
#End If
    Const FLAGS_SIZE    As Long = 3
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FLAGS_SIZE + FIXED_SIZE + 4 + LenB(uWsProp.m_strName)
    OutputBrtWsProp = WriteRecord(ucsBrtWsProp, lSize)
    WriteBlob VarPtr(uWsProp.m_flags), FLAGS_SIZE
    Debug.Assert VarPtr(uWsProp.m_colSync) + 4 - VarPtr(uWsProp.m_brtcolorTab) = FIXED_SIZE
    WriteBlob VarPtr(uWsProp.m_brtcolorTab), FIXED_SIZE
    WriteString uWsProp.m_strName
    Debug.Assert OutputBrtWsProp + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtColInfo(uColInfo As UcsBiff12BrtColInfoType) As Long
#Else
Friend Function OutputBrtColInfo(uColInfo As UcsBiff12BrtColInfoType) As Long
#End If
    Const FIXED_SIZE    As Long = 18
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtColInfo = WriteRecord(ucsBrtColInfo, lSize)
    Debug.Assert VarPtr(uColInfo.m_flags) + 2 - VarPtr(uColInfo) = FIXED_SIZE
    WriteBlob VarPtr(uColInfo), FIXED_SIZE
    Debug.Assert OutputBrtColInfo + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtRowHdr(uRowHdr As UcsBiff12BrtRowHdrType) As Long
#Else
Friend Function OutputBrtRowHdr(uRowHdr As UcsBiff12BrtRowHdrType) As Long
#End If
    Const FIXED_SIZE    As Long = 10
    Const COLSPAN_SIZE  As Long = 8
    Dim lSize           As Long

    lSize = FIXED_SIZE + 3 + 4 + uRowHdr.m_ccolspan * 8
    OutputBrtRowHdr = WriteRecord(ucsBrtRowHdr, lSize)
    Debug.Assert VarPtr(uRowHdr.m_miyRw) + 2 - VarPtr(uRowHdr) = FIXED_SIZE
    WriteBlob VarPtr(uRowHdr), FIXED_SIZE
    WriteBlob VarPtr(uRowHdr.m_flags), 3
    WriteDWord uRowHdr.m_ccolspan
    If uRowHdr.m_ccolspan > 0 Then
        Debug.Assert LenB(uRowHdr.m_rgBrtColspan(0)) = COLSPAN_SIZE
        WriteBlob VarPtr(uRowHdr.m_rgBrtColspan(0)), uRowHdr.m_ccolspan * COLSPAN_SIZE
    End If
    Debug.Assert OutputBrtRowHdr + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtMergeCell(uMergeCell As UcsBiff12UncheckedRfXType) As Long
#Else
Friend Function OutputBrtMergeCell(uMergeCell As UcsBiff12UncheckedRfXType) As Long
#End If
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtMergeCell = WriteRecord(ucsBrtMergeCell, lSize)
    Debug.Assert VarPtr(uMergeCell.m_colLast) + 4 - VarPtr(uMergeCell) = FIXED_SIZE
    WriteBlob VarPtr(uMergeCell), FIXED_SIZE
    Debug.Assert OutputBrtMergeCell + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtOleSize(uOleSize As UcsBiff12UncheckedRfXType) As Long
#Else
Friend Function OutputBrtOleSize(uOleSize As UcsBiff12UncheckedRfXType) As Long
#End If
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FIXED_SIZE
    OutputBrtOleSize = WriteRecord(ucsBrtOleSize, lSize)
    Debug.Assert LenB(uOleSize) = FIXED_SIZE
    WriteBlob VarPtr(uOleSize), FIXED_SIZE
    Debug.Assert OutputBrtOleSize + lSize = m_lPosition
End Function

#If ImplPublicClasses Then
Public Function OutputBrtFileVersion(uFileVersion As UcsBiff12BrtFileVersionType) As Long
#Else
Friend Function OutputBrtFileVersion(uFileVersion As UcsBiff12BrtFileVersionType) As Long
#End If
    Const FIXED_SIZE    As Long = 16
    Dim lSize           As Long

    lSize = FIXED_SIZE + 4 + LenB(uFileVersion.m_stAppName) + 4 + LenB(uFileVersion.m_stLastEdited) + 4 + LenB(uFileVersion.m_stLowestEdited) + 4 + LenB(uFileVersion.m_stRupBuild)
    OutputBrtFileVersion = WriteRecord(ucsBrtFileVersion, lSize)
    WriteGuid uFileVersion.m_guidCodeName
    WriteString uFileVersion.m_stAppName
    WriteString uFileVersion.m_stLastEdited
    WriteString uFileVersion.m_stLowestEdited
    WriteString uFileVersion.m_stRupBuild
    Debug.Assert OutputBrtFileVersion + lSize = m_lPosition
End Function

Public Function OutputBrtDrawing(sRelId As String) As Long
    Dim lSize           As Long
    
    lSize = 4 + LenB(sRelId)
    OutputBrtDrawing = WriteRecord(ucsBrtDrawing, lSize)
    WriteString sRelId
    Debug.Assert OutputBrtDrawing + lSize = m_lPosition
End Function

Public Function OutputBrtSstItem(sText As String) As Long
    Dim lSize           As Long
    
    lSize = 1 + 4 + LenB(sText)
    OutputBrtSstItem = WriteRecord(ucsBrtSSTItem, lSize)
    WriteByte 0
    WriteString sText
    Debug.Assert OutputBrtSstItem + lSize = m_lPosition
End Function

'= shared strings table ==================================================

Public Function SstGetIndex(sText As String) As Long
    SstGetIndex = m_uaSstHashEntries(pvSstGetInternal(sText)).Index
End Function

Public Function SstGetPosition(sText As String) As Long
    SstGetPosition = m_uaSstHashEntries(pvSstGetInternal(sText)).KeyPosition
End Function

Public Function pvSstGetInternal(sText As String) As Long
    Const FUNC_NAME     As String = "SstGetIndex"
    Dim lIdx            As Long
    Dim lPrevIdx        As Long
    Dim lPrevPosition   As Long
    
    On Error GoTo EH
    If m_lSstCount = 0 Then
        ReDim m_uaSstHashEntries(0 To LNG_SST_HASH_SIZE - 1 + 256) As UcsSstHashEntryType
    End If
    lIdx = RollingHash(StrPtr(sText), Len(sText)) And (LNG_SST_HASH_SIZE - 1)
    If m_uaSstHashEntries(lIdx).KeyPosition = 0 Then
        m_uaSstHashEntries(lIdx).KeyPosition = OutputBrtSstItem(sText) + 1
        m_uaSstHashEntries(lIdx).Index = m_lSstCount
        m_uaSstHashEntries(lIdx).NextEntry = -1
        m_lSstCount = m_lSstCount + 1
    Else
        lPrevPosition = m_lPosition
        Do
            m_lPosition = m_uaSstHashEntries(lIdx).KeyPosition
            If ReadDWord() = Len(sText) Then
                m_lPosition = m_uaSstHashEntries(lIdx).KeyPosition
                If ReadString() = sText Then
                    Exit Do
                End If
            End If
            lPrevIdx = lIdx
            lIdx = m_uaSstHashEntries(lPrevIdx).NextEntry
            If lIdx = -1 Then
                lIdx = LNG_SST_HASH_SIZE + m_lSstCapacity
                m_uaSstHashEntries(lPrevIdx).NextEntry = lIdx
                m_lSstCapacity = m_lSstCapacity + 1
                If m_lSstCapacity > UBound(m_uaSstHashEntries) - LNG_SST_HASH_SIZE Then
                    ReDim Preserve m_uaSstHashEntries(0 To LNG_SST_HASH_SIZE + 2 * m_lSstCapacity) As UcsSstHashEntryType
                End If
                m_lPosition = lPrevPosition
                m_uaSstHashEntries(lIdx).KeyPosition = OutputBrtSstItem(sText) + 1
                m_uaSstHashEntries(lIdx).Index = m_lSstCount
                m_uaSstHashEntries(lIdx).NextEntry = -1
                m_lSstCount = m_lSstCount + 1
                lPrevPosition = m_lPosition
                Exit Do
            End If
        Loop
        m_lPosition = lPrevPosition
    End If
    Erase m_laSstKeyPositions
    pvSstGetInternal = lIdx
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function SstFromPosition(ByVal lPos As Long) As String
    Const FUNC_NAME     As String = "SstFromPosition"
    Dim lPrevPosition   As Long
    
    On Error GoTo EH
    If lPos > 0 Then
        lPrevPosition = m_lPosition
        m_lPosition = lPos
        SstFromPosition = ReadString()
        m_lPosition = lPrevPosition
    End If
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function SstSizeFromPosition(ByVal lPos As Long) As Long
    Dim lPrevPosition   As Long
    Const FUNC_NAME     As String = "SstSizeFromPosition"
    
    On Error GoTo EH
    SstSizeFromPosition = 0
    If lPos > 0 Then
        lPrevPosition = m_lPosition
        m_lPosition = lPos
        SstSizeFromPosition = ReadDWord()
        m_lPosition = lPrevPosition
    End If
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function SstFromIndex(ByVal lIndex As Long) As String
    Const FUNC_NAME     As String = "SstFromIndex"
    Dim lIdx            As Long

    On Error GoTo EH
    If lIndex < m_lSstCount Then
        If Peek(ArrPtr(m_laSstKeyPositions)) = 0 Then
            ReDim m_laSstKeyPositions(0 To m_lSstCount - 1) As Long
            For lIdx = 0 To UBound(m_uaSstHashEntries)
                With m_uaSstHashEntries(lIdx)
                    If .KeyPosition <> 0 Then
                        m_laSstKeyPositions(.Index) = .KeyPosition
                    End If
                End With
            Next
        End If
        SstFromIndex = SstFromPosition(m_laSstKeyPositions(lIndex))
    End If
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function SstDumpTable() As Variant
    Dim lPrevPosition   As Long
    Dim lIdx            As Long
    '--- note: array of Variants for easier rehydration in cBiff12Writer
    Dim aRetVal()       As Variant
    
    If m_lSstCount <= 0 Then
        SstDumpTable = Split(vbNullString)
        Exit Function
    End If
    ReDim aRetVal(0 To m_lSstCount - 1) As Variant
    lPrevPosition = m_lPosition
    For lIdx = 0 To UBound(m_uaSstHashEntries)
        With m_uaSstHashEntries(lIdx)
            If .KeyPosition <> 0 Then
                m_lPosition = .KeyPosition
                aRetVal(.Index) = ReadString()
            End If
        End With
    Next
    m_lPosition = lPrevPosition
    SstDumpTable = aRetVal
End Function

#If DebugMode Then
Public Function SstDumpInfo(Optional lEntriesUsed As Long, Optional lMaxListSize As Long) As Long
    Const FUNC_NAME     As String = "SstDumpInfo"
    Dim lIdx            As Long
    Dim lNext           As Long
    Dim lListSize       As Long
    
    On Error GoTo EH
    lEntriesUsed = 0
    lMaxListSize = 0
    For lIdx = 0 To LNG_SST_HASH_SIZE - 1
        If m_uaSstHashEntries(lIdx).KeyPosition <> 0 Then
            lEntriesUsed = lEntriesUsed + 1
            lNext = m_uaSstHashEntries(lIdx).NextEntry
            lListSize = 0
            Do While lNext <> -1
                lListSize = lListSize + 1
                lNext = m_uaSstHashEntries(lNext).NextEntry
            Loop
            If lListSize > lMaxListSize Then
                lMaxListSize = lListSize
            End If
        End If
    Next
    SstDumpInfo = m_lSstCount
    DebugPrint FUNC_NAME, MODULE_NAME, "EntriesUsed=" & lEntriesUsed & ", MaxListSize=" & lMaxListSize & ", AvgListSize=" & Format$(CDbl(m_lSstCount - lEntriesUsed) / Znl(lEntriesUsed, 1), "0.0")
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function
#End If

'= zip VFS ===============================================================

Public Function VfsSetFilePointer(lPosition As Long, ByVal lMoveMethod As Long) As Long
    If Not m_oXmlDocument Is Nothing Then
        If Peek(ArrPtr(m_baBuffer)) = 0 Then
            m_baBuffer = ToUtf8Array(m_oXmlDocument.XML)
            m_lPosition = 0
            m_lSize = UBound(m_baBuffer) + 1
        End If
    End If
    Select Case lMoveMethod
    Case FILE_BEGIN
        m_lPosition = lPosition
    Case FILE_CURRENT
        m_lPosition = m_lPosition + lPosition
    Case FILE_END
        m_lPosition = m_lSize + lPosition
    End Select
    VfsSetFilePointer = m_lPosition
End Function

Public Function VfsReadFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    If m_lPosition + lSize > m_lSize Then
        lSize = m_lSize - m_lPosition
    End If
    If lSize > 0 Then
        Call CopyMemory(ByVal lPtr, m_baBuffer(m_lPosition), lSize)
        m_lPosition = m_lPosition + lSize
    End If
    VfsReadFile = lSize
End Function

Public Function VfsWriteFile(ByVal lPtr As Long, ByVal lSize As Long) As Long
    pvEnsureCapacity m_lPosition + lSize
    If lSize > 0 Then
        Call CopyMemory(m_baBuffer(m_lPosition), ByVal lPtr, lSize)
        m_lPosition = m_lPosition + lSize
        If m_lPosition > m_lSize Then
            m_lSize = m_lPosition
        End If
    End If
    VfsWriteFile = lSize
End Function

Public Sub VfsSetEndOfFile(sMetaData As String)
    #If sMetaData Then '--- touch arg
    #End If
    m_lSize = m_lPosition
End Sub

'= private ===============================================================

Private Sub pvEnsureCapacity(ByVal lNewSize As Long)
    Const MAX_STEP      As Long = 2 ^ 24 '--- 16MB
    Dim lCapacity       As Long
    
    If lNewSize > m_lSize Then
        If m_lSize = 0 Then
            lCapacity = 2 ^ Int(Log(lNewSize) / Log(2) + 1)
            ReDim m_baBuffer(0 To lCapacity - 1) As Byte
        ElseIf lNewSize > UBound(m_baBuffer) Then
            If lNewSize <= MAX_STEP Then
                lCapacity = 2 ^ Int(Log(lNewSize) / Log(2) + 1)
            Else
                lCapacity = (lNewSize + MAX_STEP - 1) And -MAX_STEP
            End If
            ReDim Preserve m_baBuffer(0 To lCapacity - 1) As Byte
        End If
        m_lSize = lNewSize
    End If
End Sub

Private Function pvHexPad(ByVal lValue As Long, ByVal lSize As Long) As String
    pvHexPad = Right$(String$(lSize, "0") & Hex$(lValue), lSize)
End Function

#If ImplUseShared = 0 Then

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

Private Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    FromUtf8Array = String$(2 * UBound(baText), 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
    FromUtf8Array = Left$(FromUtf8Array, lSize)
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function RollingHash(ByVal lPtr As Long, ByVal lSize As Long) As Long
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

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    '--- clear peek buffer safe array
    RollingHash 0, 0
    #If DebugMode Then
        DebugInstanceTerm STR_MODULE_NAME, m_sDebugID
    #End If
End Sub

#Else

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugInstanceTerm STR_MODULE_NAME, m_sDebugID
    End Sub
#End If

#End If 'ImplUseShared

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit STR_MODULE_NAME, m_sDebugID, Me
    End Sub
#End If
