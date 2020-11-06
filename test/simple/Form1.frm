VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3156
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4452
   LinkTopic       =   "Form1"
   ScaleHeight     =   3156
   ScaleWidth      =   4452
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAvatar 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   1272
      Left            =   1260
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1272
      ScaleWidth      =   2112
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   588
      Width           =   2112
   End
   Begin VB.TextBox txtUserName 
      Height          =   372
      Left            =   1260
      TabIndex        =   1
      Text            =   "Mr. Smith"
      Top             =   84
      Width           =   2028
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   432
      Left            =   1260
      TabIndex        =   0
      Top             =   2268
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "Avatar:"
      Height          =   348
      Left            =   168
      TabIndex        =   4
      Top             =   588
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   432
      Left            =   168
      TabIndex        =   2
      Top             =   84
      Width           =   936
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Biff12Writer (c) 2018 by wqweto@gmail.com
'
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
'
'=========================================================================
Option Explicit
DefObj A-Z

Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long

Private Sub Command1_Click()
    Const COL_COUNT     As Long = 5
    Const PIC_SCALE     As Long = 650
    Dim oStyle          As cBiff12CellStyle
    Dim oWriter         As cBiff12Writer
    Dim sFile           As String
    
    '-- setup styles to be used
    Set oStyle = New cBiff12CellStyle
    oStyle.FontName = "Tahoma"
    oStyle.FontSize = 9
    '--- setup xlsb writer
    Set oWriter = New cBiff12Writer
    oWriter.Init COL_COUNT, UseSST:=True
    '--- first row
    oWriter.AddRow
    oWriter.AddStringCell 0, txtUserName.Text, oStyle
    oWriter.AddImage 1, SaveAsPng(picAvatar.Picture), 0, 0, picAvatar.Picture.Width * PIC_SCALE, picAvatar.Picture.Height * PIC_SCALE
    '--- second row
    oWriter.AddRow
    oWriter.AddStringCell 0, "Profile picture:", oStyle
    '--- third row
    oWriter.AddRow
    oWriter.AddStringCell 0, "More info here...", oStyle
    '--- save
    sFile = Environ$("TMP") & "\output.xlsb"
    oWriter.SaveToFile sFile
    If MsgBox(sFile & " saved sucessfully", vbExclamation Or vbOKCancel) = vbOK Then
        Shell "cmd /c " & Environ$("TMP") & "\output.xlsb"
    End If
End Sub

Public Function SaveAsPng(pPic As IPicture) As Byte()
    Const adTypeBinary  As Long = 1
    Const wiaFormatPNG  As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Const CC_STDCALL    As Long = 4
    Dim oStream         As Object ' ADODB.Stream
    Dim oImageFile      As Object ' WIA.ImageFile
    Dim IID_IStream(3)  As Long
    Dim pStream         As IUnknown
    Dim vParams(0 To 1) As Variant
    Dim vType(0 To 1)   As Integer
    Dim vPtr(0 To 1)    As Long
    
    '--- load pPic in WIA.ImageFile
    Do While oImageFile Is Nothing
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Type = adTypeBinary
        oStream.Open
        '--- call IUnknown::QI on oStream for IStream interface and store in pStream
        IID_IStream(0) = &HC
        IID_IStream(2) = &HC0
        IID_IStream(3) = &H46000000
        vParams(0) = VarPtr(IID_IStream(0))
        vParams(1) = VarPtr(pStream)
        vType(0) = VarType(vParams(0))
        vType(1) = VarType(vParams(1))
        vPtr(0) = VarPtr(vParams(0))
        vPtr(1) = VarPtr(vParams(1))
        Call DispCallFunc(ObjPtr(oStream), 0, CC_STDCALL, vbLong, UBound(vParams) + 1, vType(0), vPtr(0), Empty)
        '--- NO magic anymore, only business as usual
        pPic.SaveAsFile ByVal ObjPtr(pStream), True, 0
        If oStream.Size = 0 Then
            GoTo QH
        End If
        oStream.Position = 0
        With CreateObject("WIA.Vector")
            .BinaryData = oStream.Read
            If pPic.Type <> vbPicTypeBitmap Then
                '--- this converts pPic to vbPicTypeBitmap subtype
                Set pPic = .Picture
            Else
                Set oImageFile = .ImageFile
            End If
        End With
    Loop
    '--- serialize WIA.ImageFile to PNG file format
    With CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("Convert").FilterID
        .Filters(.Filters.Count).Properties("FormatID").Value = wiaFormatPNG
        SaveAsPng = .Apply(oImageFile).FileData.BinaryData
    End With
QH:
End Function
