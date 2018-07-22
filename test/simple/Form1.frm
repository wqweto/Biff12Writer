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
    oWriter.Init COL_COUNT
    '--- first row
    oWriter.AddRow
    oWriter.AddStringCell 0, txtUserName.Text, oStyle
    oWriter.AddImage 1, SaveToPng(picAvatar.Picture), 0, 0, picAvatar.Picture.Width * PIC_SCALE, picAvatar.Picture.Height * PIC_SCALE
    '--- second row
    oWriter.AddRow
    oWriter.AddStringCell 0, "Profile picture:", oStyle
    '--- third row
    oWriter.AddRow
    oWriter.AddStringCell 0, "More info here...", oStyle
    '--- save
    sFile = Environ$("TMP") & "\output.xlsb"
    oWriter.SaveToFile sFile
    MsgBox sFile & " saved sucessfully", vbExclamation
End Sub

Private Function SaveToPng(oPic As StdPicture) As Byte()
    With New cDibSection
        .LoadFromPicture oPic
        SaveToPng = .SaveToByteArray("image/png")
    End With
End Function
