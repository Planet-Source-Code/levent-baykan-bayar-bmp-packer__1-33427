VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "InfoBox"
      Height          =   1140
      Left            =   45
      TabIndex        =   3
      Top             =   2610
      Width           =   1905
      Begin VB.Label Label2 
         Height          =   870
         Left            =   45
         TabIndex        =   4
         Top             =   225
         Width           =   1815
      End
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Click on an item to draw it to the right"
      Top             =   360
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   2025
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   0
      Top             =   45
      Width           =   2940
   End
   Begin VB.Label Label1 
      Caption         =   "Images:"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   1905
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnopen 
         Caption         =   "&Open BMP Pack"
      End
      Begin VB.Menu mnsave 
         Caption         =   "&Save BMP Pack"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "How to use:" + vbCr + _
       "Choose menu File/Open BMP Pack" + vbCr + _
       "Choose a bpk file."
End Sub

Private Sub List1_Click()
If List1.ListIndex < 0 Then Exit Sub
'Picture1.Cls
Picture1.Height = MyBMP(List1.ListIndex).Hei + 2
Picture1.Width = MyBMP(List1.ListIndex).Wid + 2

BitBlt Picture1.hdc, 0, 0, MyBMP(List1.ListIndex).Wid, MyBMP(List1.ListIndex).Hei, MyBMP(List1.ListIndex).GFX, 0, 0, vbSrcCopy
Picture1.Refresh
Label2.Caption = "Picture Size:" & MyBMP(List1.ListIndex).Wid & "x" & MyBMP(List1.ListIndex).Hei & vbCr & "BitPerPixel : " & MyBMP(List1.ListIndex).BPP & vbCr & "Image File Size : " & MyBMP(List1.ListIndex).FSize

End Sub

Private Sub mnopen_Click()
OpenBMPPack CD(Me.hWnd, "BMP Pack Files(*.bpk)" + Chr(0) + "*.bpk", "Open BMP Pack", "*.bmp"), Form1.List1
End Sub

Private Sub mnsave_Click()
SaveBMPPack CD(Me.hWnd, "BMP Pack Files(*.bpk)" + Chr(0) + "*.bpk", "Save BMP Pack", "*.bmp"), Form1.List1
End Sub
