VERSION 5.00
Begin VB.Form frmSlide 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "The Parthenon"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   28
      Top             =   6480
      Width           =   615
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3480
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   26
      Top             =   6480
      Width           =   615
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox tbxTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Text            =   "5"
      Top             =   6480
      Width           =   855
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   0
      Left            =   2520
      Picture         =   "frmSlide.frx":0000
      ScaleHeight     =   5700
      ScaleWidth      =   3615
      TabIndex        =   9
      Top             =   240
      Width           =   3645
   End
   Begin VB.Timer tmrShow 
      Interval        =   5000
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   5880
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Timer tmrMenu 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   7320
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4110
      Index           =   15
      Left            =   600
      Picture         =   "frmSlide.frx":106C8
      ScaleHeight     =   4080
      ScaleWidth      =   7500
      TabIndex        =   24
      Top             =   1320
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   14
      Left            =   960
      Picture         =   "frmSlide.frx":1B256
      ScaleHeight     =   5700
      ScaleWidth      =   7005
      TabIndex        =   23
      Top             =   480
      Width           =   7035
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   13
      Left            =   600
      Picture         =   "frmSlide.frx":26C6F
      ScaleHeight     =   5445
      ScaleWidth      =   7500
      TabIndex        =   22
      Top             =   600
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   12
      Left            =   600
      Picture         =   "frmSlide.frx":32C8F
      ScaleHeight     =   4995
      ScaleWidth      =   7500
      TabIndex        =   21
      Top             =   840
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3960
      Index           =   11
      Left            =   600
      Picture         =   "frmSlide.frx":3E1CF
      ScaleHeight     =   3930
      ScaleWidth      =   7500
      TabIndex        =   20
      Top             =   1320
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5355
      Index           =   10
      Left            =   600
      Picture         =   "frmSlide.frx":533F5
      ScaleHeight     =   5325
      ScaleWidth      =   7500
      TabIndex        =   19
      Top             =   600
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   9
      Left            =   1920
      Picture         =   "frmSlide.frx":72055
      ScaleHeight     =   5700
      ScaleWidth      =   5010
      TabIndex        =   18
      Top             =   240
      Width           =   5040
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   8
      Left            =   2520
      Picture         =   "frmSlide.frx":87A41
      ScaleHeight     =   5700
      ScaleWidth      =   3810
      TabIndex        =   17
      Top             =   240
      Width           =   3840
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   7
      Left            =   2520
      Picture         =   "frmSlide.frx":8F64B
      ScaleHeight     =   5700
      ScaleWidth      =   3810
      TabIndex        =   16
      Top             =   240
      Width           =   3840
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5130
      Index           =   6
      Left            =   600
      Picture         =   "frmSlide.frx":98215
      ScaleHeight     =   5100
      ScaleWidth      =   7500
      TabIndex        =   15
      Top             =   720
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5730
      Index           =   5
      Left            =   2400
      Picture         =   "frmSlide.frx":AA282
      ScaleHeight     =   5700
      ScaleWidth      =   3870
      TabIndex        =   14
      Top             =   240
      Width           =   3900
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4920
      Index           =   4
      Left            =   600
      Picture         =   "frmSlide.frx":B42A4
      ScaleHeight     =   4890
      ScaleWidth      =   7500
      TabIndex        =   13
      Top             =   840
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4275
      Index           =   2
      Left            =   600
      Picture         =   "frmSlide.frx":C4991
      ScaleHeight     =   4245
      ScaleWidth      =   7500
      TabIndex        =   12
      Top             =   1200
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   3
      Left            =   600
      Picture         =   "frmSlide.frx":DB687
      ScaleHeight     =   5625
      ScaleWidth      =   7500
      TabIndex        =   11
      Top             =   480
      Width           =   7530
   End
   Begin VB.PictureBox Photo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3825
      Index           =   1
      Left            =   600
      Picture         =   "frmSlide.frx":EA8CB
      ScaleHeight     =   3795
      ScaleWidth      =   7500
      TabIndex        =   10
      Top             =   1200
      Width           =   7530
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 of 4"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPic As Integer

Private Sub HidePhotos()
 For i = 0 To (Photo.Count - 1)
  Photo(i).Visible = False
 Next i
 Photo(nPic).Visible = True
 Photo(nPic).ZOrder 0
 lblStatus = (nPic + 1) & " of " & Photo.Count
 lblStatus.ZOrder 0
End Sub

Private Sub MenuClick(Index As Integer)
 Select Case Index
  Case 0
   Unload Me
  Case 1
   nPic = nPic - 1
   If nPic < 0 Then: nPic = Photo.Count - 1
   Call HidePhotos
  Case 2
   nPic = nPic + 1
   If nPic >= Photo.Count Then: nPic = 0
   Call HidePhotos
  Case 3
   If tmrShow.Enabled Then
    lblMenu(3).Caption = "Play"
   Else
    lblMenu(3).Caption = "Stop"
   End If
   tmrShow.Enabled = Not tmrShow.Enabled
  Case 4
   tbxTime = Val(tbxTime) + 1
   If Val(tbxTime) > 60 Then: tbxTime = 60
   tmrShow.Interval = Val(tbxTime) * 1000
  Case 5
   If Val(tbxTime) > 1 Then: tbxTime = Val(tbxTime) - 1
   tmrShow.Interval = Val(tbxTime) * 1000
 End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.Visible = True
End Sub

Private Sub lblMenu_Click(Index As Integer)
 Call MenuClick(Index)
End Sub

Private Sub pbxMenu_Click(Index As Integer)
 Call MenuClick(Index)
End Sub

Private Sub pbxMenu_Paint(Index As Integer)
 Dim lBorder As Long
 
 pbxMenu(Index).ScaleHeight = 255
 pbxMenu(Index).ScaleWidth = 10

 For i = 0 To 255
    pbxMenu(Index).Line (0, i)-(10, i), &HFFFFFF - RGB(i, i, i)
 Next i
 
 pbxMenu(Index).DrawWidth = 2
 If MouseOver(pbxMenu(Index).hwnd) Then: lBorder = RGB(0, 255, 0): Else: lBorder = RGB(255, 255, 255)
 pbxMenu(Index).Line (0, 0)-(10, 0), lBorder
 pbxMenu(Index).Line (0, 0)-(0, 255), lBorder
 pbxMenu(Index).Line (10, 0)-(10, 255), lBorder
 pbxMenu(Index).Line (0, 255)-(10, 255), lBorder
End Sub

Private Sub Form_Load()
 
 nPic = 0
 Call HidePhotos
 
 For i = 0 To (lblMenu.Count - 1)
  With lblMenu(i)
   .Left = (pbxMenu(i).ScaleWidth - .Width) / 2
   .Top = (pbxMenu(i).ScaleHeight - .Height) / 2
  End With
 Next i
 
 For i = 0 To Photo.Count - 1
  With Photo(i)
   .Left = (ScaleWidth - .Width) / 2
   .Top = (ScaleHeight - .Height) / 2
  End With
 Next i

 SetTopMost Me.hwnd

End Sub

Private Sub tbxTime_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
  KeyAscii = 0
 Else
  If Val(tbxTime) > 60 Then: tbxTime = 60
  tmrShow.Interval = Val(tbxTime) * 1000
 End If
End Sub

Private Sub tmrMenu_Timer()
 For i = 0 To (lblMenu.Count - 1)
    If MouseOver(pbxMenu(i).hwnd) Then
        lblMenu(i).ForeColor = &HFF00&
    Else
        lblMenu(i).ForeColor = &HFFFFFF
    End If
 Next i
End Sub

Private Sub tmrShow_Timer()
   nPic = nPic + 1
   If nPic >= Photo.Count Then
    nPic = 0
    tmrShow.Enabled = False
    lblMenu(3).Caption = "Play"
   End If
   Call HidePhotos
End Sub
