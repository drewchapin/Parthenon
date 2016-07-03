VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Credits"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroller 
      Interval        =   100
      Left            =   240
      Top             =   720
   End
   Begin VB.Timer tmrMenu 
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox pbxMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
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
   Begin VB.PictureBox pbxCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   6705
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.Label lblCredits 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   3960
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MenuClick(Index As Integer)
 Select Case Index
  Case 0
   Unload Me
 End Select
End Sub

Private Sub pbxCredits_Paint()
 With pbxCredits
  .DrawWidth = 2
  pbxCredits.Line (0, 0)-(.ScaleWidth, 0), RGB(255, 255, 255)
  pbxCredits.Line (0, 0)-(0, .ScaleHeight), RGB(255, 255, 255)
  pbxCredits.Line (.ScaleWidth, 0)-(.ScaleWidth, .ScaleHeight), RGB(255, 255, 255)
  pbxCredits.Line (0, .ScaleHeight)-(.ScaleWidth, .ScaleHeight), RGB(255, 255, 255)
 End With
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

Private Sub Form_Load()
 For i = 0 To (lblMenu.Count - 1)
  With lblMenu(i)
   .Left = (pbxMenu(i).ScaleWidth - .Width) / 2
   .Top = (pbxMenu(i).ScaleHeight - .Height) / 2
  End With
 Next i
 lblCredits = "The Parthenon" & vbNewLine & vbNewLine & _
              "By: Drew Chapin (archer282@msn.com)" & vbNewLine & vbNewLine & _
              "Made with: Visual Basic 6" & vbNewLine & vbNewLine & _
              "Made for Mr. Fowlers Arts && Humanities Class" & vbNewLine & vbNewLine & vbNewLine & _
              "Resources" & vbNewLine & _
              "-------------------------------" & vbNewLine & vbNewLine & _
              "http://www.greatbuildings.com/" & vbNewLine & _
              "http://www.perseus.tufts.edu/" & vbNewLine & _
              "http://www.goddess-athena.org/" & vbNewLine & _
              "http://www.goddessgift.com/" & vbNewLine & _
              "http://www.about.com"
 lblCredits.Left = (pbxCredits.ScaleWidth - lblCredits.Width) / 2
 SetTopMost Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.Show
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

Private Sub lblMenu_Click(Index As Integer)
 Call MenuClick(Index)
End Sub

Private Sub pbxMenu_Click(Index As Integer)
  Call MenuClick(Index)
End Sub

Private Sub tmrScroller_Timer()
 If lblCredits.Top <= lblCredits.Height * -1 Then
  lblCredits.Top = pbxCredits.ScaleHeight
 Else
  lblCredits.Top = lblCredits.Top - Screen.TwipsPerPixelY * 2
 End If
End Sub
