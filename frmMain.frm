VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Parthenon"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   2220
   ClientWidth     =   11760
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbxInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   2520
      ScaleHeight     =   5745
      ScaleWidth      =   6945
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox tbxInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmMain.frx":0000
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   6720
         TabIndex        =   2
         Top             =   0
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.PictureBox pbxParthenon 
      AutoSize        =   -1  'True
      Height          =   6060
      Left            =   2400
      Picture         =   "frmMain.frx":0006
      ScaleHeight     =   6000
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      Begin VB.PictureBox pbxFloorPlan 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   6945
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   6975
         Begin VB.PictureBox Image1 
            AutoSize        =   -1  'True
            Height          =   4185
            Left            =   60
            Picture         =   "frmMain.frx":15159
            ScaleHeight     =   4125
            ScaleWidth      =   6750
            TabIndex        =   7
            Top             =   360
            Width           =   6810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   6720
            TabIndex        =   6
            Top             =   0
            Width           =   105
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   6975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   6720
            TabIndex        =   5
            Top             =   0
            Width           =   105
         End
      End
   End
   Begin VB.Timer tmrMenu 
      Interval        =   1
      Left            =   10440
      Top             =   6000
   End
   Begin VB.PictureBox pbxMenuHolder 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   2055
      TabIndex        =   8
      Top             =   720
      Width           =   2055
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   23
         Top             =   0
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "History"
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
            Left            =   600
            TabIndex        =   24
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   21
         Top             =   480
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Architecture"
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
            TabIndex        =   22
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   19
         Top             =   960
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Floor Plan"
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
            TabIndex        =   20
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth of Athena"
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
            TabIndex        =   18
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   2025
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fast Facts on Athena"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   0
            Width           =   1320
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   13
         Top             =   3240
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Slide show"
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
            Left            =   240
            TabIndex        =   14
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   11
         Top             =   4560
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
            Index           =   6
            Left            =   240
            TabIndex        =   12
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.PictureBox pbxMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   9
         Top             =   4080
         Width           =   2055
         Begin VB.Label lblMenu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credits"
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
            Index           =   7
            Left            =   240
            TabIndex        =   10
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2040
         Y1              =   3840
         Y2              =   3840
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MenuClick(Index As Integer)
 Select Case Index
  Case 0
    tbxInfo = "When The Anthenian Empire was at the peak of its power, it began work on the " & _
              "Parthenon in c. 447 B.C. and lasted until ca. 432 B.C. (15 years). The Parthenon " & _
              "was built over an existing temple called The Pre-Parthenon, which was constructed before " & _
              "The Persian War, and destroyed by the Pursians in ca. 480 B.C. The Parthenon was designed by " & _
              "Iktinos, Kallikrates, Pheidias. It was dedicated to Athena Parthenos (""The Virgin""), hence the " & _
              "the name Parthenon. It is thought that the western cella were stengthened with bronze bars, and that " & _
              "it was used as a treasury. It appears that the Parthenon was damaged by fire, but the exact date " & _
              "is still be debated. The most likely suggestions range from 150 B.C. to 267 A.D. during the invasion of " & _
              "the Herulians. The Parthenon was converted to Christian Church in ca. 600 A.D. and in 1687 as small " & _
              "mosque was built in the cella."
    pbxInfo.Visible = True
    pbxFloorPlan.Visible = False
  Case 1
    tbxInfo = "The Parthenon is a Doric, Peripteral Temple, meaning that it has a rectangular floor with a series of low " & _
              "step on each side, and a colonnade (8x17) of Doric columns extending around the entire structure. Every entrance " & _
              "has an additional six columns in front of it. Naos, the largest of the two interior rooms contains the cult statue " & _
              "Opisthomodos, the smallest of the two, was used as a treasury."
    pbxInfo.Visible = True
    pbxFloorPlan.Visible = False
  Case 2
    pbxFloorPlan.Visible = True
    pbxInfo.Visible = False
  Case 3
    tbxInfo = "When Zues swallowed his wife, Metis, she was pregnant with a child. Later Zues was tortured by an intolerable head-ache. " & _
              "To cure him, Hephaestus split open his head with a bronze axe and from the gaping wound sprang Athena, Fully armed and " & _
              "Brandishing a javelin. All of the immortals were struct with astonishment and filled with awe. Athena became Zues' favorite child. " & _
              "His preference for her was marked and his undulgence towards her was so extreme that it aroused the jealousy of the other gods."
    pbxInfo.Visible = True
    pbxFloorPlan.Visible = False
  Case 4
    tbxInfo = "Symbol/Attribute: Owl, Signifying watchfullness and wisdom." & vbNewLine & vbNewLine & _
              "Strengths: Rational, Intelligent, A powerfull defender in war, but also a potent peace maker." & vbNewLine & vbNewLine & _
              "Weaknesses: Reason rules her, she is not usually emotional or compassionate." & vbNewLine & vbNewLine & _
              "Birth place: From the forehead of her father Zues." & vbNewLine & vbNewLine & _
              "Spouse: none." & vbNewLine & vbNewLine & _
              "Children: none." & vbNewLine & vbNewLine & _
              "Interesting fact: One of her epithets (titles) is ""Grey-eyed"". Her gift to the Greeks was the useful olive tree. The underside of the olive tree's leaf is grey, and when the wind lifts the leaves, it shows Athena's many ""eyes""."
    pbxInfo.Visible = True
    pbxFloorPlan.Visible = False
  Case 5
    pbxInfo.Visible = False
    pbxFloorPlan.Visible = False
    frmSlide.Show
    Me.Visible = False
  Case 6
    Unload Me
  Case 7
    pbxInfo.Visible = False
    pbxFloorPlan.Visible = False
    frmCredits.Show
    Me.Visible = False
 End Select
End Sub

Private Sub Form_Load()
 For i = 0 To (lblMenu.Count - 1)
  With lblMenu(i)
   .Left = (pbxMenu(i).ScaleWidth - .Width) / 2
   .Top = (pbxMenu(i).ScaleHeight - .Height) / 2
  End With
 Next i
 SetTopMost Me.hwnd
 With pbxMenuHolder
  .Left = 120
  .Top = 120
 End With
 With pbxParthenon
  .Left = pbxMenuHolder.Left + pbxMenuHolder.Width + 120
  .Top = 120
 End With
 With Me
  .Width = pbxParthenon.Left + pbxParthenon.Width + 120 + Width - ScaleWidth
  .Height = pbxParthenon.Top + pbxParthenon.Height + 120 + Height - ScaleHeight
 End With
End Sub

Private Sub Label1_Click()
 pbxInfo.Visible = False
End Sub

Private Sub Label3_Click()
 pbxFloorPlan.Visible = False
End Sub

Private Sub pbxFloorPlan_Click()
 pbxFloorPlan.Visible = False
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

Private Sub tmrMenu_Timer()
 For i = 0 To (lblMenu.Count - 1)
    If MouseOver(pbxMenu(i).hwnd) Then
        lblMenu(i).ForeColor = &HFF00&
    Else
        lblMenu(i).ForeColor = &HFFFFFF
    End If
 Next i
End Sub
