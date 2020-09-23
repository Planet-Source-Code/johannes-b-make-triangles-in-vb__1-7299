VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "RND"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Backcolor (CLS)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Fore color"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1320
      Width           =   255
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   720
      Width           =   255
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   120
      Width           =   255
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   30
      Left            =   360
      Max             =   255
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   30
      Left            =   360
      Max             =   255
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   30
      Left            =   360
      Max             =   255
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
If Option1.Value = True Then
Form1.Picture2.BackColor = Form2.BackColor
Form1.Picture1.ForeColor = Form2.BackColor
End If

If Option2.Value = True Then
Form1.Picture3.BackColor = Form2.BackColor
Form1.Picture1.BackColor = Form2.BackColor
End If

End Sub

Private Sub Command2_Click()
HScroll1.Value = Rnd * 255
HScroll2.Value = Rnd * 255
HScroll3.Value = Rnd * 255
End Sub


Private Sub HScroll1_Change()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture1.BackColor = RGB(HScroll1.Value, 0, 0)
End Sub

Private Sub HScroll1_Scroll()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture1.BackColor = RGB(HScroll1.Value, 0, 0)
End Sub

Private Sub HScroll2_Change()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture2.BackColor = RGB(0, HScroll2.Value, 0)
End Sub


Private Sub HScroll2_Scroll()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture2.BackColor = RGB(0, HScroll2.Value, 0)
End Sub


Private Sub HScroll3_Change()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture3.BackColor = RGB(0, 0, HScroll3.Value)
End Sub


Private Sub HScroll3_Scroll()
Form2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Picture3.BackColor = RGB(0, 0, HScroll3.Value)
End Sub


