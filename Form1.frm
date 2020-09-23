VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Triangle by JB!"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save As..."
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLS"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   960
      Width           =   7935
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   1800
         Top             =   720
         Width           =   1575
         Visible         =   0   'False
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back color"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fore color"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Color (click to change)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Line width"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim o, p As Integer
Dim JB As Integer

Private Sub Command1_Click()
Picture1.Cls
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Width = Form1.Width - 100
Picture1.Height = Form1.Height - 1350
Label3.Caption = "Size = " & Picture1.ScaleWidth & ", " & Picture1.ScaleHeight
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

o = X
p = Y
JB = 1
Shape2.Shape = 0

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If JB = 1 Then
Shape2.Visible = True


'Picture1.Line (a, b)-(x, y)
Shape2.Left = o
Shape2.Top = p

If X > o Then Shape2.Width = X - o
If Y > p Then Shape2.Height = Y - p

If X < a Or Y < B Then
Shape2.Left = X
Shape2.Top = Y

Shape2.Width = o
Shape2.Height = p
    End If
     End If
      
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        JB = 0

If X < o Then
Shape2.Visible = False
Exit Sub
End If

If Y < p Then
Shape2.Visible = False
Exit Sub
End If

Shape2.Visible = False
If o <> 0 And p <> 0 Then Picture1.Line (Shape2.Left + Shape2.Width / 2, Shape2.Top)-(X, Y)
If o <> 0 And p <> 0 Then Picture1.Line (Shape2.Left, Shape2.Top + Shape2.Height)-(X, Y)
If o <> 0 And p <> 0 Then Picture1.Line (Shape2.Left + Shape2.Width / 2, Shape2.Top)-(X - Shape2.Width, Y)
o = 0
p = 0




End Sub

Private Sub Picture2_Click()
Form2.Show
Form2.Option1.Value = True
End Sub

Private Sub Picture3_Click()
Form2.Show
Form2.Option2.Value = True
End Sub


Private Sub Text1_Change()
Picture1.DrawWidth = Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Beep
     End If
End Sub
