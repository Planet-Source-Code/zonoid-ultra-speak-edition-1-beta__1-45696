VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Text Color"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox List2 
      BackColor       =   &H8000000F&
      Height          =   2010
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Appearance"
      Height          =   1335
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2880
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Size:"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Sample Text"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim b
Dim c
Dim ret As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
  Label4.FontBold = True
Else
  Label4.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Label4.FontItalic = True
Else
  Label4.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Label4.FontUnderline = True
Else
  Label4.FontUnderline = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
  Label4.FontStrikethru = True
Else
  Label4.FontStrikethru = False
End If
End Sub

Private Sub Command1_Click()
SimpleDict.txtspeech.Font.Name = Label4.Font.Name
SimpleDict.txtspeech.Font.Size = Label4.Font.Size
SimpleDict.txtspeech.Font.Bold = Label4.Font.Bold
SimpleDict.txtspeech.Font.Italic = Label4.Font.Italic
SimpleDict.txtspeech.Font.Underline = Label4.Font.Underline
SimpleDict.txtspeech.SelColor = SimpleDict.CommonDialog1.Color
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub Command3_Click()
SimpleDict.CommonDialog1.ShowColor
Label4.ForeColor = SimpleDict.CommonDialog1.Color
End Sub

Private Sub Form_Load()
ret = frm
Debug.Print ret
For x = 1 To Screen.FontCount
List1.AddItem Screen.Fonts(x)
Next
For x = 5 To 72: List2.AddItem Str$(x): Next
For x = 0 To List1.ListCount - 1
 If SimpleDict.txtspeech.Font = List1.List(x) Then
  List1.ListIndex = x
  Label4.FontName = List1.List(x)
  Exit For
 End If
Next
For x = 0 To List2.ListCount - 1
 If Int(Val(SimpleDict.txtspeech)) = Val(List2.List(x)) Then
  List2.ListIndex = x
  Label4.FontSize = Val(List2.List(x))
  Text1.Text = List2.List(x)
  Exit For
 End If
Next
If SimpleDict.txtspeech.Font.Bold = True Then
 Label4.FontBold = True
 Check1.Value = 1
End If
If SimpleDict.txtspeech.Font.Italic = True Then
 Label4.FontItalic = True
 Check2.Value = 1
End If
If SimpleDict.txtspeech.Font.Underline = True Then
 Label4.Font.Underline = True
 Check3.Value = 1
End If
If SimpleDict.txtspeech.Font.Strikethrough = True Then
 Label4.Font.Strikethrough = True
 Check4.Value = 1
End If
End Sub
Private Sub Label3_Click()
CommonDialog1.ShowColor
Label3.BackColor = CommonDialog1.Color
Label4.ForeColor = CommonDialog1.Color
End Sub



Private Sub List1_Click()
Label4.FontName = List1.List(List1.ListIndex)
End Sub
Private Sub List2_Click()
SimpleDict.Invisible.Text = List2.List(List2.ListIndex)
Label4.FontSize = Val(SimpleDict.Invisible.Text)
End Sub
Private Sub Text1_Change()
For x = 0 To List2.ListCount - 1
If Val(Text1.Text) = Val(List2.List(x)) Then
  List2.ListIndex = x
  Label4.FontSize = Val(Text1.Text)
 Exit For
End If
Next
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then
Else
a = True
b = x
c = Y
End If
End Sub
Private Sub Image5_Click()
Me.WindowState = 1
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
a = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If a Then
Me.Move (Me.Left + x - b), (Me.Top + Y - c)
End If
End Sub
Private Sub Dm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If a Then Me.Move (Me.Left + x - b), (Me.Top + Y - c)
End Sub
Private Sub Dm_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
a = False
End Sub
