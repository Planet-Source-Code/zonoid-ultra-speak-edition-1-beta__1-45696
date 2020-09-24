VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SimpleDict 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultra Speak Edition 1"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "SimpleDictFull.frx":0000
   LinkTopic       =   "SimpleDict"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Invisible 
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox txtspeech 
      Height          =   6135
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10821
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"SimpleDictFull.frx":08CA
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   375
      Left            =   3360
      OleObjectBlob   =   "SimpleDictFull.frx":094C
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR VR 
      Height          =   255
      Left            =   3120
      OleObjectBlob   =   "SimpleDictFull.frx":0970
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Read"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Font"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Mic. On"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Mic. Off"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   6705
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/18/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:10 AM"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SimpleDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Grammar As ISpeechRecoGrammar
Dim m_bRecoRunning As Boolean
Dim m_cChars As Integer
Dim WithEvents RecoContext As SpSharedRecoContext
Attribute RecoContext.VB_VarHelpID = -1
Private Sub Command1_Click()
On Error GoTo Errorbug:
 Dim sFile As String
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "*.txt"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SimpleDict.txtspeech.LoadFile sFile
    SimpleDict.Caption = sFile
Errorbug:
End Sub
Private Sub Command2_Click()
Dim sFile As String
    If SimpleDict Is Nothing Then Exit Sub
    With CommonDialog1
        .DialogTitle = "Save"
        .CancelError = False
        .Filter = ".txt"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SimpleDict.Caption = "Ultra Speak Edition 1 " + sFile
    SimpleDict.txtspeech.SaveFile sFile
End Sub
Private Sub Command3_Click()
Form2.Show
End Sub
Private Sub Command4_Click()
On Error GoTo be
TextToSpeech1.Speak txtspeech.Text
be:
End Sub

Private Sub Command5_Click()
On Error GoTo le
TextToSpeech1.StopSpeaking
le:
End Sub
Private Sub Command6_click()
txtspeech.Text = ""
SimpleDict.Caption = "Ultra Speak Edition 1"
End Sub
Private Sub Command7_Click()
CommonDialog1.CancelError = True
On Error GoTo Errhandler
Errhandler:
 On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Ultra Speak Edition 1"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If SimpleDict.txtspeech.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            SimpleDict.txtspeech.SelPrint .hDC
        End If
    End With
End Sub
Public Sub Form_Load()
On Local Error Resume Next
Call VR.Deactivate
Call VR.Activate
    SetState False
    m_cChars = 0
End Sub
Public Sub btnStart_Click()
    On Error GoTo se
    Debug.Assert Not m_bRecoRunning
    If (RecoContext Is Nothing) Then
        Debug.Print "Initializing SAPI reco context object..."
    Set RecoContext = New SpSharedRecoContext
Set Grammar = RecoContext.CreateGrammar(1)
Grammar.DictationLoad
Grammar.DictationSetState SGDSActive
    End If
    Grammar.DictationSetState SGDSActive
    SetState True
se:
End Sub
Private Sub btnStop_Click()
    Debug.Assert m_bRecoRunning
    Grammar.DictationSetState SGDSInactive
    SetState False
End Sub
Private Sub RecoContext_Recognition(ByVal StreamNumber As Long, _
                                    ByVal StreamPosition As Variant, _
                                    ByVal RecognitionType As SpeechRecognitionType, _
                                    ByVal Result As ISpeechRecoResult _
                             )
    Dim strText As String
    strText = Result.PhraseInfo.GetText
    Debug.Print "Recognition: " & strText & ", " & _
        StreamNumber & ", " & StreamPosition
    txtspeech.SelStart = m_cChars
    txtspeech.SelText = strText & " "
    m_cChars = m_cChars + 1 + Len(strText)
End Sub
Public Sub SetState(ByVal bNewState As Boolean)
    m_bRecoRunning = bNewState
    btnStart.Enabled = Not m_bRecoRunning
    btnStop.Enabled = m_bRecoRunning
End Sub
