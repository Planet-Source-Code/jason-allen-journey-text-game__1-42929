VERSION 5.00
Begin VB.Form frmJourney 
   Caption         =   "Jouney"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmJourney.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help!"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtActionPane 
      BackColor       =   &H80000004&
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5400
      Width           =   4575
   End
   Begin VB.TextBox txtAction 
      Height          =   375
      Left            =   240
      MaxLength       =   25
      TabIndex        =   9
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Game Status"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.Label lblDirection 
         Caption         =   "North, East, South, West, Down, Up"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label lblTravel 
         Caption         =   "You can travel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblCarrying 
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCarry 
         Caption         =   "You are carrying:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblItems 
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItemList 
         Caption         =   "Items in this area:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblWhereAreYou 
         Caption         =   "You are in..."
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblCommand 
      Caption         =   "What do you want to do?"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuFileLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHow 
         Caption         =   "&How To Play"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Journey"
      End
   End
End
Attribute VB_Name = "frmJourney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHelp_Click()
Call ShowHelpBox
End Sub

Private Sub cmdLoad_Click()
Call OpenFile
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Call SaveFile
End Sub

Private Sub Form_Load()
    Call MapSetup
    Call DefaultInventory
    
    myRs.MoveFirst
    
    Call UpdateCaptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Answer
Answer = MsgBox("Are you sure you want to Exit?", vbYesNo, "Confirm Exit")
If Answer = vbNo Then
    Cancel = True
    Exit Sub
End If
End
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileLoad_Click()
Call OpenFile
End Sub

Private Sub mnuFileSave_Click()
Call SaveFile
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuHelpHow_Click()
Call ShowHelpBox
End Sub

Private Sub txtAction_KeyDown(KeyCode As Integer, Shift As Integer)
'reset result
Result = ""
'if enter key pressed, check for text, call action function, and display final result in pane.
If KeyCode = 13 Then
    If txtAction.Text = "" Then KeyCode = 0: Exit Sub
    Call Game(txtAction.Text)
    If Result <> "" Then txtActionPane.Text = "    " & Result & vbCrLf & txtActionPane.Text
    txtAction.Text = ""
    KeyCode = 0
    'keep actionpane caption length under 1000 characters to conserve memory
    If Len(txtActionPane.Text) > 1000 Then txtActionPane.Text = Left(txtActionPane.Text, 900)
End If
End Sub
