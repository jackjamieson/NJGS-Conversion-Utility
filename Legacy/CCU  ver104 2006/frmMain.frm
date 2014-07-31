VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main menu"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame fraSelect 
      BackColor       =   &H00C0FFFF&
      Height          =   2775
      Left            =   2453
      TabIndex        =   4
      Top             =   2040
      Width           =   4095
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Batch Conversion"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   3615
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Latitude / Longitude <=> Atlas sheet coordinates"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   3855
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Latitude/Longitude <=> NJ state plane"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Decimal degrees <=> ddmmss.ss"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   3615
      End
   End
   Begin VB.Label cmdTitle2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Coordinate Conversion Utility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2753
      TabIndex        =   8
      Top             =   720
      Width           =   3495
   End
   Begin VB.Image imgDepLogo 
      Height          =   1275
      Left            =   120
      Picture         =   "frmMain.frx":4F0A
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image imgNjgsLogo 
      Height          =   1200
      Left            =   7680
      Picture         =   "frmMain.frx":568F
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lblTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "New Jersey Geological Survey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1493
      TabIndex        =   5
      Top             =   240
      Width           =   6015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRun 
         Caption         =   "Run"
         Begin VB.Menu mnuFileRunStateplane 
            Caption         =   "NJ state plane"
         End
         Begin VB.Menu mnuFileRunDecimaldegrees 
            Caption         =   "Decimal degrees"
         End
         Begin VB.Menu mnuFileRunAtlassheetcoordinates 
            Caption         =   "Atlas sheet coordinates"
         End
         Begin VB.Menu mnuRunBatchconversion 
            Caption         =   "Batch conversion"
         End
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdRun_Click()
Dim intCount As Integer
    For intCount = 0 To 3
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    frmLatXy.Show
                    Unload frmMain
                Case 1
                    frmDegrees.Show
                    Unload frmMain
                Case 2
                    frmLatToAt.Show
                    Unload frmMain
                Case 3
                    frmBatch.Show
                    Unload frmMain
            End Select
        End If
    Next intCount
    
End Sub

Private Sub Form_Load()
    frmMain.Top = (Screen.Height - frmMain.Height) / 2
    frmMain.Left = (Screen.Width - frmMain.Width) / 2

End Sub

Private Sub mnuFileRunAtlassheetcoordinates_Click()
    frmLatToAt.Show
    Unload frmMain
End Sub

Private Sub mnuFileRunDecimaldegrees_Click()
    frmDegrees.Show
    Unload frmMain
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileRunStateplane_Click()
    frmLatXy.Show
    Unload frmMain
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub


Private Sub mnuHelpDecimaldegreesmodule_Click()
    frmHelpDdm.Show
End Sub

Private Sub mnuHelpNAD27_Click()
    frmHelpNAD27.Show
End Sub

Private Sub mnuHelpTopicsAtlassheetcoordinatesmodule_Click()
    frmHelpAsc.Show
End Sub


Private Sub mnuHelpTopicsConven_Click()
    frmHelpConven.Show
End Sub

Private Sub mnuHelpTopicsNAD27_Click()
    frmHelpNAD27.Show
End Sub

Private Sub mnuHelpTopicStateplanemodule_Click()
    frmHelpSp.Show
End Sub


Private Sub mnuHelpTopics_Click()
    frmHelpMain.Show
End Sub


Private Sub mnuRunBatchconversionmodule_Click()

End Sub


Private Sub mnuRunBatchconversion_Click()
    frmBatch.Show
    Unload frmMain
End Sub


