VERSION 5.00
Begin VB.Form frmDegrees 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decimal degrees <=> ddmmss.ss"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9000
   ForeColor       =   &H00000000&
   Icon            =   "frmDegrees.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInputB 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtInputA 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame fraUnits 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Input Units"
      Height          =   1215
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ddmmss.ss"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "decimal degrees"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
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
      Left            =   7440
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Re&set"
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
      Left            =   7440
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   7440
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Main menu"
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
      Left            =   7440
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
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
      Left            =   7440
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblLongitude 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Longitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2520
      TabIndex        =   16
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label lblLatitude 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   840
      TabIndex        =   15
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblOutputB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   2280
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblUnitsB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ddmmss.ss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   3960
      TabIndex        =   13
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblUnitsA 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "decimal degrees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   3960
      TabIndex        =   12
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblTitle3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Decimal degrees <=> ddmmss.ss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2168
      TabIndex        =   11
      Top             =   240
      Width           =   4665
   End
   Begin VB.Label lblOutputA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFIleGoto 
         Caption         =   "Goto"
         Begin VB.Menu mnuFileGotoMainpage 
            Caption         =   "Main menu"
         End
         Begin VB.Menu mnuFileGotoStateplanemodule 
            Caption         =   "State plane module"
         End
         Begin VB.Menu mnuFileGotoDecimaldegreesmodule 
            Caption         =   "Decimal degrees module"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFileGotoAtlassheetcoordinates 
            Caption         =   "Atlas sheet coordinates module"
         End
         Begin VB.Menu mnuFileGotoBatchconversionmodule 
            Caption         =   "Batch conversion module"
         End
      End
      Begin VB.Menu mnuExit 
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
Attribute VB_Name = "frmDegrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()

blnBatch = False

On Error GoTo Err_cmdCalculate_Click

Dim intCount As Integer
    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0  'degrees, minutes, seconds to decimal degreees
                    dblLatdecimal = txtInputA.Text
                    dblLondecimal = txtInputB.Text
                    
                    'check for valid latitude and longitude
                    If dblLatdecimal > 90 Or dblLondecimal > 180 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                    
                    'convert decimal degrees to degrees, minutes and decimal seconds
                    Call DecimalDegrees
                    
                    'display values on screen
                    lblOutputA.Caption = dblLatdd
                    lblOutputB.Caption = dblLondd
                    
                Case 1  'decimal degrees to degrees, minutes, seconds
                    dblLatdd = txtInputA.Text
                    dblLondd = txtInputB.Text
                    
                    'disassemble dblLatdd and check for valid latitude
                    If dblLatdd > 900000 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                    
                    dblLatDeg = Left(dblLatdd, 2)
                        
                    dblLatMin = Mid(dblLatdd, 3, 2)
                    If dblLatMin > 59 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                        
                    dblLatSec = Mid(dblLatdd, 5, 2)
                    If dblLatSec > 59 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                    
                    'disassemble dblLondd and check for valid longitude
                    If dblLondd > 1800000 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                    
                    If dblLondd > 999999 Then
                        dblLonDeg = Left(dblLondd, 3)
                        dblLonMin = Mid(dblLondd, 4, 2)
                        dblLonSec = Mid(dblLondd, 6, 2)
                    Else
                        dblLonDeg = Left(dblLondd, 2)
                        dblLonMin = Mid(dblLondd, 3, 2)
                        dblLonSec = Mid(dblLondd, 5, 2)
                    End If
                    
                    If dblLonMin > 59 Then
                        GoTo Err_cmdCalculate_Click
                    End If

                    If dblLonSec > 59 Then
                        GoTo Err_cmdCalculate_Click
                    End If
                        
                    'convert degrees, minutes and decimal seconds to decimal degrees
                    Call Ddmmss
                    
                    'display values on screen
                    lblOutputA.Caption = Round(dblLatdecimal, 6)
                    lblOutputB.Caption = Round(dblLondecimal, 6)
                    
            End Select
        End If
    Next intCount
    txtInputA.SetFocus

Exit Sub
       
Err_cmdCalculate_Click:
    MsgBox prompt:="The coordinate entered is not valid!"
    lblOutputA.Caption = ""
    lblOutputB.Caption = ""
    txtInputA.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdExit_Click()
    End
    
End Sub


Private Sub cmdPrint_Click()
    PrintForm
    
End Sub


Private Sub cmdReset_Click()
    txtInputA.Text = ""
    txtInputB.Text = ""
    lblOutputA.Caption = ""
    lblOutputB.Caption = ""
    txtInputA.SetFocus
End Sub

Private Sub cmdReturn_Click()
    frmMain.Show
    Unload frmDegrees

End Sub

Private Sub Form_Load()
    frmDegrees.Top = (Screen.Height - frmDegrees.Height) / 2
    frmDegrees.Left = (Screen.Width - frmDegrees.Width) / 2
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        End
    End If
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFileGotoAtlassheetcoordinates_Click()
    frmLatToAt.Show
    Unload frmDegrees
End Sub

Private Sub mnuFileGotoBatchconversionmodule_Click()
    frmBatch.Show
    Unload frmDegrees
End Sub


Private Sub mnuFileGotoMainpage_Click()
    frmMain.Show
    Unload frmDegrees
End Sub

Private Sub mnuFileGotoStateplanemodule_Click()
    frmLatXy.Show
    Unload frmDegrees
End Sub

Private Sub mnuFilePrint_Click()
    PrintForm
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpTopics_Click()
    frmHelpMain.Show
End Sub

Private Sub optPick_Click(Index As Integer)
Dim intCount As Integer
    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    lblUnitsA.Caption = "decimal degrees"
                    lblUnitsB.Caption = "ddmmss.ss"
                Case 1
                    lblUnitsB.Caption = "decimal degrees"
                    lblUnitsA.Caption = "ddmmss.ss"
            End Select
        End If
    Next intCount
    
txtInputA.Text = ""
txtInputB.Text = ""
lblOutputA.Caption = ""
lblOutputB.Caption = ""
txtInputA.SetFocus

End Sub


Private Sub txtInputA_GotFocus()
    txtInputA.SelStart = 0
    txtInputA.SelLength = Len(txtInputA.Text)

End Sub


Private Sub txtInputB_GotFocus()
    txtInputB.SelStart = 0
    txtInputB.SelLength = Len(txtInputB.Text)

End Sub


