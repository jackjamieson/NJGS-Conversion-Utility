VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLatToAt 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Latitude/Longitude to Atlas Sheet Coordinate"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9000
   Icon            =   "frmLatToAt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnits 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lat / Lon Units"
      Height          =   1215
      Left            =   2640
      TabIndex        =   23
      Top             =   1200
      Width           =   1935
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "decimal degrees"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ddmmss"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fraSelect 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Lat/Lon to ASC"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ASC to Lat/Lon"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.TextBox txtInputA 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Width           =   1215
   End
   Begin VB.TextBox txtInputB 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Width           =   1215
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
   Begin MSMask.MaskEdBox mskInputC 
      Height          =   325
      Left            =   1560
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##:###"
      PromptChar      =   "_"
   End
   Begin VB.Label lblOutputF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   2760
      TabIndex        =   27
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label lblQuad 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "USGS Quadrangle - "
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
      Left            =   600
      TabIndex        =   26
      Top             =   5760
      Width           =   2115
   End
   Begin VB.Label lblUnitsB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ddmmss.ss (NAD27)"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblUnitsA 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ddmmss.ss (NAD27)"
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
      Left            =   3840
      TabIndex        =   19
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblXcoorB 
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
      Left            =   780
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblYcoorB 
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
      Left            =   2355
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblAscCoorA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Atlas Sheet Coordinate"
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
      Visible         =   0   'False
      Width           =   2520
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
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
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
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblTitle3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Latitude / Longitude <=> Atlas Sheet Coordinate"
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
      Left            =   1155
      TabIndex        =   11
      Top             =   240
      Width           =   6690
   End
   Begin VB.Label lblOutputC 
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
      Left            =   1200
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblXcoorA 
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
      Left            =   780
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblYcoorA 
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
      Left            =   2355
      TabIndex        =   8
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label lblAscCoorB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Atlas Sheet Coordinate"
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
      TabIndex        =   7
      Top             =   4560
      Width           =   2520
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileGoto 
         Caption         =   "Goto"
         Begin VB.Menu mnuFileGotoMainpage 
            Caption         =   "Main menu"
         End
         Begin VB.Menu mnuFileGotoStateplanemodule 
            Caption         =   "State plane module"
         End
         Begin VB.Menu mnuFileGotoDecimaldegreesmodule 
            Caption         =   "Decimal degrees module"
         End
         Begin VB.Menu mnuFileGotoAtlassheetcoordinatesmodule 
            Caption         =   "Atlas sheet coordinate module"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFileGotoBatchconversionmodule 
            Caption         =   "Batch conversion module"
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
Attribute VB_Name = "frmLatToAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intLatDeg As Integer, intLatMin As Integer, intLatSec As Integer
    Dim intLonDeg As Integer, intLonMin As Integer, intLonSec As Integer
    Dim sngSouth As Single, sngEast As Single
    Dim intBlkSouth As Integer, intBlkEast As Integer, intBlk As Integer
    Dim intR1 As Integer, intR2 As Integer, intR3 As Integer
    Dim strSheet As String, strBlk As String, strRect As String
    Dim intCount As Integer

Private Sub cmdCalculate_Click()

blnBatch = False

Dim intCount As Integer
    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0  'convert latitude/longitude to atlas sheet coordinates
                    On Error GoTo Err_cmdCalculate_Click
                    If optPick(2).Value = True Then 'input in ddmmss.ss
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
                            
                        dblLatSec = Mid(dblLatdd, 5, 6)
                        If Int(dblLatSec) > 59 Then
                            GoTo Err_cmdCalculate_Click
                        End If
                        
                        'disassemble dblLondd and check for valid longitude
                        If dblLondd > 1800000 Then
                            GoTo Err_cmdCalculate_Click
                        End If
                        
                        If dblLondd > 999999 Then
                            dblLonDeg = Left(dblLondd, 3)
                            dblLonMin = Mid(dblLondd, 4, 2)
                            dblLonSec = Mid(dblLondd, 6, 6)
                        Else
                            dblLonDeg = Left(dblLondd, 2)
                            dblLonMin = Mid(dblLondd, 3, 2)
                            dblLonSec = Mid(dblLondd, 5, 6)
                        End If
                        
                        If dblLonMin > 59 Then
                            GoTo Err_cmdCalculate_Click
                        End If
    
                        If Int(dblLonSec) > 59 Then
                            GoTo Err_cmdCalculate_Click
                        End If
                        
                            Call LatToAt
                            
                            'display results
                            lblOutputC.Caption = strAscCoor
                        Else    'input in decimal degrees
                            dblLatdecimal = txtInputA.Text
                            dblLondecimal = txtInputB.Text
                            
                            'check for valid latitude and longitude
                            If dblLatdecimal > 90 Or dblLondecimal > 180 Then
                                GoTo Err_cmdCalculate_Click
                            End If
                            
                            Call DecimalDegrees
                            Call LatToAt
                            
                            'display results
                            lblOutputC.Caption = strAscCoor
                        End If
    
                    txtInputA.SelStart = 0
                    txtInputA.SelLength = Len(txtInputA.Text)
                    txtInputA.SetFocus

                Case 1  'convert atlas sheet coordinates to latitude/longitude
                    On Error GoTo Err1_cmdCalculate_Click
                    strAscCoor = mskInputC
                    
                    
                    If optPick(2).Value = True Then
                        Call AtToLat
                        lblOutputA = Round(dblLatdd, 0)
                        lblOutputB = Round(dblLondd, 0)
                    Else
                        Call AtToLat
                        Call Ddmmss
                        lblOutputA = Round(dblLatdecimal, 3)
                        lblOutputB = Round(dblLondecimal, 3)
                    End If
                    
                    mskInputC.SelStart = 0
                    mskInputC.SelLength = Len(mskInputC.Text)
                    mskInputC.SetFocus
                    
            End Select
        End If
    Next intCount

On Error GoTo Err2_cmdCalculate_Click
Open "NAD27.dat" For Input As #1
    
Do While Not EOF(1)
    Input #1, strQnum, strQnam, dblNwlat, dblNwlon, dblNw, dblS3, dblNelat, dblNelon, _
              dblNe, dblS4, dblSwlat, dblSwlon, dblSw, dblS1, dblSelat, dblSelon, _
              dblSe, dblS2
        If dblLatdd <= dblNwlat And dblLatdd >= dblSelat And _
        dblLondd <= dblNwlon And dblLondd >= dblSelon Then
            Exit Do
        End If
Loop
Close #1
    
lblOutputF.Caption = strQnam

Exit Sub

Err_cmdCalculate_Click:
    MsgBox prompt:="The coordinate entered is not valid!"
    lblOutputC.Caption = ""
    txtInputA.SetFocus
    Exit Sub
    
Err1_cmdCalculate_Click:
    MsgBox prompt:="The coordinate entered is not valid!"
    lblOutputA.Caption = ""
    lblOutputB.Caption = ""
    mskInputC.SetFocus
    Exit Sub
    
Err2_cmdCalculate_Click:
    MsgBox prompt:="The file NAD27.dat is missing!"
    lblOutputF.Caption = "file NAD27.dat missing!"
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    PrintForm
End Sub

Private Sub cmdReset_Click()
    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    txtInputA.Text = ""
                    txtInputB.Text = ""
                    lblOutputC.Caption = ""
                    lblOutputF.Caption = ""
                    txtInputA.SetFocus
                    
                Case 1
                    mskInputC.Mask = ""
                    mskInputC.Text = ""
                    mskInputC.Mask = "##:##:###"
                    lblOutputA.Caption = ""
                    lblOutputB.Caption = ""
                    lblOutputF.Caption = ""
                    mskInputC.SetFocus
            End Select
        End If
    Next intCount
    
End Sub

Private Sub cmdReturn_Click()
    frmMain.Show
    Unload frmLatToAt
End Sub

Private Sub Form_Load()
    frmLatToAt.Top = (Screen.Height - frmLatToAt.Height) / 2
    frmLatToAt.Left = (Screen.Width - frmLatToAt.Width) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        End
    End If
End Sub


Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileGotoBatchconversionmodule_Click()
    frmBatch.Show
    Unload frmLatToAt
End Sub

Private Sub mnuFileGotoDecimaldegreesmodule_Click()
    frmDegrees.Show
    Unload frmLatToAt
End Sub

Private Sub mnuFileGotoMainpage_Click()
    frmMain.Show
    Unload frmLatToAt
End Sub

Private Sub mnuFileGotoStateplanemodule_Click()
    frmLatXy.Show
    Unload frmLatToAt
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

Private Sub mskInputC_GotFocus()
    mskInputC.SelStart = 0
    mskInputC.SelLength = Len(mskInputC.Text)
End Sub

Private Sub optPick_Click(Index As Integer)

    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    lblXcoorA.Visible = True
                    lblYCoorA.Visible = True
                    lblAscCoorA.Visible = False
                    lblXcoorB.Visible = False
                    lblYcoorB.Visible = False
                    lblAscCoorB.Visible = True
                    txtInputA.Visible = True
                    txtInputB.Visible = True
                    mskInputC.Visible = False
                    lblOutputA.Visible = False
                    lblOutputB.Visible = False
                    lblOutputC.Visible = True
                    lblUnitsA.Visible = True
                    lblUnitsB.Visible = False
                    
                    txtInputA.Text = ""
                    txtInputB.Text = ""
                    lblOutputC.Caption = ""
                    lblOutputF.Caption = ""
                    txtInputA.SetFocus
                    
                    If optPick(2).Value = True Then
                        lblUnitsA.Caption = "ddmmss.ss (NAD27)"
                    Else
                        lblUnitsA.Caption = "decimal degrees (NAD27)"
                    End If
                Case 1
                    lblXcoorA.Visible = False
                    lblYCoorA.Visible = False
                    lblAscCoorA.Visible = True
                    lblXcoorB.Visible = True
                    lblYcoorB.Visible = True
                    lblAscCoorB.Visible = False
                    txtInputA.Visible = False
                    txtInputB.Visible = False
                    mskInputC.Visible = True
                    lblOutputA.Visible = True
                    lblOutputB.Visible = True
                    lblOutputC.Visible = False
                    lblUnitsA.Visible = False
                    lblUnitsB.Visible = True
                    
                    mskInputC.Mask = ""
                    mskInputC.Text = ""
                    mskInputC.Mask = "##:##:###"
                    lblOutputA.Caption = ""
                    lblOutputB.Caption = ""
                    lblOutputF.Caption = ""
                    mskInputC.SetFocus
                    
                    If optPick(2).Value = True Then
                        lblUnitsB.Caption = "ddmmss.ss (NAD27)"
                    Else
                        lblUnitsB.Caption = "decimal degrees (NAD27)"
                    End If
            End Select
        End If
    Next intCount
End Sub

Private Sub txtInputA_GotFocus()
    txtInputA.SelStart = 0
    txtInputA.SelLength = Len(txtInputA.Text)
End Sub

Private Sub txtInputB_GotFocus()
    txtInputB.SelStart = 0
    txtInputB.SelLength = Len(txtInputB.Text)
End Sub
