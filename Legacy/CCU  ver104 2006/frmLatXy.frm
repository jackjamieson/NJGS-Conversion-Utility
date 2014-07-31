VERSION 5.00
Begin VB.Form frmLatXy 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Latitude / Longitude <=> NJ State Plane"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLatXy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInputB 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   2280
      TabIndex        =   32
      Top             =   2982
      Width           =   1455
   End
   Begin VB.TextBox txtInputA 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   600
      TabIndex        =   31
      Top             =   2982
      Width           =   1455
   End
   Begin VB.Frame fraDatum 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Input Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NAD83"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NAD27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraUnits 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lat / Lon Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      TabIndex        =   19
      Top             =   1200
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ddmmss.ss"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "decimal degrees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame fraCoordinate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Input Coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   18
      Top             =   1200
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NJ State Plane"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Lat / Lon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
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
      TabIndex        =   6
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
      TabIndex        =   7
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
      TabIndex        =   8
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
      TabIndex        =   9
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
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblUnitsA 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "U.S. survey feet"
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   4080
      TabIndex        =   33
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblQuad 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "USGS Quadrangle - "
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   600
      TabIndex        =   30
      Top             =   5883
      Width           =   2115
   End
   Begin VB.Label lblOutputF 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2760
      TabIndex        =   29
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Label lblUnitsC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ddmmss.ss (NAD27)"
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   4080
      TabIndex        =   28
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label lblOutputD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   600
      TabIndex        =   27
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblOutputA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   600
      TabIndex        =   26
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblOutputE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   2280
      TabIndex        =   25
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblYcoorC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Longitude"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2475
      TabIndex        =   24
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label lblXcoorC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Latitude"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   900
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblAscCoor 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Atlas Sheet Coordinate"
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   4080
      TabIndex        =   21
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lblOutputC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   2280
      TabIndex        =   20
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblXcoorB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Latitude"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   900
      TabIndex        =   17
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblYcoorB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Longitude"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2475
      TabIndex        =   16
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label lblUnitsB 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ddmmss.ss (NAD83)"
      ForeColor       =   &H00C00000&
      Height          =   325
      Left            =   4080
      TabIndex        =   15
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label lblOutputB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   2280
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblXcoorA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Northing"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   870
      TabIndex        =   13
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label lblYCoorA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Easting"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2595
      TabIndex        =   12
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label lblTitle3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Latitude / Longitude <=> NJ State Plane"
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
      Left            =   1733
      TabIndex        =   11
      Top             =   240
      Width           =   5535
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
         Begin VB.Menu mnuFileGotoStatepalemodule 
            Caption         =   "State plane module"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFileGotoDecimaldegreesmodule 
            Caption         =   "Decimal degrees module"
         End
         Begin VB.Menu mnuFileGotoAtlassheetcoordinatemodule 
            Caption         =   "Atlas sheet coordinate module"
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
Attribute VB_Name = "frmLatXy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim intCount As Integer

Private Sub cmdCalculate_Click()

blnBatch = False

On Error GoTo Err_cmdCalculate_Click

Dim intCount As Integer
    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0      'state plane to latitude/longitude
                    If optPick(2).Value = True Then     'output in ddmmss.ss
                        dblNorthing = txtInputA.Text
                        dblEasting = txtInputB.Text
                        
                        If dblNorthing >= -73899.055 And dblNorthing <= 927030.312 Then
                            If dblEasting >= 170739.055 And dblEasting <= 732257.404 Then
                                'convert state plane coordinates to latitude/longitude NAD83
                                Call xytoll83

                                'convert decimal degrees to degrees, minutes and decimal seconds
                                Call DecimalDegrees
                                lblOutputA.Caption = dblLatdd
                                lblOutputB.Caption = dblLondd
                                
                                'calculate the atlas sheet coordinate
                                Call LatToAt
                                lblOutputC.Caption = strAscCoor
                                
                                'convert the latitude/longitude NAD83 to NAD27
                                Call Datum
                                dblLatdecimal = dblLatDatum
                                dblLondecimal = dblLonDatum
                                
                                'convert decimal degrees to degrees, minutes and decimal seconds
                                Call DecimalDegrees
                                lblOutputD.Caption = dblLatdd
                                lblOutputE.Caption = dblLondd

                             Else
                                GoTo Err_cmdCalculate_Click
                            End If
                        Else
                            GoTo Err_cmdCalculate_Click
                        End If
                    Else                                'output in decimal degrees
                        dblNorthing = txtInputA.Text
                        dblEasting = txtInputB.Text
                        
                        If dblNorthing >= -73899.055 And dblNorthing <= 927030.312 Then
                            If dblEasting >= 170739.055 And dblEasting <= 732257.404 Then
                            
                                'convert state plane coordinates to latitude/longitude NAD83
                                Call xytoll83
                                lblOutputA.Caption = Round(dblLatdecimal, 6)
                                lblOutputB.Caption = Round(dblLondecimal, 6)
                                
                                'convert decimal degrees to degrees, minutes and decimal
                                Call DecimalDegrees
                                
                                'calculate the atlas sheet coordinate
                                Call LatToAt
                                lblOutputC.Caption = strAscCoor
                                
                                'convert the latitude/longitude NAD83 to NAD27
                                Call Datum
                                lblOutputD.Caption = Round(dblLatDatum, 6)
                                lblOutputE.Caption = Round(dblLonDatum, 6)
                                
                            Else
                                GoTo Err_cmdCalculate_Click
                            End If
                        Else
                            GoTo Err_cmdCalculate_Click
                        End If
                    End If

                Case 1      'latitude/longitude to state plane
                    If optPick(2).Value = True Then     'input units ddmmss.ss
                        dblLatdd = txtInputA.Text
                        dblLondd = txtInputB.Text

                        'disassemble dblLatdd and check for valid latitude
                        If dblLatdd < 383730 Or dblLatdd > 412230 Then
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
                        If dblLondd > 753730 Or dblLondd < 733730 Then
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

                        'calculate the atlas sheet coordinate
                        Call LatToAt
                        lblOutputC.Caption = strAscCoor
                                
                        If optPick(5).Value = True Then
                                
                            'convert degrees, minutes and decimal seconds to decimal degrees
                            Call Ddmmss
                    
                            'convert latitude/longitude to state plane coordinates
                            Call lltoxy83
                            lblOutputA.Caption = Round(dblNorthing, 1)
                            lblOutputB.Caption = Round(dblEasting, 1)
                                    
                        End If
                                
                        'convert the latitude/longitude NAD27 <=> NAD83
                        Call Datum
                        dblLatdecimal = dblLatDatum
                        dblLondecimal = dblLonDatum
                                
                        'convert decimal degrees to degrees, minutes and decimal seconds
                        Call DecimalDegrees
                        lblOutputD.Caption = dblLatdd
                        lblOutputE.Caption = dblLondd
                                
                        If optPick(4).Value = True Then
                                
                            'convert latitude/longitude to state plane coordinates
                            Call lltoxy83
                            lblOutputA.Caption = Round(dblNorthing, 1)
                            lblOutputB.Caption = Round(dblEasting, 1)
                                    
                        End If
                                
                    Else
                        'input units decimal degrees
                        dblLatdecimal = txtInputA.Text
                        dblLondecimal = txtInputB.Text
                        
                        If dblLatdecimal >= 38.625 And dblLatdecimal <= 41.375 Then
                            If dblLondecimal <= 75.625 And dblLondecimal >= 73.625 Then
                            
                                If optPick(5).Value = True Then
                                                  
                                    'convert latitude/longitude to state plane coordinates
                                    Call lltoxy83
                                    lblOutputA.Caption = Round(dblNorthing, 0)
                                    lblOutputB.Caption = Round(dblEasting, 0)
                                    
                                End If
                                
                                'convert decimal degrees to degrees, minutes and decimal seconds
                                Call DecimalDegrees
                                
                                'disassemble dblLatdd
                                dblLatDeg = Left(dblLatdd, 2)
                                dblLatMin = Mid(dblLatdd, 3, 2)
                                dblLatSec = Mid(dblLatdd, 5, 2)
                                
                                'disassemble dblLondd
                                dblLonDeg = Left(dblLondd, 2)
                                dblLonMin = Mid(dblLondd, 3, 2)
                                dblLonSec = Mid(dblLondd, 5, 2)
                                
                                'calculate the atlas sheet coordinate
                                Call LatToAt
                                lblOutputC.Caption = strAscCoor
                                
                                'convert the latitude/longitude NAD83 <=> NAD27
                                Call Datum
                                dblLatdecimal = dblLatDatum
                                dblLondecimal = dblLonDatum
                                lblOutputD.Caption = Round(dblLatDatum, 6)
                                lblOutputE.Caption = Round(dblLonDatum, 6)
                                
                                If optPick(4).Value = True Then
                                                  
                                    'convert latitude/longitude to state plane coordinates
                                    Call lltoxy83
                                    lblOutputA.Caption = Round(dblNorthing, 0)
                                    lblOutputB.Caption = Round(dblEasting, 0)
                                                                        
                                End If
                                
                            Else
                                GoTo Err_cmdCalculate_Click
                            End If
                        Else
                            GoTo Err_cmdCalculate_Click
                        End If
                    End If
            End Select
        End If
    Next intCount
lblOutputF.Caption = strQnam
txtInputA.SetFocus
Exit Sub
    
Err_cmdCalculate_Click:
    MsgBox prompt:="The coordinate entered is not valid!"
    lblOutputA.Caption = ""
    lblOutputB.Caption = ""
    lblOutputC.Caption = ""
    lblOutputD.Caption = ""
    lblOutputE.Caption = ""
    lblOutputF.Caption = ""
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
    lblOutputC.Caption = ""
    lblOutputD.Caption = ""
    lblOutputE.Caption = ""
    lblOutputF.Caption = ""
    txtInputA.SetFocus
End Sub

Private Sub cmdReturn_Click()
    frmMain.Show
    Unload frmLatXy

End Sub

Private Sub Form_Load()
    frmLatXy.Top = (Screen.Height - frmLatXy.Height) / 2
    frmLatXy.Left = (Screen.Width - frmLatXy.Width) / 2
   
End Sub

Private Sub mnuFileCalculate_Click()

End Sub

Private Sub imgDepLogo_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        End
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileGotoAtlassheetcoordinatemodule_Click()
    frmLatToAt.Show
    Unload frmLatXy
End Sub

Private Sub mnuFileGotoBatchconversionmodule_Click()
    frmBatch.Show
    Unload frmLatXy
End Sub

Private Sub mnuFileGotoDecimaldegreesmodule_Click()
    frmDegrees.Show
    Unload frmLatXy
End Sub

Private Sub mnuFileGotoMainpage_Click()
    frmMain.Show
    Unload frmLatXy
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

    For intCount = 0 To 1
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    blnNAD83 = False
                    fraDatum.Visible = False
                    If optPick(2).Value = True Then
                        lblUnitsA.Caption = "state plane feet"
                        lblUnitsB.Caption = "ddmmss.ss (NAD83)"
                        lblUnitsC.Caption = "ddmmss.ss (NAD27)"
                        lblXcoorA.Caption = "Northing"
                        lblYCoorA.Caption = "Easting"
                        lblXcoorB.Caption = "Latitude"
                        lblYcoorB.Caption = "Longitude"
                    Else
                        lblUnitsA.Caption = "state plane feet"
                        lblUnitsB.Caption = "decimal degrees (NAD83)"
                        lblUnitsC.Caption = "decimal degrees (NAD27)"
                        lblXcoorA.Caption = "Northing"
                        lblYCoorA.Caption = "Easting"
                        lblXcoorB.Caption = "Latitude"
                        lblYcoorB.Caption = "Longitude"
                    End If
                Case 1
                    fraDatum.Visible = True
                    If optPick(2).Value = True Then
                        If optPick(4).Value = True Then
                            blnNAD83 = True
                            lblUnitsA.Caption = "ddmmss.ss (NAD27)"
                            lblUnitsC.Caption = "ddmmss.ss (NAD83)"
                        Else
                            blnNAD83 = False
                            lblUnitsA.Caption = "ddmmss.ss (NAD83)"
                            lblUnitsC.Caption = "ddmmss.ss (NAD27)"
                        End If
                        lblUnitsB.Caption = "state plane feet"
                        lblXcoorA.Caption = "Latitude"
                        lblYCoorA.Caption = "Longitude"
                        lblXcoorB.Caption = "Northing"
                        lblYcoorB.Caption = "Easting"
                    Else
                        If optPick(4).Value = True Then
                            blnNAD83 = True
                            lblUnitsA.Caption = "decimal degrees (NAD27)"
                            lblUnitsC.Caption = "decimal degrees (NAD83)"
                        Else
                            blnNAD83 = False
                            lblUnitsA.Caption = "decimal degrees (NAD83)"
                            lblUnitsC.Caption = "decimal degrees (NAD27)"
                        End If
                        lblUnitsB.Caption = "state plane feet"
                        lblXcoorA.Caption = "Latitude"
                        lblYCoorA.Caption = "Longitude"
                        lblXcoorB.Caption = "Northing"
                        lblYcoorB.Caption = "Easting"
                    End If
            End Select
        End If
    Next intCount
    
txtInputA.Text = ""
txtInputB.Text = ""
lblOutputA.Caption = ""
lblOutputB.Caption = ""
lblOutputC.Caption = ""
lblOutputD.Caption = ""
lblOutputE.Caption = ""
lblOutputF.Caption = ""
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
