VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBatch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Conversion"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9000
   FillColor       =   &H00404040&
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
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
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
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ASC"
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
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
         TabIndex        =   5
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRunBatch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Run &batch"
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
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   120
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files(*.*)|*.*"
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
      Height          =   1575
      Left            =   2520
      TabIndex        =   10
      Top             =   960
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
         Index           =   6
         Left            =   120
         TabIndex        =   19
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
         Index           =   5
         Left            =   120
         TabIndex        =   18
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
      Height          =   1575
      Left            =   4560
      TabIndex        =   9
      Top             =   960
      Width           =   1815
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "decimal"
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPick 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ddmmss"
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Re&set"
      Enabled         =   0   'False
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
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
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
      TabStop         =   0   'False
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
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblInputD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "                     3164855, 31:02:229                     (Atlas Sheet Coordinate ASC)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   3960
      Width           =   6735
   End
   Begin VB.Label lblInputC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Examples:       3164855, 368154.62, 312006.32   (NJ State Plane - Lat/Lon)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3720
      Width           =   6735
   End
   Begin VB.Label lblOutputA 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBatch.frx":4F0A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   13
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label lblIntputB 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBatch.frx":5097
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label lblInputA 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBatch.frx":5166
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   6735
   End
   Begin VB.Label lblTitle3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Batch Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileGoto 
         Caption         =   "Goto"
         Begin VB.Menu mnuFileGotoMainpage 
            Caption         =   "Main menu"
         End
         Begin VB.Menu mnuFileGotoStatepalemodule 
            Caption         =   "State plane module"
         End
         Begin VB.Menu mnuFileGotoDecimaldegreesmodule 
            Caption         =   "Decimal degrees module"
         End
         Begin VB.Menu mnuFileGotoAtlassheetcoordinatemodule 
            Caption         =   "Atlas sheet coordinate module"
         End
         Begin VB.Menu mnuFileGotoBatchconversionmodule 
            Caption         =   "Batch conversion module"
            Enabled         =   0   'False
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
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim intCount As Integer, intTimer As Integer
    Public strId As String, strPoints As String, strInputA As String, _
    strInputB As String, strInputC As String, strOutputA As String, strOutputB As String, _
    strOutputC As String, strOutputD As String, strOutputE As String
Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdReturn_Click()
    frmMain.Show
    Unload frmBatch

End Sub

Private Sub cmdRunBatch_Click()

blnBatch = True

dlgCommon.FileName = ""     'get file for conversion
dlgCommon.ShowOpen
strPoints = dlgCommon.FileName

If strPoints = "" Then GoTo Err_dlgCommon_Select

On Error GoTo Err_cmdRunBatch_Click

Open strPoints For Input As #3
Open "convert.dat" For Output As #4
Open "problem.dat" For Output As #5
Do While Not EOF(3)

If optPick(2).Value = True Then
    Input #3, strId, strInputC
Else
    Input #3, strId, strInputA, strInputB
End If

Dim intCount As Integer, intErr As Integer
intErr = 0
    For intCount = 0 To 2
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0      'state plane to latitude/longitude
                    If optPick(3).Value = True Then     'output in ddmmss.ss
                        dblNorthing = strInputA
                        dblEasting = strInputB
                        
                        If dblNorthing >= -73899.055 And dblNorthing <= 927030.312 Then
                            If dblEasting >= 170739.055 And dblEasting <= 732257.404 Then
                                'convert state plane coordinates to latitude/longitude NAD83
                                Call xytoll83

                                'convert decimal degrees to degrees, minutes and decimal seconds
                                Call DecimalDegrees
                                strOutputA = dblLatdd
                                strOutputB = dblLondd
                                
                                'calculate the atlas sheet coordinate
                                Call LatToAt
                                strOutputC = strAscCoor
                                
                                'convert the latitude/longitude NAD83 to NAD27
                                Call Datum
                                dblLatdecimal = dblLatDatum
                                dblLondecimal = dblLonDatum
                                
                                'convert decimal degrees to degrees, minutes and decimal seconds
                                Call DecimalDegrees
                                strOutputD = dblLatdd
                                strOutputE = dblLondd
                                
                                'write variables to file convert.dat
                                Write #4, strId, strInputA, strInputB, strOutputA, _
                                strOutputB, strOutputD, strOutputE, strOutputC, strQnam
                             Else
                                GoTo Err_cmdRunBatch_Click
                            End If
                        Else
                            GoTo Err_cmdRunBatch_Click
                        End If
                    Else                                'output in decimal degrees
                        dblNorthing = strInputA
                        dblEasting = strInputB
                        
                        If dblNorthing >= -73899.055 And dblNorthing <= 927030.312 Then
                            If dblEasting >= 170739.055 And dblEasting <= 732257.404 Then
                            
                                'convert state plane coordinates to latitude/longitude NAD83
                                Call xytoll83
                                strOutputA = Round(dblLatdecimal, 6)
                                strOutputB = Round(dblLondecimal, 6)
                                
                                'convert decimal degrees to degrees, minutes and decimal
                                Call DecimalDegrees
                                
                                'calculate the atlas sheet coordinate
                                Call LatToAt
                                strOutputC = strAscCoor
                                
                                'convert the latitude/longitude NAD83 to NAD27
                                Call Datum
                                strOutputD = Round(dblLatDatum, 6)
                                strOutputE = Round(dblLonDatum, 6)
                                
                                'write variables to file convert.dat
                                Write #4, strId, strInputA, strInputB, strOutputA, _
                                strOutputB, strOutputD, strOutputE, strOutputC, strQnam
                            Else
                                GoTo Err_cmdRunBatch_Click
                            End If
                        Else
                            GoTo Err_cmdRunBatch_Click
                        End If
                    End If

                Case 1      'latitude/longitude to state plane
                    If optPick(3).Value = True Then     'input units ddmmss.ss
                        dblLatdd = strInputA
                        dblLondd = strInputB
                        
                        'disassemble dblLatdd and check for valid latitude
                        If dblLatdd < 383730 Or dblLatdd > 412230 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                        
                        dblLatDeg = Left(dblLatdd, 2)
                            
                        dblLatMin = Mid(dblLatdd, 3, 2)
                        If dblLatMin > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                            
                        dblLatSec = Mid(dblLatdd, 5, 6)
                        If Int(dblLatSec) > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                        
                        'disassemble dblLondd and check for valid longitude
                        If dblLondd > 753730 Or dblLondd < 733730 Then
                            GoTo Err_cmdRunBatch_Click
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
                            GoTo Err_cmdRunBatch_Click
                        End If
    
                        If Int(dblLonSec) > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                                
                        'calculate the atlas sheet coordinate
                        Call LatToAt
                        strOutputC = strAscCoor
                                
                        If optPick(6).Value = True Then
                                
                            'convert degrees, minutes and decimal seconds to decimal degrees
                            Call Ddmmss
                    
                            'convert latitude/longitude to state plane coordinates
                            Call lltoxy83
                            strOutputA = Round(dblNorthing, 1)
                            strOutputB = Round(dblEasting, 1)
                                    
                        End If
                                
                        'convert the latitude/longitude NAD27 <=> NAD83
                        Call Datum
                        dblLatdecimal = dblLatDatum
                        dblLondecimal = dblLonDatum
                                
                        'convert decimal degrees to degrees, minutes and decimal seconds
                        Call DecimalDegrees
                        strOutputD = dblLatdd
                        strOutputE = dblLondd
                                
                        If optPick(5).Value = True Then
                                
                            'convert latitude/longitude to state plane coordinates
                            Call lltoxy83
                            strOutputA = Round(dblNorthing, 1)
                            strOutputB = Round(dblEasting, 1)
                                    
                        End If
                        
                        If optPick(6).Value = True Then   'write variables to file convert.dat
                            Write #4, strId, strOutputA, strOutputB, strInputA, _
                            strInputB, strOutputD, strOutputE, strOutputC, strQnam
                        Else
                            Write #4, strId, strOutputA, strOutputB, strOutputD, _
                            strOutputE, strInputA, strInputB, strOutputC, strQnam
                        End If
                        
                    Else
                        'input units decimal degrees
                        dblLatdecimal = strInputA
                        dblLondecimal = strInputB
                        
                        If dblLatdecimal >= 38.625 And dblLatdecimal <= 41.375 Then
                            If dblLondecimal <= 75.625 And dblLondecimal >= 73.625 Then
                            
                                If optPick(6).Value = True Then
                                                  
                                    'convert latitude/longitude to state plane coordinates
                                    Call lltoxy83
                                    strOutputA = Round(dblNorthing, 0)
                                    strOutputB = Round(dblEasting, 0)
                                    
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
                                strOutputC = strAscCoor
                                
                                'convert the latitude/longitude NAD83 <=> NAD27
                                Call Datum
                                dblLatdecimal = dblLatDatum
                                dblLondecimal = dblLonDatum
                                strOutputD = Round(dblLatDatum, 6)
                                strOutputE = Round(dblLonDatum, 6)
                                
                                If optPick(5).Value = True Then
                                                  
                                    'convert latitude/longitude to state plane coordinates
                                    Call lltoxy83
                                    strOutputA = Round(dblNorthing, 0)
                                    strOutputB = Round(dblEasting, 0)
                                                                        
                                End If
                                
                                If optPick(6).Value = True Then   'write variables to file convert.dat
                                    Write #4, strId, strOutputA, strOutputB, strInputA, _
                                    strInputB, strOutputD, strOutputE, strOutputC, strQnam
                                Else
                                    Write #4, strId, strOutputA, strOutputB, strOutputD, _
                                    strOutputE, strInputA, strInputB, strOutputC, strQnam
                                End If
                                
                            Else
                                GoTo Err_cmdRunBatch_Click
                            End If
                        Else
                            GoTo Err_cmdRunBatch_Click
                        End If
                    End If
                    
                Case 2      'atlas sheet coordinate to state plane and latitude/longitude
                        strAscCoor = strInputC
                        strOutputC = strAscCoor
                        
                        Call AtToLat
                        strInputA = Round(dblLatdd, 0)
                        strInputB = Round(dblLondd, 0)
                        
                        dblLatdd = strInputA
                        dblLondd = strInputB
                        
                        'disassemble dblLatdd and check for valid latitude
                        If dblLatdd < 383730 Or dblLatdd > 412230 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                        
                        dblLatDeg = Left(dblLatdd, 2)
                            
                        dblLatMin = Mid(dblLatdd, 3, 2)
                        If dblLatMin > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                            
                        dblLatSec = Mid(dblLatdd, 5, 6)
                        If Int(dblLatSec) > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                        
                        'disassemble dblLondd and check for valid longitude
                        If dblLondd > 753730 Or dblLondd < 733730 Then
                            GoTo Err_cmdRunBatch_Click
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
                            GoTo Err_cmdRunBatch_Click
                        End If
    
                        If Int(dblLonSec) > 59 Then
                            GoTo Err_cmdRunBatch_Click
                        End If
                              
                        'convert the latitude/longitude NAD27 <=> NAD83
                        Call Datum
                        dblLatdecimal = dblLatDatum
                        dblLondecimal = dblLonDatum
                                
                        'convert decimal degrees to degrees, minutes and decimal seconds
                        Call DecimalDegrees
                        strOutputD = dblLatdd
                        strOutputE = dblLondd
                                
                        'convert latitude/longitude to state plane coordinates
                        Call lltoxy83
                        strOutputA = Round(dblNorthing, 1)
                        strOutputB = Round(dblEasting, 1)
                        
                        Write #4, strId, strOutputA, strOutputB, strOutputD, _
                        strOutputE, strInputA, strInputB, strOutputC, strQnam
            End Select
        End If
    Next intCount

Err_Resume:

If intErr = 3 And optPick(2).Value = True Then
    Write #5, strId, strInputC
    intErr = 0
End If

If intErr = 3 And optPick(2).Value = False Then
    Write #5, strId, strInputA, strInputB
    intErr = 0
End If

Loop

Close #3
Close #4
Close #5

Exit Sub

Err_dlgCommon_Select:
    MsgBox prompt:="A file was not selected!"
    Exit Sub
    
Err_cmdRunBatch_Click:
    intErr = 3
    Resume Err_Resume
        
End Sub

Private Sub Command1_Click()
dlgCommon.FileName = ""     'get file for conversion
dlgCommon.ShowOpen
strPoints = dlgCommon.FileName

If strPoints = "" Then GoTo Err_dlgCommon_Select

Exit Sub

Err_dlgCommon_Select:
    MsgBox prompt:="A file was not selected!"
    Exit Sub
    
End Sub


Private Sub Form_Load()
    frmBatch.Top = (Screen.Height - frmBatch.Height) / 2
    frmBatch.Left = (Screen.Width - frmBatch.Width) / 2
   
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
    Unload frmBatch
End Sub

Private Sub mnuFileGotoDecimaldegreesmodule_Click()
    frmDegrees.Show
    Unload frmBatch
End Sub

Private Sub mnuFileGotoMainpage_Click()
    frmMain.Show
    Unload frmBatch
End Sub

Private Sub mnuFileGotoStatepalemodule_Click()
    frmLatXy.Show
    Unload frmBatch
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

    For intCount = 0 To 2
        If optPick(intCount).Value = True Then
            Select Case intCount
                Case 0
                    blnNAD83 = False
                    fraDatum.Visible = False
                    fraUnits.Visible = True
                Case 1
                    fraDatum.Visible = True
                    fraUnits.Visible = True
                    If optPick(3).Value = True Then
                        If optPick(5).Value = True Then
                            blnNAD83 = True
                        Else
                            blnNAD83 = False
                        End If
                    Else
                        If optPick(5).Value = True Then
                            blnNAD83 = True
                        Else
                            blnNAD83 = False
                        End If
                    End If
                Case 2
                    blnNAD83 = True
                    fraDatum.Visible = False
                    fraUnits.Visible = False
            End Select
        End If
    Next intCount

End Sub

