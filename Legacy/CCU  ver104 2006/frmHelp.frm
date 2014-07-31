VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmHelpMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NJGS Utility Help"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtfLimit 
      Height          =   5415
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      FileName        =   "C:\Documents and Settings\bmennel\My Documents\VBworkver3\Limitations.rtf"
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.CommandButton cmdNad27 
      Caption         =   "NAD27"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpcs 
      Caption         =   "State Plane"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAscs 
      Caption         =   "Atlas Sheet"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEntry 
      Caption         =   "Entry"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimit 
      Caption         =   "Limitations"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtfEntry 
      Bindings        =   "frmHelp.frx":0297
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Documents and Settings\bmennel\My Documents\VBworkver3\Conventions.rtf"
      TextRTF         =   $"frmHelp.frx":02A2
   End
   Begin RichTextLib.RichTextBox rtfAscs 
      Bindings        =   "frmHelp.frx":0861
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Documents and Settings\bmennel\My Documents\VBworkver3\Ascs.rtf"
      TextRTF         =   $"frmHelp.frx":086C
   End
   Begin RichTextLib.RichTextBox rtfSpcs 
      Bindings        =   "frmHelp.frx":1304
      Height          =   5415
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "Sp.rtf"
      TextRTF         =   $"frmHelp.frx":130F
   End
   Begin RichTextLib.RichTextBox rtfNAD27 
      Bindings        =   "frmHelp.frx":1764
      Height          =   5415
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Documents and Settings\bmennel\My Documents\VBworkver3\Nad27.rtf"
      TextRTF         =   $"frmHelp.frx":176F
   End
End
Attribute VB_Name = "frmHelpMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAscs_Click()
    rtfLimit.Visible = False
    rtfEntry.Visible = False
    rtfAscs.Visible = True
    rtfSpcs.Visible = False
    rtfNAD27.Visible = False
End Sub

Private Sub cmdEntry_Click()
    rtfLimit.Visible = False
    rtfEntry.Visible = True
    rtfAscs.Visible = False
    rtfSpcs.Visible = False
    rtfNAD27.Visible = False
End Sub
Private Sub cmdLimit_Click()
    rtfLimit.Visible = True
    rtfEntry.Visible = False
    rtfAscs.Visible = False
    rtfSpcs.Visible = False
    rtfNAD27.Visible = False
End Sub

Private Sub cmdNad27_Click()
    rtfLimit.Visible = False
    rtfEntry.Visible = False
    rtfAscs.Visible = False
    rtfSpcs.Visible = False
    rtfNAD27.Visible = True
End Sub

Private Sub cmdSpcs_Click()
    rtfLimit.Visible = False
    rtfEntry.Visible = False
    rtfAscs.Visible = False
    rtfSpcs.Visible = True
    rtfNAD27.Visible = False
End Sub
