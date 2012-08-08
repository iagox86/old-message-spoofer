VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About D2backstab.com Message Spoofer"
   ClientHeight    =   2655
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1832.528
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2265
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Programmed by iago <iago@d2backstab.com>"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1439.104
      Y2              =   1439.104
   End
   Begin VB.Label lblDescription 
      Caption         =   "This is a program for spoofing your messages on battle.net while playing Brood War."
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "D2backstab.com Message Spoofer 0.5"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1449.457
      Y2              =   1449.457
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
