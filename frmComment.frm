VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmComment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Question/Comment/Bug report"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws 
      Left            =   1680
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.d2backstab.com"
      RemotePort      =   80
      LocalPort       =   4787
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtmessage 
      Height          =   285
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Your question/comment/bug report"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Your email address"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Your name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSend_Click()
    On Error Resume Next
    
    ws.LocalPort = Int(Rnd * 25000)
    
    ws.Connect
    
End Sub

Private Sub Form_Load()
    On Error Resume Next

    txtName.SetFocus
End Sub

Private Sub ws_Close()
    ws.Close
End Sub

Private Sub ws_Connect()
    ws.SendData "GET /cgi-bin/email.pl?name=" & Replace(txtName.Text, " ", "+") & "&email=" & Replace(txtEmail.Text, " ", "+") & "&message=" & Replace(txtmessage.Text, " ", "+") & " http/1.1" & vbCrLf & _
                "Host:www.d2backstab.com" & vbCrLf & _
                vbCrLf
   
    MsgBox "Comment sent by email.  I check my email at least once a day so I'll probably respond soon!", vbInformation, "Woohoo!"
    Me.Hide
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock error!" & vbCrLf & "Error no: " & Number & vbCrLf & "Description: " & Description, vbCritical, "Winsock Error"
    ws.Close
    Call ws_Close
End Sub

Private Sub ws_SendComplete()
    ws.Close
End Sub
