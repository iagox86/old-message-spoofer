VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D2backstab.com Message Spoofer *beta*"
   ClientHeight    =   3510
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5430
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer tmrCheckIfGameIsBeingPlayed 
      Interval        =   5000
      Left            =   4200
      Top             =   0
   End
   Begin VB.Timer tmrSpoofer 
      Interval        =   500
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer tmrSlowText 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   0
   End
   Begin VB.PictureBox picSystemTray 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4920
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame4 
      Caption         =   "Username (Appears in front of /.../ messages) [/name $name/]"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   5175
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Align (/.../)"
      Height          =   1335
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
      Begin VB.OptionButton optRight 
         Caption         =   "Right [/right]"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optBCenter 
         Caption         =   "BCenter [/bcenter]"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Center [/center]"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left [/noalign]"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Default color (for /.../ messages)"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
      Begin VB.OptionButton optColor 
         Caption         =   "Green [/green]"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Red [/red]"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Grey [/grey]"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optColor 
         Caption         =   "White [/white]"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Yellow [/yellow]"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optColor 
         Caption         =   "None [/nocolor]"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkTeam 
         Caption         =   "Show team name always [/team]"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "This is useful for TvB games, but bad for melee"
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkBanned 
         Caption         =   "Allow banned characters (ALT-003, etc) [/banned]"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Disable this if crashes occur"
         Top             =   480
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkReplaceColors 
         Caption         =   "Check colors as you type [/colors]"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Disable this if crashes occur"
         Top             =   240
         Value           =   1  'Checked
         Width           =   3135
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHow 
         Caption         =   "&How does this damn thing work?..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpBugs 
         Caption         =   "&Send Bug Report..."
      End
   End
   Begin VB.Menu mnuST 
      Caption         =   "SystemTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSTRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuSTExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrFindText As String
Const MAX_LENGTH As Integer = &H50
Const MEM_OFFSET As Long = &H88
'MemoryOffset is the beginning of the chat message
Dim MemoryOffset As Long
'BadMemoryOffset is MemoryOffset - MEM_OFFSET
Dim BadMemoryOffset As Long
Dim Restoring As Boolean
'The next ones for slow text/marquee
Dim strSlowText As String
Dim strMarquee As String
Dim iPosition As Integer
Dim bMarquee As Boolean
Dim bMarqueeLeft As Boolean
Dim iOffset As Integer
Dim iMarqueeSpeed As Integer
'These are for the animation(s):
Dim strAnimation(1 To 10) As String * 80
Dim iFrame As Integer


'Some flags for different effects:
Dim bRight As Boolean
Dim bCenter As Boolean
Dim bCenterBottom As Boolean
Dim strColor As String * 1
Dim strName As String

'In strings:
Dim bwEND As String
Dim bwNOCOLOR As String
Dim bwBLUE As String
Dim bwYELLOW As String
Dim bwWHITE As String
Dim bwGREY As String
Dim bwRED As String
Dim bwGREEN As String
Dim bwCRLF As String
Dim bwCRASH As String
Dim bwRIGHT As String
Dim bwCENTER As String

Dim AddLineFeeds As String

'Nickspoofer stuff:
Const BROODWAR As Long = &H190350A0
Const BROODWAR_REAL As Long = &H19035338
Const ADDR_PTR As Long = &H170001C
Const CMD_LENGTH As Long = &H19033B10
Const CMD_SELSTART As Long = &H19033B18
Const CMD_CURSORPOSITION As Long = &H19033B1C

Dim lCommandAddr As Long
Dim strCurrentNickname As String * 21

Private Sub FindMemAddress()
    Dim iIndex As Long
    Dim strReadString As String
   
   'debug
    'Exit Sub
   
    For iIndex = &H10000 To &H10000000 Step &H10000
        DoEvents
        strReadString = Left(ReadMemory(iIndex, Len(mstrFindText)), Len(mstrFindText))
        If strReadString = mstrFindText Then
            tmrSpoofer.Enabled = True
            MemoryOffset = iIndex + MEM_OFFSET
            BadMemoryOffset = iIndex
            Exit Sub
        End If
    Next
    
End Sub

Private Sub chkBanned_Click()
    SaveSetting "D2backstab.com Message Spoofer", "Options", "Banned", chkBanned.Value
    DoBannedChange
End Sub

Private Sub chkReplaceColors_Click()
    SaveSetting "D2backstab.com Message Spoofer", "Options", "Colors", chkReplaceColors.Value
End Sub

Private Sub chkTeam_Click()
    SaveSetting "D2backstab.com Message Spoofer", "Options", "Teams", chkTeam.Value
    DoTeamChange
End Sub

Private Sub Form_Activate()
    bwEND = Chr(0)
    bwNOCOLOR = Chr(1)
    bwBLUE = Chr(2)
    bwYELLOW = Chr(3)
    bwWHITE = Chr(4)
    bwGREY = Chr(5)
    bwRED = Chr(6)
    bwGREEN = Chr(7)
    bwCRLF = Chr(&HA)
    bwCRASH = Chr(&HC)
    bwRIGHT = Chr(&H12)
    bwCENTER = Chr(&H13)
    AddLineFeeds = bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF & bwCRLF
    
    strColor = Chr(1)
    iMarqueeSpeed = 1
    
    chkReplaceColors.Value = Val(GetSetting("D2backstab.com Message Spoofer", "Options", "Colors"))
    chkBanned.Value = Val(GetSetting("D2backstab.com Message Spoofer", "Options", "Banned"))
    chkTeam.Value = Val(GetSetting("D2backstab.com Message Spoofer", "Options", "Teams"))
    FindMemAddress
    DoBannedChange
    DoTeamChange
End Sub

Private Sub Form_Load()
    mstrFindText = Chr(0) & Chr(0) & Chr(0) & Chr(0) & _
                  Chr(&HAF) & Chr(&H87) & Chr(&H4D) & Chr(&H58) _
                  & Chr(&HAF)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close the handle we opened.
    CloseHandle Phand
    
    'End the program right
    End
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelpBugs_Click()
    frmComment.Show vbModal
End Sub

Private Sub mnuHelpHow_Click()
    MsgBox "I have included a help file in .html format which is " & vbCrLf & _
           "easily printable.  I suggest you print that and use it " & vbCrLf & _
           "as reference.", vbInformation, "How does this damn thing work?"
End Sub




Private Sub optColor_Click(Index As Integer)
    If Index <> 1 Then
        strColor = Chr(Index)
    Else
        strColor = ""
    End If
End Sub

Private Sub optLeft_Click()
    bCenter = False
    bRight = False
    bCenterBottom = False
End Sub
Private Sub optRight_Click()
    bCenter = False
    bRight = True
    bCenterBottom = False
End Sub
Private Sub optCenter_Click()
    bCenter = True
    bRight = False
    bCenterBottom = False
End Sub
Private Sub optBCenter_Click()
    bCenter = False
    bRight = False
    bCenterBottom = True
End Sub

Private Sub tmrAnimation_Timer()
    Dim strTemp As String * 1
    Dim strMessage As String
    Dim lLen As Integer


    
    strTemp = ReadMemory(MemoryOffset, 1)
    
    If strTemp = Chr(0) Then
        iFrame = iFrame + 1
        If iFrame > UBound(strAnimation) Then
            tmrAnimation.Enabled = False
            Exit Sub
        End If
        strMessage = strAnimation(iFrame)
        WriteMemory strMessage, MemoryOffset, Len(strMessage)
    End If
    
End Sub

Private Sub tmrCheckIfGameIsBeingPlayed_Timer()
    'This checks if the user is in the game that the memory address is pointing to.
    'If they aren't, it searches for the game
    Dim strFoundText
    strFoundText = ReadMemory(BadMemoryOffset, Len(mstrFindText))
    
    If strFoundText <> mstrFindText Then
        FindMemAddress
    End If
End Sub


Private Sub tmrSlowText_Timer()
    Dim strTemp As String * 1
    Dim strStringToWrite As String
    
    On Error Resume Next
    
    If iPosition >= Len(strSlowText) Then
        If Not (bMarquee) Then
            tmrSlowText.Enabled = False
            Exit Sub
        Else
            bMarqueeLeft = True
            iPosition = iPosition - iOffset
        End If
    End If
    
    If bMarqueeLeft And iPosition = 12 Then
        tmrSlowText.Enabled = False
        Exit Sub
    End If
    
    If bMarqueeLeft And iPosition < 12 Then
        strMarquee = ""
        iPosition = 12 + iMarqueeSpeed
    End If
    
    strTemp = ReadMemory(MemoryOffset, 1)
    If strTemp = Chr(0) Then
        If bMarqueeLeft And bMarquee Then
            iPosition = iPosition - iMarqueeSpeed
        Else
            iPosition = iPosition + iMarqueeSpeed
        End If
        
        If bMarqueeLeft Then
            strStringToWrite = Right(strMarquee, iPosition - 12) & Chr(0)
            'strStringToWrite = Replace(strStringToWrite, " ", "", 1, Len(strSlowText) - iPosition)
            strStringToWrite = AddLineFeeds & strStringToWrite
        Else
            strStringToWrite = Left(strSlowText, iPosition) & Chr(0)
        End If
        
        WriteMemory strStringToWrite, MemoryOffset, Len(strStringToWrite)
    End If
End Sub

Private Sub tmrSpoofer_Timer()
    'Holds the message they typed
    Dim strMessage As String
    'Hold the unchanged message to determine if a write is needed
    Dim strOldMessage As String
    'Holds an array if the message requires parameters
    Dim strMessageArray() As String
    'Loop variable
    Dim iIndex As Integer
    
    On Error Resume Next
    
    tmrSpoofer.Enabled = False
    'Read the message that's in the message buffer
    strMessage = ReadMemory(MemoryOffset, MAX_LENGTH)
    
    'Cut down the message at the null-termination
    strMessage = Left(strMessage, InStr(1, strMessage, Chr(0)) - 1)
    strOldMessage = strMessage
    
    'These flags affect messages enclosed in /'s
    Select Case strMessage
    Case "/normal"
        bCenter = False
        bCenterBottom = False
        bRight = False
        optColor(Asc(bwNOCOLOR)).Value = True
        txtUsername.Text = ""
    Case "/center"
        optCenter.Value = True
        strMessage = bwGREEN & "Center has been set." & AddLineFeeds
    Case "/bcenter"
        optBCenter.Value = True
        strMessage = bwGREEN & "BottomCenter has been set." & AddLineFeeds
    Case "/right"
        optRight.Value = True
        strMessage = bwGREEN & "Right has been set." & AddLineFeeds
    Case "/noalign"
        optLeft.Value = True
        strMessage = bwGREEN & "Left has been set." & AddLineFeeds
    Case "/left"
        optLeft.Value = True
        strMessage = bwGREEN & "Left has been set." & AddLineFeeds
    Case "/yellow"
        optColor(Asc(bwYELLOW)).Value = 1
        strMessage = bwYELLOW & "Yellow has been set." & AddLineFeeds
    Case "/white"
        optColor(Asc(bwWHITE)).Value = 1
        strMessage = bwWHITE & "White has been set." & AddLineFeeds
    Case "/grey"
        optColor(Asc(bwGREY)).Value = 1
        strMessage = bwGREY & "Grey has been set." & AddLineFeeds
    Case "/red"
        optColor(Asc(bwRED)).Value = 1
        strMessage = bwRED & "Red has been set." & AddLineFeeds
    Case "/green"
        optColor(Asc(bwGREEN)).Value = 1
        strMessage = bwGREEN & "Green has been set." & AddLineFeeds
    Case "/nocolor"
        optColor(Asc(bwNOCOLOR)).Value = 1
        strMessage = bwNOCOLOR & "Color has been removed." & AddLineFeeds
    Case "/colors"
        If chkReplaceColors.Value = 1 Then
            chkReplaceColors.Value = 0
            strMessage = bwRED & "Replace Colors Disabled" & AddLineFeeds
        Else
            chkReplaceColors.Value = 1
            strMessage = bwGREEN & "Replace Colors Enabled" & AddLineFeeds
        End If
    Case "/banned"
        If chkBanned.Value = 1 Then
            chkBanned.Value = 0
            strMessage = bwRED & "Banned Characters Disabled" & AddLineFeeds
        Else
            chkBanned.Value = 1
            strMessage = bwGREEN & "Banned Colors Enabled" & AddLineFeeds
        End If
    Case "/team"
        If chkTeam.Value = 1 Then
            chkTeam.Value = 0
            strMessage = bwRED & "Force Team Disabled" & AddLineFeeds
        Else
            chkTeam.Value = 1
            strMessage = bwGREEN & "Force Team Enabled" & AddLineFeeds
        End If
    End Select
    
    If Left(strMessage, Len("/name")) = "/name" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        strMessage = bwYELLOW & "Name is now " & strMessageArray(1) & AddLineFeeds & Chr(0)
        txtUsername.Text = strMessageArray(1)
    ElseIf strMessage = "/noname" Then
        txtUsername.Text = ""
        strMessage = bwRED & "Name has been removed." & AddLineFeeds
    End If
    
    If Left(strMessage, Len("/slowtext")) = "/slowtext" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        strMessage = "Hold [enter] to continue" & AddLineFeeds
        iPosition = 12
        strSlowText = AddLineFeeds & strMessageArray(1)
        tmrSlowText.Enabled = True
        bMarquee = False
    End If
    
    
    If Left(strMessage, Len("/mspeed")) = "/mspeed" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        If Val(strMessageArray(1)) <= 20 And Val(strMessageArray(1)) >= 1 Then
            strMessage = "Marquee Speed Set" & AddLineFeeds
            iMarqueeSpeed = Val(strMessageArray(1))
        Else
            strMessage = "Speed must be from 1-20" & AddLineFeeds
        End If
    End If
    
    If Left(strMessage, Len("/marquee")) = "/marquee" And Right(strMessage, 1) = "/" Then
        Dim strSpaces As String
        
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        strMessage = "Hold [enter] to continue" & AddLineFeeds
        iPosition = 12
        tmrSlowText.Enabled = True
        bMarquee = True
        bMarqueeLeft = False
        
        iOffset = (15 / Len(strMessageArray(1))) - 1
        If iOffset < 1 Then
            'for long messages
            iOffset = Int(-(1 / iOffset) * 4)
        Else
            'for short messages
            iOffset = Int(iOffset / 2)
        End If
        
        'as much as I hate doing it, I have to pad this with spaces
        For iIndex = Len(AddLineFeeds) + Len(strMessageArray(1)) To 78
            strSpaces = " " & strSpaces
        Next

        strSlowText = AddLineFeeds & bwRIGHT & strMessageArray(1) & strSpaces
        strMarquee = strSpaces & strMessageArray(1)
    End If
    
    If strMessage = "/smile" Then
        strAnimation(1) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "0  0" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " |" & bwCRLF & _
                          Chr(3) & "\_____/"
                          
        strAnimation(2) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "^  ^" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " |" & bwCRLF & _
                          Chr(3) & "\_____/"
                          
        strAnimation(3) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "0  0" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " |" & bwCRLF & _
                          Chr(3) & "\_____/"
                          
        strAnimation(4) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "^  ^" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " |" & bwCRLF & _
                          Chr(3) & "\_____/"
                          
        strAnimation(5) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "0  0" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " | -- kekeke" & bwCRLF & _
                          Chr(3) & "\_____/"
                          
        strAnimation(6) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          Chr(3) & " _____" & bwCRLF & _
                          Chr(3) & "/ " & Chr(7) & "^  ^" & Chr(1) & " \" & bwCRLF & _
                          Chr(3) & "|   v    |" & bwCRLF & _
                          Chr(3) & "| " & Chr(6) & "\___/" & Chr(1) & " | -- kekeke" & bwCRLF & _
                          Chr(3) & "\_____/"
                          

        iFrame = 0
        tmrAnimation.Enabled = True
        strMessage = "Press [enter] to continue." & AddLineFeeds
    End If
    'Moo!
    If strMessage = "/cow" Then

        strAnimation(1) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          " //--------\\" & bwCRLF & _
                          "^^       ^^"
        strAnimation(2) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          "  | |-----| |" & bwCRLF & _
                          "  ^^    ^^"

        strAnimation(3) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          " //--------\\" & bwCRLF & _
                          "^^       ^^"
        strAnimation(4) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          "  | |-----| |" & bwCRLF & _
                          "  ^^    ^^"

        strAnimation(5) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          " //--------\\" & bwCRLF & _
                          "^^       ^^"
        strAnimation(6) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          "  | |-----| |" & bwCRLF & _
                          "  ^^    ^^"

        strAnimation(7) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          " //--------\\" & bwCRLF & _
                          "^^       ^^"
        strAnimation(8) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*      (__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          "  | |-----| |" & bwCRLF & _
                          "  ^^    ^^"
         strAnimation(8) = bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          bwCRLF & _
                          "*" & Chr(4) & "'Moo!'" & Chr(1) & "(__)" & bwCRLF & _
                          " \     (oo)" & bwCRLF & _
                          "  \-------\/" & bwCRLF & _
                          "  | |-----| |" & bwCRLF & _
                          "  ^^    ^^"
        
        iFrame = 0
        tmrAnimation.Enabled = True
        strMessage = "Press [enter] to continue." & AddLineFeeds
    End If
        
    '"Nuclear launch detected"
    If strMessage = "/nuke" Then
        strMessage = AddLineFeeds
        strMessage = strMessage & bwWHITE & bwCENTER & "Nuclear launch detected."
    End If
    
    'Cheat enabled
    If strMessage = "/cheat" Then
        strMessage = AddLineFeeds
        strMessage = strMessage & bwWHITE & bwCENTER & "Cheat enabled"
    End If
    
    'You have been backstabbed!
    If strMessage = "/bs" Then
        strMessage = AddLineFeeds
        strMessage = strMessage & bwRED & bwCENTER & "You have been backstabbed!" & bwCRLF & _
                                  bwWHITE & bwCENTER & "http://www.d2backstab.com"
    End If
    
    '$user has left the game
    If Left(strMessage, Len("/leave")) = "/leave" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        
        strMessage = AddLineFeeds
        strMessage = strMessage & bwYELLOW & strMessageArray(1) & " has left the game."
    End If
    
    '$user was eliminated
    If Left(strMessage, Len("/kill")) = "/kill" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        
        strMessage = AddLineFeeds
        strMessage = strMessage & bwYELLOW & strMessageArray(1) & " was eliminated."
    End If
    
    If Left(strMessage, Len("/drop")) = "/drop" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        
        strMessage = AddLineFeeds
        strMessage = strMessage & bwYELLOW & strMessageArray(1) & " was dropped from the game."
    End If
    
    If Left(strMessage, Len("/join")) = "/join" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 2)
        
        strMessage = AddLineFeeds
        strMessage = strMessage & bwYELLOW & strMessageArray(1) & " has joined the game."
    End If
    
    If Left(strMessage, Len("/latency")) = "/latency" And Right(strMessage, 1) = "/" Then
        strMessage = Left(strMessage, Len(strMessage) - 1)
        strMessageArray = Split(strMessage, " ", 3)
        
        strMessage = AddLineFeeds
        strMessage = strMessage & bwYELLOW & "Player " & strMessageArray(1) & " set network for " & strMessageArray(2) & " latency"
    End If
    
    
    'Messages starting and ending with / indicate commands have 12 linefeeds added to them
    If Left(strMessage, 1) = "/" And Right(strMessage, 1) = "/" And Len(strMessage) > 1 Then
        'First, we get rid of the / at the beginning and end
        strMessage = Mid(strMessage, 2, Len(strMessage) - 2)
        
        If strName <> "" Then
            strMessage = strName & ": " & strMessage
        End If
                
        strMessage = strColor & strMessage
        strMessage = Replace(strMessage, "\n", "\n" & strColor)
        
        If (bCenter) Then
            strMessage = bwCENTER & strMessage
            strMessage = Left(AddLineFeeds, 6) & strMessage & Left(AddLineFeeds, 6)
        Else
            If bCenterBottom Then
                strMessage = bwCENTER & strMessage
            ElseIf bRight Then
                strMessage = bwRIGHT & strMessage
            End If
            
            'Now we add 12 newlines to the beginning so you don't see name: before your message
            strMessage = AddLineFeeds & strMessage
        End If
    
    'If the checkbox is unchecked, do colors
    If chkReplaceColors.Value = 0 Then
        strMessage = ChangeColors(strMessage)
    End If
        
    strMessage = Replace(strMessage, "\n", bwCRLF)
    
    End If
    
    'Now we change \3, \4, \5 and \n to yellow, white, grey, and linefeed
    If chkReplaceColors.Value = 1 Then
        strMessage = ChangeColors(strMessage)
    End If
    If strMessage <> strOldMessage Then
        strMessage = strMessage & bwEND
        WriteMemory strMessage, MemoryOffset, Len(strMessage)
    End If
    
    tmrSpoofer.Enabled = True
    
   
End Sub


Private Function ChangeColors(strMessage As String) As String
    strMessage = Replace(strMessage, "\1", Chr(1) & Chr(1))
    strMessage = Replace(strMessage, "\2", Chr(2) & Chr(2))
    strMessage = Replace(strMessage, "\3", Chr(3) & Chr(3))
    strMessage = Replace(strMessage, "\4", Chr(4) & Chr(4))
    strMessage = Replace(strMessage, "\5", Chr(5) & Chr(5))
    strMessage = Replace(strMessage, "\6", Chr(6) & Chr(6))
    strMessage = Replace(strMessage, "\7", Chr(7) & Chr(7))
    
    ChangeColors = strMessage
End Function

Private Sub txtUsername_Change()
    strName = txtUsername.Text
End Sub












Private Sub Form_Resize()
    If frmMain.WindowState = 1 And Restoring <> True Then
        App.TaskVisible = False
        Me.Hide
        CreateIcon
    ElseIf frmMain.WindowState = 0 Then
        Me.Show
        App.TaskVisible = True
        Restoring = False
        DeleteIcon
    End If
End Sub

Private Sub picSystemTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX
    Select Case X
        Case WM_LBUTTONDOWN
            Me.PopupMenu mnuST
        Case WM_RBUTTONDOWN
            Me.PopupMenu mnuST
        Case WM_LBUTTONDBLCLK
            Restoring = True
            App.TaskVisible = True
            Me.Show
            Me.WindowState = 0
            DeleteIcon
    End Select
End Sub

Private Sub mnuSTExit_Click()
    DeleteIcon
    Unload Me
    End
End Sub

Private Sub mnuSTRestore_Click()
    Restoring = True
    App.TaskVisible = True
    Me.Show
    Me.WindowState = 0
    DeleteIcon
End Sub



