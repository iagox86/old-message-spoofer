Attribute VB_Name = "modMessageSpoofer"
'modMessageSpoofer.bas

'This module is for spoofing messages so that they contain colors
'or other funky effects :)

'Process Handle
Public Phand As Long


'for retrieving the game process handle
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'get processID from hwnd returned above
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'return handle to target process
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'for writing game memory
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'for reading game memory
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'close each open handle
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'for determining OS version
Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
'User defined type for OSVERSIONINFO
    Public Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long           '1 = Windows 95.
                                      '2 = Windows NT

       szCSDVersion As String * 128
    End Type
'Necessary flags for NT
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)


Function ReadMemory(lMemLocation As Long, iLength As Integer) As String
    'Length returned
    Dim LengthReturned As Long
    'String for containing nickname
    Dim strMessage As String * &HFF
    
    If (Phand = 0) Then
        GetPhand
    End If
    
    If (Phand = 0) Then
        ReadMemory = "Error! Unable to get process handle!"
        Exit Function
    End If
    
    ReadProcessMemory Phand, lMemLocation, strMessage, iLength, LengthReturned
    
    If (LengthReturned = 0) Then
        ReadMemory = "Error, could not read memory, continuing search..."
        Exit Function
    End If
    
    'And finally, return the nickname read.
    ReadMemory = Left(strMessage, iLength)
    'CloseHandle Phand
End Function

Function WriteMemory(strData As String, lMemLocation As Long, iLength As Integer) As String
    'Length returned
    Dim LengthReturned As Long
    'String for containing nickname
    Dim strMessage As String
    
    'If we don't already have a program handle, get it
    If (Phand = 0) Then
        GetPhand
    End If
    
    If (Phand = 0) Then
        Exit Function
    End If
    
    
    WriteProcessMemory Phand, lMemLocation, strData, iLength, LengthReturned
    
    If (LengthReturned = 0) Then
        Exit Function
    End If

End Function
Function getVersion() As Long
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
getVersion = osinfo.dwPlatformId
End Function

Sub GetPhand()
    'Window Handles
    Dim bwWnd As Long
    'Process ID
    Dim Pid As Long
    
    'Get processid's
    bwWnd = FindWindow(vbNullString, "Brood War")
    If bwWnd = 0 Then
        Phand = 0
        Exit Sub
    End If
        
    'Convert window handle to process id
    GetWindowThreadProcessId bwWnd, Pid
    
    'Get process handle
    GetWindowThreadProcessId bwWnd, Pid
    Phand = OpenProcess(PROCESS_ALL_ACCESS, False, Pid)


    DoBannedChange
    DoTeamChange
End Sub

Sub DoBannedChange()
    If frmMain.chkBanned.Value = 1 Then
        WriteMemory Chr(&H90) & Chr(&H90), &H4E62EC, 2
    Else
        WriteMemory Chr(&H74) & Chr(&H13), &H4E62EC, 2
    End If
End Sub

Sub DoTeamChange()
    If frmMain.chkTeam.Value = 1 Then
        WriteMemory Chr(&H90) & Chr(&H90), &H46DE02, 2
    Else
        WriteMemory Chr(&HEB) & Chr(&H2C), &H46DE02, 2
    End If
End Sub

