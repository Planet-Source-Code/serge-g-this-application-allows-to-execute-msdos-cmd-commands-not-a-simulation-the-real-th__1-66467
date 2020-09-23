VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command Line"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Default         =   -1  'True
      DownPicture     =   "frmMain.frx":08CA
      Height          =   375
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":2F90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   5895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Redirects output from console program to textbox.
'Requires two textboxes and one command button.
'Set MultiLine property of Text2 to true.
'
'Note: don't run plain DOS programs with this example
'under Windows 95,98 and ME, as the program freezes when
'execution of program is finnished.
'
'Thanx again to AllAPI.net

Option Explicit
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_HIDE = 0
Private Const EM_SETSEL = &HB1
Private Const EM_REPLACESEL = &HC2


Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbEnter Then
        Command1_Click
    Else
        proccessKeyUp (KeyCode)
    End If
    
End Sub

Private Sub Command1_Click()

    Text1.Text = Combo1.Text
    
    If Text1.Text = "" Then Exit Sub
    If LCase(Text1.Text) = "cls" Then Text2.Text = "": Exit Sub
    Command1.Enabled = False
    Redirect Text1.Text, Text3
    Text2.Text = Text2.Text & Text3.Text
    Text3.Text = ""
    Command1.Enabled = True
    
'    If Text1.Text = "" Then Exit Sub
'    If LCase(Text1.Text) = "cls" Then Text2.Text = "": Exit Sub
'    Command1.Enabled = False
'    Redirect Text1.Text, Text3
'    Text2.Text = Text2.Text & Text3.Text
'    Text3.Text = ""
'    Command1.Enabled = True

    Combo1.AddItem Text1.Text
    
    Text2.SelStart = Len(Text2.Text)
    Combo1.Text = ""
    
End Sub
Private Sub Form_Load()

    Text1.Text = "help"
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Command1.Enabled = False Then Cancel = True
    
End Sub

Sub Redirect(cmdLine As String, objTarget As Object)

    Dim i%, t$
    Dim pa As SECURITY_ATTRIBUTES
    Dim pra As SECURITY_ATTRIBUTES
    Dim tra As SECURITY_ATTRIBUTES
    Dim pi As PROCESS_INFORMATION
    Dim sui As STARTUPINFO
    Dim hRead As Long
    Dim hWrite As Long
    Dim bRead As Long
    Dim lpBuffer(1024) As Byte
    pa.nLength = Len(pa)
    pa.lpSecurityDescriptor = 0
    pa.bInheritHandle = True
    
    pra.nLength = Len(pra)
    tra.nLength = Len(tra)

    If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
        sui.cb = Len(sui)
        GetStartupInfo sui
        sui.hStdOutput = hWrite
        sui.hStdError = hWrite
        sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        sui.wShowWindow = SW_HIDE
        If CreateProcess(vbNullString, cmdLine, pra, tra, True, 0, Null, vbNullString, sui, pi) <> 0 Then
            SetWindowText objTarget.hwnd, ""
            Do
                Erase lpBuffer()
                If ReadFile(hRead, lpBuffer(0), 1023, bRead, ByVal 0&) Then
                    SendMessage objTarget.hwnd, EM_SETSEL, -1, 0
                    SendMessage objTarget.hwnd, EM_REPLACESEL, False, lpBuffer(0)
                    DoEvents
                Else
                    CloseHandle pi.hThread
                    CloseHandle pi.hProcess
                    Exit Do
                End If
                CloseHandle hWrite
            Loop
            CloseHandle hRead
        End If
    End If
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbEnter Then
        Command1_Click
    End If
    
End Sub

Private Sub proccessKeyUp(KeyCode As Integer)

    Dim currLen As Integer
    Dim currText As String
    Dim x As Integer
        
    If KeyCode = vbKeySpace Then GoTo allow
    If KeyCode > 47 Or KeyCode < 1 Then
allow:
        currLen = Len(Combo1.Text)
        For x = 0 To Combo1.ListCount - 1
            If Left(LCase(Combo1.Text), currLen) = Left(LCase(Combo1.List(x)), currLen) Then
                Combo1.ListIndex = x
                Combo1.SelStart = currLen
                Combo1.SelLength = Len(Combo1.Text)
                Exit For
            End If
        Next x
    End If
    
End Sub
