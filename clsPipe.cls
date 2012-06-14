VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
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
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type PIPE
    hReadPipe As Long
    hWritePipe As Long
End Type
Private Const INFINITE = -1&

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const CREATE_NEW_CONSOLE = &H10

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESIZE = &H2
Private Const STARTF_USEPOSITION = &H4
Private Const STARTF_USECOUNTCHARS = &H8
Private Const STARTF_USEFILLATTRIBUTE = &H10
Private Const STARTF_RUNFULLSCREEN = &H20
Private Const STARTF_FORCEONFEEDBACK = &H40
Private Const STARTF_FORCEOFFFEEDBACK = &H80
Private Const STARTF_USESTDHANDLES = &H100

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Const INVALID_HANDLE_VALUE = -1
Private Const STILL_ACTIVE = &H103&
Private Declare Function CreatePipe Lib "kernel32" _
 (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes _
 As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" _
 (ByVal hFile As Long, ByVal lpBuffer As String, _
 ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead _
 As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" _
 (ByVal hFile As Long, ByVal lpBuffer As String, _
 ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten _
 As Long, lpOverlapped As Any) As Long
Private Declare Function CreateProcess Lib "kernel32" _
 Alias "CreateProcessA" (ByVal lpApplicationName As Long, _
 ByVal lpCommandLine As String, lpProcessAttributes As Any, _
 lpThreadAttributes As Any, ByVal bInheritHandles As _
 Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment _
 As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo _
 As Any, lpProcessInformation As Any) As Long
Private Declare Function GetExitCodeProcess Lib _
 "kernel32" (ByVal hProcess As Long, lpExitCode _
 As Long) As Long
Private Declare Function TerminateProcess Lib _
 "kernel32" (ByVal hProcess As Long, ByVal uExitCode _
 As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" _
 (ByVal nStdHandle As Long) As Long
Private Declare Function SetStdHandle Lib "kernel32" _
 (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
 (ByVal hObject As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" _
 (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal _
 nBufferSize As Long, lpBytesRead As Long, _
 lpTotalBytesAvail As Long, lpBytesLeftThisMessage _
 As Long) As Long
Private pipeOut As PIPE, pipeIn As PIPE
Private Process As PROCESS_INFORMATION
Private mvarCommandLine As String
Private mvarRunning As Boolean

Public Sub Terminate()
    TerminateProcess Process.hProcess, 0
    CloseHandle Process.hProcess
    CloseHandle Process.hThread
    CloseHandle pipeIn.hReadPipe
    CloseHandle pipeIn.hWritePipe
    CloseHandle pipeOut.hReadPipe
    CloseHandle pipeOut.hWritePipe
   
    mvarRunning = False
End Sub

Public Function Read(Optional ByVal Bytes As Long = -1) As String
    Dim tBytesR As Long, Buffer As String
    Dim tBytesA As Long, tMsg As Long
    Dim i As Long, result As Long
    Dim ReturnStr As String
    If Not mvarRunning Then Exit Function
    result = PeekNamedPipe(pipeOut.hReadPipe, ByVal 0&, 0, tBytesR, tBytesA, tMsg)
    If result <> 0 And tBytesA > 0 Then
        Buffer = String(tBytesA, " ")
        result = ReadFile(pipeOut.hReadPipe, Buffer, IIf(Bytes = -1, Len(Buffer), Bytes), tBytesR, ByVal 0&)
        If result = 0 Then Err.Raise vbObjectError + 504, "DOSShell Class", "Error: ReadFile failed. " & Err.LastDllError
        ReturnStr = Left(Buffer, tBytesR)
        Read = DOSDecode(ReturnStr)
    End If
End Function
Public Function Write_(ByVal Data As String) As Long
    Dim tBytesW As Long
    Dim i As Long, result As Long
    If Not Right(Data, 2) = Chr(13) & Chr(10) Then Data = Data & Chr(13) & Chr(10)
    result = WriteFile(pipeIn.hWritePipe, Data, Len(Data), tBytesW, ByVal 0&)
    If result = 0 Then Err.Raise 503, "DOSWrite", "Error: WriteFile failed. " & Err.LastDllError
    If result = 0 Then Err.Raise vbObjectError + 507, "DOSShell Class", "Error: FlushFileBuffers failed. " & Err.LastDllError
    Write_ = Len(Data) - 1
End Function

Public Function Execute(Optional ByVal CommandLine As String = "") As Long
    Dim result As Long
    Dim StartInfo As STARTUPINFO
    Dim Attribs As SECURITY_ATTRIBUTES
    Dim tIn As Long, tOut As Long
    On Error GoTo ErrHandler
    If CommandLine <> "" Then mvarCommandLine = CommandLine
    Attribs.nLength = Len(Attribs)
    Attribs.bInheritHandle = 1
    Attribs.lpSecurityDescriptor = 0&
    result = CreatePipe(pipeIn.hReadPipe, pipeIn.hWritePipe, _
 Attribs, ByVal 0&)
    If result = 0 Then _
        Err.Raise vbObjectError + 501, "DOSShell Class", _
 "Error: CreatePipe failed. " & Err.LastDllError
   
    result = CreatePipe(pipeOut.hReadPipe, pipeOut.hWritePipe, _
 Attribs, ByVal 0&)
    If result = 0 Then _
        Err.Raise vbObjectError + 501, "DOSShell Class", _
 "Error: CreatePipe failed. " & Err.LastDllError
       
    StartInfo.cb = Len(StartInfo)
    StartInfo.hStdInput = pipeIn.hReadPipe
    StartInfo.hStdOutput = pipeOut.hWritePipe
    StartInfo.hStdError = pipeOut.hWritePipe
    StartInfo.dwFlags = STARTF_USESTDHANDLES + _
                              STARTF_USESHOWWINDOW
    'StartInfo.wShowWindow = SW_SHOWNORMAL 'SW_HIDE
    StartInfo.wShowWindow = SW_HIDE
    result = CreateProcess(0&, mvarCommandLine, Attribs, _
 Attribs, ByVal 1&, CREATE_NEW_CONSOLE, ByVal 0&, ByVal _
 0&, StartInfo, Process)
    If result = 0 Then Err.Raise vbObjectError + 502, "DOSShell Class", "Error: CreateProcess failed. " & Err.LastDllError
    Execute = 1
    mvarRunning = True
    Exit Function
ErrHandler:
    Execute = Err.Number
End Function
Public Property Get Running() As Boolean
    Dim ExitCode As Long
    If Not mvarRunning Then
        Running = False
    Else
        GetExitCodeProcess Process.hProcess, ExitCode
        Running = (ExitCode = STILL_ACTIVE)
    End If
End Property
Public Property Let CommandLine(ByVal vData As String)
    mvarCommandLine = vData
End Property
Public Property Get CommandLine() As String
    CommandLine = mvarCommandLine
End Property
Private Function DOSDecode(ByVal Str As String) As String
    Dim i As Long
    For i = 239 To 192 Step -1
        Str = Replace(Str, Chr(i), Chr(i + 16))
    Next i
    For i = 191 To 128 Step -1
        Str = Replace(Str, Chr(i), Chr(i + 64))
    Next i
    Str = Replace(Str, Chr(0), "")
    DOSDecode = Str
End Function

