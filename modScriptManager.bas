Attribute VB_Name = "modScriptManager"
Option Explicit

' Windows API 声明
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hwnd As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
     ByVal lpsz2 As String) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 进程快照相关API
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
    (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32" _
    (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" _
    (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As Long

' 常量定义
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Private Const STARTF_USESHOWWINDOW = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const TH32CS_SNAPPROCESS = &H2
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const PROCESS_TERMINATE = &H1

' PROCESSENTRY32 结构体定义
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

' 结构体定义
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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

' 脚本信息类型
Public Type ScriptInfo
    scriptPath As String
    ScriptName As String
    processId As Long      ' cmd.exe 的进程ID
    ProcessHandle As Long  ' cmd.exe 的进程句柄
    pythonProcessId As Long ' python.exe 的进程ID
    WindowHandle As Long
    IsRunning As Boolean
End Type

' 全局变量
Public Scripts() As ScriptInfo
Public ScriptCount As Integer

' 通过进程ID查找窗口
Public Function FindWindowByProcessId(ByVal processId As Long) As Long
    Dim hwnd As Long
    Dim windowProcessId As Long
    
    hwnd = FindWindowEx(0, 0, vbNullString, vbNullString)
    
    Do While hwnd <> 0
        GetWindowThreadProcessId hwnd, windowProcessId
        
        If windowProcessId = processId Then
            FindWindowByProcessId = hwnd
            Exit Function
        End If
        
        hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
    Loop
    
    FindWindowByProcessId = 0
End Function

' 查找Python子进程
Private Function FindPythonProcessId(ByVal parentProcessId As Long) As Long
    Dim hSnapshot As Long
    Dim pe As PROCESSENTRY32
    Dim ret As Long
    
    ' 创建进程快照
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnapshot = INVALID_HANDLE_VALUE Then
        FindPythonProcessId = 0
        Exit Function
    End If
    
    ' 初始化结构体大小
    pe.dwSize = Len(pe)
    ret = Process32First(hSnapshot, pe)
    
    Do While ret <> 0
        ' 查找python.exe进程
        Dim exeName As String
        exeName = Left(pe.szExeFile, InStr(pe.szExeFile, Chr(0)) - 1)
        
        If LCase(exeName) = "python.exe" Or LCase(exeName) = "python" Then
            ' 检查父进程ID
            If pe.th32ParentProcessID = parentProcessId Then
                FindPythonProcessId = pe.th32ProcessID
                CloseHandle hSnapshot
                Exit Function
            End If
        End If
        ret = Process32Next(hSnapshot, pe)
    Loop
    
    CloseHandle hSnapshot
    FindPythonProcessId = 0
End Function

' 更新单个脚本的信息
Public Sub UpdateScriptInfo(ByVal scriptIndex As Integer)
    If scriptIndex >= 0 And scriptIndex < ScriptCount Then
        If Scripts(scriptIndex).IsRunning Then
            ' 更新窗口句柄
            Scripts(scriptIndex).WindowHandle = FindWindowByProcessId(Scripts(scriptIndex).processId)
            
            ' 查找Python子进程
            If Scripts(scriptIndex).pythonProcessId = 0 Then
                Scripts(scriptIndex).pythonProcessId = FindPythonProcessId(Scripts(scriptIndex).processId)
            End If
        End If
    End If
End Sub

' 更新所有脚本的信息
Public Sub UpdateAllScriptsInfo()
    Dim i As Integer
    
    For i = 0 To ScriptCount - 1
        UpdateScriptInfo i
    Next i
End Sub

' 运行Python脚本 - 完整修复版
Public Function RunPythonScript(ByVal scriptPath As String, ByVal scriptIndex As Integer, _
                                Optional ByVal hideWindow As Boolean = False) As Boolean
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim cmdLine As String
    Dim result As Long
    Dim retry As Integer
    Dim scriptDir As String
    Dim scriptFile As String
    
    On Error GoTo ErrorHandler
    
    ' 获取脚本所在目录和文件名
    scriptDir = GetFileDirectory(scriptPath)
    scriptFile = GetFileName(scriptPath)
    
    ' 调试信息
    Debug.Print "启动脚本: " & scriptPath
    Debug.Print "脚本目录: " & scriptDir
    Debug.Print "脚本文件: " & scriptFile
    
    ' 构建命令行
    ' 方法1: 使用 cd /d 切换目录（兼容性好）
    cmdLine = "cmd.exe /c cd /d " & Chr(34) & scriptDir & Chr(34) & _
              " & title Python-" & Scripts(scriptIndex).ScriptName & _
              " & python " & Chr(34) & scriptFile & Chr(34)
    
    ' 初始化结构体
    si.cb = Len(si)
    
    If hideWindow Then
        si.dwFlags = STARTF_USESHOWWINDOW
        si.wShowWindow = SW_HIDE
    End If
    
    ' 创建进程 - 设置工作目录为脚本所在目录
    result = CreateProcess(vbNullString, cmdLine, 0, 0, 0, _
                          NORMAL_PRIORITY_CLASS, 0, scriptDir, si, pi)
    
    If result <> 0 Then
        ' 保存进程信息
        Scripts(scriptIndex).processId = pi.dwProcessId
        Scripts(scriptIndex).ProcessHandle = pi.hProcess
        Scripts(scriptIndex).IsRunning = True
        Scripts(scriptIndex).pythonProcessId = 0
        
        ' 关闭线程句柄
        CloseHandle pi.hThread
        
        ' 等待窗口创建并获取句柄
        For retry = 1 To 10
            Sleep 300
            UpdateScriptInfo scriptIndex
            If Scripts(scriptIndex).WindowHandle <> 0 Then
                Exit For
            End If
        Next retry
        
        RunPythonScript = True
        Debug.Print "启动成功，进程ID: " & Scripts(scriptIndex).processId
    Else
        RunPythonScript = False
        Debug.Print "启动失败，错误码: " & Err.LastDllError
        MsgBox "启动脚本失败！错误码: " & Err.LastDllError, vbExclamation, "错误"
    End If
    
    Exit Function
    
ErrorHandler:
    RunPythonScript = False
    Debug.Print "启动异常: " & Err.Description
End Function

' 获取文件所在目录
Public Function GetFileDirectory(ByVal fullPath As String) As String
    Dim pos As Integer
    
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileDirectory = Left(fullPath, pos - 1)
    Else
        GetFileDirectory = ""
    End If
End Function

' 停止Python脚本 - 同时结束cmd和python进程
Public Function StopPythonScript(ByVal scriptIndex As Integer) As Boolean
    Dim pythonProcessId As Long
    Dim pythonHandle As Long
    
    On Error GoTo ErrorHandler
    
    If Scripts(scriptIndex).IsRunning Then
        ' 1. 先结束Python进程
        pythonProcessId = Scripts(scriptIndex).pythonProcessId
        
        If pythonProcessId <> 0 Then
            ' 打开Python进程
            pythonHandle = OpenProcess(PROCESS_TERMINATE, 0, pythonProcessId)
            If pythonHandle <> 0 Then
                TerminateProcess pythonHandle, 0
                CloseHandle pythonHandle
            End If
        End If
        
        ' 2. 结束CMD进程
        If Scripts(scriptIndex).ProcessHandle <> 0 Then
            TerminateProcess Scripts(scriptIndex).ProcessHandle, 0
            CloseHandle Scripts(scriptIndex).ProcessHandle
        End If
        
        ' 3. 清理信息
        Scripts(scriptIndex).IsRunning = False
        Scripts(scriptIndex).ProcessHandle = 0
        Scripts(scriptIndex).processId = 0
        Scripts(scriptIndex).pythonProcessId = 0
        Scripts(scriptIndex).WindowHandle = 0
        
        StopPythonScript = True
    End If
    
    Exit Function
    
ErrorHandler:
    StopPythonScript = False
End Function

' 隐藏脚本窗口
Public Function HideScriptWindow(ByVal scriptIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    If Scripts(scriptIndex).IsRunning Then
        ' 更新窗口句柄
        UpdateScriptInfo scriptIndex
        
        If Scripts(scriptIndex).WindowHandle <> 0 Then
            If ShowWindow(Scripts(scriptIndex).WindowHandle, SW_HIDE) <> 0 Then
                HideScriptWindow = True
                Exit Function
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    HideScriptWindow = False
End Function

' 显示脚本窗口
Public Function ShowScriptWindow(ByVal scriptIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    If Scripts(scriptIndex).IsRunning Then
        ' 更新窗口句柄
        UpdateScriptInfo scriptIndex
        
        If Scripts(scriptIndex).WindowHandle <> 0 Then
            If ShowWindow(Scripts(scriptIndex).WindowHandle, SW_SHOW) <> 0 Then
                ShowScriptWindow = True
                Exit Function
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ShowScriptWindow = False
End Function

' 添加脚本到列表
Public Sub AddScript(ByVal scriptPath As String)
    Dim i As Integer
    
    ' 检查是否已存在
    For i = 0 To ScriptCount - 1
        If Scripts(i).scriptPath = scriptPath Then
            MsgBox "脚本已存在！", vbExclamation, "提示"
            Exit Sub
        End If
    Next i
    
    ' 扩展数组
    ReDim Preserve Scripts(ScriptCount)
    
    ' 添加新脚本
    Scripts(ScriptCount).scriptPath = scriptPath
    Scripts(ScriptCount).ScriptName = GetFileName(scriptPath)
    Scripts(ScriptCount).IsRunning = False
    Scripts(ScriptCount).processId = 0
    Scripts(ScriptCount).ProcessHandle = 0
    Scripts(ScriptCount).pythonProcessId = 0
    Scripts(ScriptCount).WindowHandle = 0
    
    ScriptCount = ScriptCount + 1
    
    ' 刷新列表显示
    RefreshScriptList
End Sub

' 删除脚本
Public Sub DeleteScript(ByVal index As Integer)
    Dim i As Integer
    
    If index >= 0 And index < ScriptCount Then
        If Scripts(index).IsRunning Then
            StopPythonScript index
        End If
        
        For i = index To ScriptCount - 2
            Scripts(i) = Scripts(i + 1)
        Next i
        
        ScriptCount = ScriptCount - 1
        ReDim Preserve Scripts(ScriptCount)
        
        RefreshScriptList
    End If
End Sub

' 刷新脚本列表显示
Public Sub RefreshScriptList()
    Dim i As Integer
    Dim displayText As String
    
    On Error Resume Next
    frmMain.lstScripts.Clear
    
    For i = 0 To ScriptCount - 1
        displayText = Scripts(i).ScriptName & " - "
        
        If Scripts(i).IsRunning Then
            displayText = displayText & "[运行中]"
        Else
            displayText = displayText & "[已停止]"
        End If
        
        frmMain.lstScripts.AddItem displayText
    Next i
    On Error GoTo 0
End Sub

' 更新状态显示
Public Sub UpdateStatus(ByVal message As String)
    On Error Resume Next
    frmMain.lblStatus.Caption = message
    frmMain.lblStatus.Refresh
    On Error GoTo 0
End Sub

' 获取文件名
Public Function GetFileName(ByVal fullPath As String) As String
    Dim pos As Integer
    
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function

' 检查进程是否还在运行
Public Sub CheckProcessStatus()
    Dim i As Integer
    
    For i = 0 To ScriptCount - 1
        If Scripts(i).IsRunning Then
            If Scripts(i).ProcessHandle <> 0 Then
                If WaitForSingleObject(Scripts(i).ProcessHandle, 0) = 0 Then
                    ' 进程已结束
                    Scripts(i).IsRunning = False
                    CloseHandle Scripts(i).ProcessHandle
                    Scripts(i).ProcessHandle = 0
                    Scripts(i).processId = 0
                    Scripts(i).pythonProcessId = 0
                    Scripts(i).WindowHandle = 0
                End If
            End If
        End If
    Next i
    
    RefreshScriptList
End Sub

' 启动所有脚本
Public Sub StartAllScripts()
    Dim i As Integer
    Dim successCount As Integer
    
    successCount = 0
    UpdateStatus "正在启动所有脚本..."
    
    For i = 0 To ScriptCount - 1
        If Not Scripts(i).IsRunning Then
            If RunPythonScript(Scripts(i).scriptPath, i, False) Then
                successCount = successCount + 1
                DoEvents
                Sleep 500
            End If
        End If
    Next i
    
    RefreshScriptList
    UpdateStatus "已启动 " & successCount & " 个脚本"
End Sub

' 停止所有脚本
Public Sub StopAllScripts()
    Dim i As Integer
    Dim stopCount As Integer
    
    If MsgBox("确定要停止所有正在运行的脚本吗？", vbYesNo + vbQuestion, "确认停止") = vbYes Then
        stopCount = 0
        UpdateStatus "正在停止所有脚本..."
        
        For i = 0 To ScriptCount - 1
            If Scripts(i).IsRunning Then
                If StopPythonScript(i) Then
                    stopCount = stopCount + 1
                End If
                DoEvents
            End If
        Next i
        
        RefreshScriptList
        UpdateStatus "已停止 " & stopCount & " 个脚本"
    End If
End Sub

' 隐藏所有窗口
Public Sub HideAllWindows()
    Dim i As Integer
    Dim hideCount As Integer
    
    hideCount = 0
    UpdateStatus "正在隐藏所有窗口..."
    
    ' 先更新所有信息
    UpdateAllScriptsInfo
    
    For i = 0 To ScriptCount - 1
        If Scripts(i).IsRunning Then
            If HideScriptWindow(i) Then
                hideCount = hideCount + 1
            End If
            DoEvents
        End If
    Next i
    
    UpdateStatus "已隐藏 " & hideCount & " 个窗口"
End Sub

' 显示所有窗口
Public Sub ShowAllWindows()
    Dim i As Integer
    Dim showCount As Integer
    
    showCount = 0
    UpdateStatus "正在显示所有窗口..."
    
    ' 先更新所有信息
    UpdateAllScriptsInfo
    
    For i = 0 To ScriptCount - 1
        If Scripts(i).IsRunning Then
            If ShowScriptWindow(i) Then
                showCount = showCount + 1
            End If
            DoEvents
        End If
    Next i
    
    UpdateStatus "已显示 " & showCount & " 个窗口"
End Sub


