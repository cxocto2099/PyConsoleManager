VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Python脚本管理器 - 简化版"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9000
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdProperties 
      Caption         =   "属性"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "显示窗口"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隐藏窗口"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "停止所选"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "启动所选"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "全部显示"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdHideAll 
      Caption         =   "全部隐藏"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStopAll 
      Caption         =   "全部停止"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartAll 
      Caption         =   "全部启动"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除脚本"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加脚本"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstScripts 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8775
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' 初始化
    ScriptCount = 0
    ReDim Scripts(0)
    UpdateStatus "就绪"
    lblStatus.BackColor = &H80000018
    
    ' 加载保存的脚本列表
    LoadScriptList
    CheckProcessStatus  ' 使用 CheckProcessStatus 替代 RefreshAllStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 关闭所有运行的脚本
    Dim i As Integer
    
    For i = 0 To ScriptCount - 1
        If Scripts(i).IsRunning Then
            StopPythonScript i
        End If
    Next i
    
    ' 保存脚本列表
    SaveScriptList
End Sub

' 添加脚本
Private Sub cmdAdd_Click()
    Dim commonDialog As Object
    Dim filePath As String
    
    On Error Resume Next
    Set commonDialog = CreateObject("MSComDlg.CommonDialog")
    
    If Not commonDialog Is Nothing Then
        commonDialog.Filter = "Python脚本 (*.py)|*.py|所有文件 (*.*)|*.*"
        commonDialog.ShowOpen
        
        If commonDialog.FileName <> "" Then
            AddScript commonDialog.FileName
            UpdateStatus "已添加脚本：" & GetFileName(commonDialog.FileName)
        End If
    Else
        ' 如果没有CommonDialog控件，使用InputBox
        filePath = InputBox("请输入Python脚本的完整路径：", "添加脚本")
        If filePath <> "" Then
            If Dir(filePath) <> "" Then
                AddScript filePath
                UpdateStatus "已添加脚本：" & GetFileName(filePath)
            Else
                MsgBox "文件不存在！", vbExclamation, "错误"
            End If
        End If
    End If
End Sub

' 删除脚本
Private Sub cmdDelete_Click()
    Dim selectedIndex As Integer
    
    selectedIndex = lstScripts.ListIndex
    If selectedIndex >= 0 Then
        If MsgBox("确定要删除脚本 " & Scripts(selectedIndex).ScriptName & " 吗？", _
                  vbYesNo + vbQuestion, "确认删除") = vbYes Then
            DeleteScript selectedIndex
            UpdateStatus "已删除脚本"
        End If
    Else
        MsgBox "请先选择要删除的脚本！", vbExclamation, "提示"
    End If
End Sub

' 启动所选脚本
Private Sub cmdStart_Click()
    Dim selectedIndex As Integer
    selectedIndex = lstScripts.ListIndex
    
    If selectedIndex >= 0 Then
        If Not Scripts(selectedIndex).IsRunning Then
            If RunPythonScript(Scripts(selectedIndex).scriptPath, selectedIndex, False) Then
                UpdateStatus "已启动脚本：" & Scripts(selectedIndex).ScriptName
                RefreshScriptList
            Else
                UpdateStatus "启动失败：" & Scripts(selectedIndex).ScriptName
                MsgBox "启动脚本失败！请确保：" & vbCrLf & _
                       "1. Python已正确安装（python.exe在系统PATH中）" & vbCrLf & _
                       "2. 脚本路径正确", vbExclamation, "错误"
            End If
        Else
            MsgBox "脚本已在运行中！", vbInformation, "提示"
        End If
    Else
        MsgBox "请先选择要启动的脚本！", vbExclamation, "提示"
    End If
End Sub

' 停止所选脚本
Private Sub cmdStop_Click()
    Dim selectedIndex As Integer
    selectedIndex = lstScripts.ListIndex
    
    If selectedIndex >= 0 Then
        If Scripts(selectedIndex).IsRunning Then
            If StopPythonScript(selectedIndex) Then
                UpdateStatus "已停止脚本：" & Scripts(selectedIndex).ScriptName
                RefreshScriptList
            End If
        Else
            MsgBox "脚本未运行！", vbInformation, "提示"
        End If
    Else
        MsgBox "请先选择要停止的脚本！", vbExclamation, "提示"
    End If
End Sub

' 隐藏所选脚本窗口
Private Sub cmdHide_Click()
    Dim selectedIndex As Integer
    selectedIndex = lstScripts.ListIndex
    
    If selectedIndex >= 0 Then
        If Scripts(selectedIndex).IsRunning Then
            If HideScriptWindow(selectedIndex) Then
                UpdateStatus "已隐藏窗口：" & Scripts(selectedIndex).ScriptName
            Else
                UpdateStatus "隐藏窗口失败"
                MsgBox "无法找到脚本窗口！", vbExclamation, "提示"
            End If
        Else
            MsgBox "脚本未运行，无法隐藏窗口！", vbInformation, "提示"
        End If
    Else
        MsgBox "请先选择脚本！", vbExclamation, "提示"
    End If
End Sub

' 显示所选脚本窗口
Private Sub cmdShow_Click()
    Dim selectedIndex As Integer
    selectedIndex = lstScripts.ListIndex
    
    If selectedIndex >= 0 Then
        If Scripts(selectedIndex).IsRunning Then
            If ShowScriptWindow(selectedIndex) Then
                UpdateStatus "已显示窗口：" & Scripts(selectedIndex).ScriptName
            Else
                UpdateStatus "显示窗口失败"
                'MsgBox "无法找到脚本窗口！", vbExclamation, "提示"
            End If
        Else
            MsgBox "脚本未运行，无法显示窗口！", vbInformation, "提示"
        End If
    Else
        MsgBox "请先选择脚本！", vbExclamation, "提示"
    End If
End Sub

' 查看脚本属性
Private Sub cmdProperties_Click()
    Dim selectedIndex As Integer
    selectedIndex = lstScripts.ListIndex
    
    If selectedIndex >= 0 Then
        Dim statusText As String
        statusText = IIf(Scripts(selectedIndex).IsRunning, "运行中", "已停止")
        
        MsgBox "【脚本属性】" & vbCrLf & vbCrLf & _
               "脚本名称：" & Scripts(selectedIndex).ScriptName & vbCrLf & _
               "完整路径：" & Scripts(selectedIndex).scriptPath & vbCrLf & _
               "运行状态：" & statusText & vbCrLf & _
               "CMD进程ID：" & IIf(Scripts(selectedIndex).processId > 0, Scripts(selectedIndex).processId, "无") & vbCrLf & _
               "Python进程ID：" & IIf(Scripts(selectedIndex).pythonProcessId > 0, Scripts(selectedIndex).pythonProcessId, "无") & vbCrLf & _
               "窗口句柄：" & IIf(Scripts(selectedIndex).WindowHandle > 0, Hex(Scripts(selectedIndex).WindowHandle), "无"), _
               vbInformation, "脚本属性"
    Else
        MsgBox "请先选择脚本！", vbExclamation, "提示"
    End If
End Sub

' 全部启动
Private Sub cmdStartAll_Click()
    StartAllScripts
End Sub

' 全部停止
Private Sub cmdStopAll_Click()
    StopAllScripts
End Sub

' 全部隐藏
Private Sub cmdHideAll_Click()
    HideAllWindows
End Sub

' 全部显示
Private Sub cmdShowAll_Click()
    ShowAllWindows
End Sub

' 双击列表项启动/停止
Private Sub lstScripts_DblClick()
    Dim selectedIndex As Integer
    
    selectedIndex = lstScripts.ListIndex
    If selectedIndex >= 0 Then
        If Scripts(selectedIndex).IsRunning Then
            If StopPythonScript(selectedIndex) Then
                UpdateStatus "已停止脚本：" & Scripts(selectedIndex).ScriptName
            End If
        Else
            If RunPythonScript(Scripts(selectedIndex).scriptPath, selectedIndex, False) Then
                UpdateStatus "已启动脚本：" & Scripts(selectedIndex).ScriptName
            Else
                UpdateStatus "启动失败：" & Scripts(selectedIndex).ScriptName
            End If
        End If
        RefreshScriptList
    End If
End Sub

' 保存脚本列表到文件
Private Sub SaveScriptList()
    Dim fileNum As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open App.Path & "\scripts.dat" For Output As #fileNum
    
    Write #fileNum, ScriptCount
    
    For i = 0 To ScriptCount - 1
        Write #fileNum, Scripts(i).scriptPath
    Next i
    
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    If fileNum <> 0 Then Close #fileNum
End Sub

' 加载脚本列表
Private Sub LoadScriptList()
    Dim fileNum As Integer
    Dim count As Integer
    Dim i As Integer
    Dim scriptPath As String
    
    On Error GoTo ErrorHandler
    
    If Dir(App.Path & "\scripts.dat") = "" Then Exit Sub
    
    fileNum = FreeFile
    Open App.Path & "\scripts.dat" For Input As #fileNum
    
    Input #fileNum, count
    
    For i = 1 To count
        Input #fileNum, scriptPath
        If Dir(scriptPath) <> "" Then
            AddScript scriptPath
        End If
    Next i
    
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    If fileNum <> 0 Then Close #fileNum
End Sub

' 可选：添加一个定时器来定期检查进程状态
' 如果需要在窗体上添加 Timer 控件，取消下面代码的注释
'Private Sub Timer1_Timer()
'    Static lastCheck As Date
'
'    ' 每3秒检查一次
'    If DateDiff("s", lastCheck, Now) >= 3 Then
'        CheckProcessStatus
'        lastCheck = Now
'    End If
'End Sub

