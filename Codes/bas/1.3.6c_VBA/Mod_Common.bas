Attribute VB_Name = "Mod_Common"
'1.3.6c_VBA
Option Explicit
Option Private Module
'----------------------------------------------------------------------------------------------------
'2025/03/18 04:27:44
'----------------------------------------------------------------------------------------------------
Public Const MIN_ROW = 1
Public Const MAX_ROW = 1048576
Public Const MIN_COL = 1
Public Const MAX_COL = 16384
Public FSO As New FileSystemObject
Public REG As New RegExp
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 15:31:21
'----------------------------------------------------------------------------------------------------
Public Sub AfterProcess(Optional ByVal calculation As XlCalculation, Optional ByRef excelApp As Excel.Application, Optional ByVal isWbOpening As Boolean)

    '引数 既定値判定
    If excelApp Is Nothing Then Set excelApp = Excel.Application

    '各プロセス 再開
    With excelApp
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        If Not isWbOpening Then
            .CutCopyMode = False
            .StatusBar = False
        End If
        If .Workbooks.Count > 0 Then
            .Calculation = calculation
        End If
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 15:31:21
'----------------------------------------------------------------------------------------------------
Public Sub BeforeProcess(Optional ByRef calculation As XlCalculation, Optional ByRef excelApp As Excel.Application, Optional ByVal isWbOpening As Boolean)

    '引数 既定値判定
    If excelApp Is Nothing Then Set excelApp = Excel.Application

    '各プロセス 停止
    With excelApp
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        If Not isWbOpening Then
            .StatusBar = False
        End If
        If .Workbooks.Count > 0 Then
            calculation = .Calculation
            .Calculation = xlCalculationManual
        End If
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/12 15:12:20
'----------------------------------------------------------------------------------------------------
Public Sub ShowAllData(ByRef targetWs As Worksheet)
On Error Resume Next

    With targetWs
        'フィルター 解除
        If .FilterMode Then Call .ShowAllData

        '行列 全表示
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/17 09:59:52
'----------------------------------------------------------------------------------------------------
Public Sub ShowErrMsg(ByVal errDescription As String, Optional ByVal errNumber As Long, Optional ByVal title As String)
On Error Resume Next

    'メッセージプロンプト 設定
    Dim prompt As String
    prompt = "エラー内容:[" & vbCrLf & errDescription & vbCrLf & "]"
    If errNumber <> 0 Then Prompt = "エラー番号:[" & errNumber & "]" & vbCrLf & prompt

    'タイトル 設定
    If title <> "" Then title = ":" & title

    'メッセージ 表示
    Call MsgBox(prompt, vbOKOnly + vbCritical, "エラー" & title)

    Err.Clear

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/29 04:11:11
'----------------------------------------------------------------------------------------------------
Public Sub ShowInfoMsg(ByVal prompt As String, Optional ByVal title As String)
On Error Resume Next

    'タイトル 設定
    If title <> "" Then title = ":" & title

    'メッセージ 表示
    Call MsgBox(prompt, vbOKOnly + vbInformation, "情報" & title)

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/29 04:11:11
'----------------------------------------------------------------------------------------------------
Public Function ShowQuestionMsg(ByVal prompt As String, Optional ByVal title As String) As VbMsgBoxResult
On Error Resume Next

    'タイトル 設定
    If title <> "" Then title = ":" & title

    'メッセージ 表示
    ShowQuestionMsg = MsgBox(prompt, vbOKCancel + vbQuestion, "確認" & title)

End Function
'----------------------------------------------------------------------------------------------------