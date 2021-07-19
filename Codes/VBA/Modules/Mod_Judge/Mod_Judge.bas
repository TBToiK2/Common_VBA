Attribute VB_Name = "Mod_Common"
Option Explicit
'----------------------------------------------------------------------------------------------------
Public Const MIN_ROW = 1
Public Const MAX_ROW = 1048576
Public Const MIN_COL = 1
Public Const MAX_COL = 16384

'----------------------------------------------------------------------------------------------------
Public Function ShowInfoMsg(ByVal Prompt As String) As Long
On Error Resume Next

    'メッセージ 表示
    ShowInfoMsg = MsgBox(Prompt, vbOKOnly + vbInformation, "情報")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function ShowQuestionMsg(ByVal Prompt As String) As Long
On Error Resume Next

    'メッセージ 表示
    ShowInfoMsg = MsgBox(Prompt, vbOKCancel + vbQuestion, "確認")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub ShowErrMsg(ByVal Description As String, Optional ByVal Number As Long)
On Error Resume Next

    'メッセージプロンプト 設定
    Dim Prompt As String
    Prompt = "エラー内容:[" & vbCrLf & Description & vbCrLf & "]"
    If Number <> 0 Then Prompt = "エラー番号:[" & Number & "]" & vbCrLf & Prompt

    'メッセージ 表示
    Call MsgBox(Prompt, vbOKOnly + vbCritical, "エラー")

    Err.Clear

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub ShowAllData(ByVal ws As Worksheet)

    With ws
        'フィルター 解除
        If .FilterMode Then Call .ShowAllData

        '行列 全表示
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub BeforeProcess(ByRef Calculation As XlCalculation, Optional ByRef ExcelApp As Application)

    'オブジェクト有無 判定
    If IsBlank(ExcelApp) Then
        With Application
            Calculation = .Calculation
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
        End With
        DoEvents
    Else
        With ExcelApp
            If .Workbooks.Count > 0 Then
                Calculation = .Calculation
                .Calculation = xlCalculationManual
            End If
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
        End With
        DoEvents
    End If

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub AfterProcess(ByVal Calculation As XlCalculation, Optional ByRef ExcelApp As Application)

    'オブジェクト有無 判定
    If IsBlank(ExcelApp) Then
        With Excel.Application
            .Calculation = Calculation
            .CutCopyMode = False
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
        End With
        DoEvents
    Else
        With ExcelApp
            If .Workbooks.Count > 0 Then
                .Calculation = Calculation
            End If
            .CutCopyMode = False
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
        End With
        DoEvents
    End If

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function BackupFile(ByVal FileSpec As String) As Boolean
On Error GoTo Err

    Dim FSO As New FileSystemObject
    Dim BaseName As String, ExtensionName As String

    'ファイル名・拡張子 取得
    BaseName = FSO.GetBaseName(FileSpec)
    ExtensionName = FSO.GetExtensionName(FileSpec)

    Call FileCopy(FileSpec, ThisWorkbook.Path & "\" & BaseName & "_" & CStr(Format(Now(), "yymmddhhmmss")) & "." & ExtensionName)

    BackupFile = True

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg("ファイルのコピー(バックアップ)が失敗しました。" & vbCrLf & Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------
