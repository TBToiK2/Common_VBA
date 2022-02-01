Attribute VB_Name = "Mod_Get"
Option Explicit
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function GetFileName(ByVal FileSpec As String) As String
On Error GoTo Err

    'ファイル名 取得
    Dim FSO As New FileSystemObject
    GetFileName = FSO.GetFileName(FileSpec)

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function GetPath(ByVal Spec As String) As String
On Error GoTo Err

    'ファイル名 取得
    Dim FSO As New FileSystemObject
    GetPath = FSO.GetParentFolderName(Spec)

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function GetSelectFileSpec() As String
On Error GoTo Err

    'ダイアログボックス 設定
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        With .Filters
            .Clear
            Call .Add("Excel ファイル", "*.xls*", 1)
        End With
        .FilterIndex = 1
        .Title = "Excel ファイル 選択"
        .InitialFileName = CurDir()

        'フォルダ選択状態 判定
        If .Show = -1 Then GetSelectFileSpec = .SelectedItems(1)
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function GetSelectFolderSpec() As String
On Error GoTo Err

    'ダイアログボックス 設定
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "フォルダ 選択"

        'フォルダ選択状態 判定
        If .Show = -1 Then GetSelectFolderSpec = .SelectedItems(1)
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------
