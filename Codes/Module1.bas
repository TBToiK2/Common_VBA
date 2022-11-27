Attribute VB_Name = "Module1"
Option Explicit
'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Sub ReplaceModule()
On Error GoTo Err:

    '各プロセス 停止
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .StatusBar = False
    End With

    'モジュール置換対象ファイル, フォルダディレクトリ 取得
    Dim targetFileSpec As String
    Dim targetFolderSpec As String
    If Not SelectFileFolderSpec(targetFileSpec, targetFolderSpec) Then Exit Sub

    If IsOpen(targetFileSpec) = -1 Then
        Call MsgBox("既に選択したファイルが開いているため、処理を中止します。" & vbLf & "閉じてからやり直してください。")
        Exit Sub
    End If

    'モジュールディクショナリ 作成
    Dim moduleFileDIC As Dictionary
    Set moduleFileDIC = GetModuleFileDIC(targetFolderSpec)

    If moduleFileDIC.Count = 0 Then Exit Sub

    'モジュール置換対象エクセルプロセス 停止
    Dim ExcelApp As New Excel.Application
    With ExcelApp
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .StatusBar = False

        Dim wb As Workbook
        Set wb = .Workbooks.Open(targetFileSpec)
    End With

    Dim VBCs As VBComponents
    Set VBCs = wb.VBProject.VBComponents

    '置換対象モジュール 絞り込み
    Dim VBC As VBComponent
    For Each VBC In VBCs
        If VBC.Type = vbext_ct_StdModule And moduleFileDIC.Exists(VBC.Name) Then
            Dim replaceVBCDIC As New Dictionary
            Call replaceVBCDIC.Add(VBC, moduleFileDIC(VBC.Name))
        End If
    Next VBC

    'モジュール 置換
    Dim VBCDICKey As Variant
    For Each VBCDICKey In replaceVBCDIC.Keys
        Call VBCs.Remove(VBCDICKey)
        Call VBCs.Import(replaceVBCDIC(VBCDICKey))
    Next VBCDICKey

    Call wb.Save

    Call MsgBox("置換完了")

Err:

    If Not wb Is Nothing Then Call wb.Close(False)

    'モジュール置換対象エクセルプロセス 再開
    With ExcelApp
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False

        Call .Quit
    End With

    '各プロセス 再開
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Function SelectFileFolderSpec(ByRef targetFileSpec As String, ByRef targetFolderSpec As String) As Boolean

    Dim FSO As New FileSystemObject

    Dim wbPath As String
    wbPath = ReplaceOneDrivePath(ThisWorkbook.path)

    'モジュール置換対象ファイル 選択
    With Application.FileDialog(msoFileDialogFilePicker)
        With .Filters
            Call .Clear
            Call .Add("Excel ファイル", "*.xlsm, *.xlam")
        End With
        .InitialFileName = FSO.BuildPath(wbPath, Application.PathSeparator)
        .FilterIndex = 1
        .AllowMultiSelect = False
        .Title = "モジュール置換対象ファイル 選択"
        .InitialFileName = vbNullString

        If .Show = -1 Then
            targetFileSpec = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With

    'モジュール格納フォルダ 選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "モジュール格納フォルダ 選択"
        .InitialFileName = FSO.BuildPath(wbPath, Application.PathSeparator)

        If .Show = -1 Then
            targetFolderSpec = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With

    SelectFileFolderSpec = True

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Function GetModuleFileDIC(ByVal targetFolderSpec As String) As Dictionary

    Dim FSO As New FileSystemObject
    Dim moduleFileDIC As New Dictionary

    Dim moduleFile As File
    For Each moduleFile In FSO.GetFolder(targetFolderSpec).Files
        Dim elm As Long
        If FSO.GetExtensionName(moduleFile.Name) = "bas" Then
            Call moduleFileDIC.Add(FSO.GetBaseName(moduleFile.Name), moduleFile.path)
        End If
    Next moduleFile

    Set GetModuleFileDIC = moduleFileDIC

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Function IsOpen(ByVal fileSpec As String) As Long
On Error Resume Next

    IsOpen = 1

    Dim fileNumber As Long
    'ファイルナンバー1-255
    fileNumber = FreeFile(0)
    'ファイルナンバー256-512
    If fileNumber = 0 Then fileNumber = FreeFile(1)

    'ファイルナンバー 全使用時
    If Err.Number = 67 Then Exit Function

    '開閉確認(Append)
    Open fileSpec For Append As #fileNumber
    Close #fileNumber

    '読み取り専用判定
    If Err.Number = 75 Then
        '開閉確認(Input)
        Open fileSpec For Input As #fileNumber
        Close #fileNumber
    End If

    IsOpen = Err.Number = 55 Or Err.Number = 70

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Function ReplaceOneDrivePath(ByVal path As String)

    'OneDriveパス 判定
    If Not path Like "https://*" Then
        ReplaceOneDrivePath = path
        Exit Function
    End If

    '環境設定 取得
    Dim OneDriveCommercialPath As String, OneDriveConsumerPath As String
    OneDriveCommercialPath = IIf(Environ("OneDriveCommercial") = "", Environ("OneDrive"), Environ("OneDriveCommercial"))
    OneDriveConsumerPath = IIf(Environ("OneDriveConsumer") = "", Environ("OneDrive"), Environ("OneDriveConsumer"))

    Dim filePathPos As Long
    '法人向け
    If path Like "*my.sharepoint.com*" Then

        filePathPos = InStr(path, "/Documents") + Len("/Documents")

        ReplaceOneDrivePath = OneDriveCommercialPath & Replace(Mid(path, filePathPos), "/", Application.PathSeparator)

    '個人向け
    Else

        filePathPos = InStr(Len("https://") + 1, path, "/")
        filePathPos = InStr(filePathPos + 1, path, "/")

        If filePathPos = 0 Then
            ReplaceOneDrivePath = OneDriveConsumerPath
        Else
            ReplaceOneDrivePath = OneDriveConsumerPath & Replace(Mid(path, filePathPos), "/", Application.PathSeparator)
        End If

    End If

End Function