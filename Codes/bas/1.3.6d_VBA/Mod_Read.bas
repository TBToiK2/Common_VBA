Attribute VB_Name = "Mod_Read"
'1.3.6d_VBA
Option Explicit
Option Private Module
'----------------------------------------------------------------------------------------------------
'2025/03/18 04:27:44
'----------------------------------------------------------------------------------------------------
Public Enum EnumFileNameType
    enFileName = 0
    enBaseName = 1
    enExtensionName = 2
End Enum
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 03:29:12
'----------------------------------------------------------------------------------------------------
Public Function ArrayValueExists(ByRef targetArr As Variant, ByRef searchValue As Variant) As Boolean
On Error GoTo Err

    'ターゲット引数 配列判定
    If Not IsArray(targetArr) Then Exit Function

On Error Resume Next
    'ターゲット引数 空判定
    If LBound(targetArr, 1) > UBound(targetArr, 1) Then Exit Function
On Error GoTo Err

    '検索引数 配列, オブジェクト判定
    If IsArray(searchValue) Or IsObject(searchValue) Then Exit Function

    Dim elmValue As Variant
    '全値 確認
    For Each elmValue In targetArr
        '配列, オブジェクト 判定
        If Not IsArray(elmValue) And Not IsObject(elmValue) Then
            If searchValue = elmValue Then
                ArrayValueExists = True
                Exit Function
            End If
        End If
    Next elmValue

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ArrayValueExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function FileExists(ByVal fileSpec As String) As Boolean
On Error GoTo Err

    'ファイル 存在確認
    FileExists = FSO.FileExists(fileSpec)

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "FileExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function FolderExists(ByVal folderSpec As String) As Boolean
On Error GoTo Err

    'フォルダー 存在確認
    FolderExists = FSO.FolderExists(folderSpec)

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "FolderExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function GetColRange(ByRef targetWs As Worksheet, ByVal targetCol As Long, Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Variant, _
                            Optional ByVal relativeFLG As Boolean) As Range
On Error GoTo Err

    If targetCol < 1 Or targetCol > MAX_COL Then
        Call ShowErrMsg("targetColに指定された値が有効列数の範囲外です。" & "1以上" & MAX_COL & "以下で入力してください。", title:="GetColRange")
        Exit Function
    End If
    If firstRow < 1 Or firstRow > MAX_ROW Then
        Call ShowErrMsg("firstRowに指定された値が有効行数の範囲外です。" & "1以上" & MAX_ROW & "以下で入力してください。", title:="GetColRange")
        Exit Function
    End If

    Dim targetLastRow As Long
    If IsMissing(lastRow) Then
        '最終行 取得
        targetLastRow = GetLastRow(targetWs, targetCol)
        If firstRow > targetLastRow Then Exit Function
    Else
        If Not IsNumber(lastRow) Then
            Call ShowErrMsg("lastRowに指定された値は数値ではありません。" & "1以上" & MAX_ROW & "以下で入力してください。", title:="GetColRange")
            Exit Function
        Else
            If relativeFLG Then
                '最終行 取得
                targetLastRow = firstRow + lastRow
                If targetLastRow < 1 Or targetLastRow > MAX_ROW Then
                    Call ShowErrMsg("指定された値が有効行数の範囲外です。" & "firstRow + lastRowが" & "1以上" & MAX_ROW & "以下となるよう入力してください。", title:="GetColRange")
                    Exit Function
                End If
            Else
                '最終行 取得
                targetLastRow = lastRow
                If targetLastRow < 1 Or targetLastRow > MAX_ROW Then
                    Call ShowErrMsg("lastRowに指定された値が有効行数の範囲外です。" & "1以上" & MAX_ROW & "以下で入力してください。", title:="GetColRange")
                    Exit Function
                End If
            End If
        End If
    End If

    With targetWs
        Set GetColRange = .Range(.Cells(firstRow, targetCol), .Cells(targetLastRow, targetCol))
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetColRange")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetFileName(ByVal fileSpec As String, Optional ByVal enFileNameType As EnumFileNameType) As String
On Error GoTo Err

    With FSO
        'ファイル名 取得
        Select Case enFileNameType
            Case enFileName
                GetFileName = .GetFileName(fileSpec)
            Case enBaseName
                GetFileName = .GetBaseName(fileSpec)
            Case enExtensionName
                GetFileName = .GetExtensionName(fileSpec)
        End Select
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetFileName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function GetLastCol(ByRef targetWs As Worksheet, ParamArray targetRows() As Variant) As Long
On Error GoTo Err

    If IsMissing(targetRows) Then Exit Function

    Dim lastCol As Long

    With targetWs
        Dim usedLastCol As Long

        '使用範囲最終列 取得
        With .UsedRange
            usedLastCol = .Columns(.Columns.Count).Column
        End With

        '最終列 取得
        Dim targetRowsElm As Variant
        For Each targetRowsElm In targetRows
            If Not IsArray(targetRowsElm) Then targetRowsElm = Array(targetRowsElm)

            Dim targetRow As Variant
            For Each targetRow In targetRowsElm
                Dim usedRowRange As Range
                Set usedRowRange = .Range(.Cells(targetRow, 1), .Cells(targetRow, usedLastCol))

                Dim rangeIndx As Long
                For rangeIndx = usedRowRange.Count To 1 Step -1
                    Dim usedCell As Range
                    Set usedCell = usedRowRange.Item(rangeIndx)
                    If CStr(usedCell.Value) <> "" And usedCell.Column > lastCol Then
                        lastCol = usedCell.Column
                        Exit For
                    End If
                Next rangeIndx
            Next targetRow
        Next targetRowsElm
    End With

    GetLastCol = lastCol

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetLastCol")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function GetLastRow(ByRef targetWs As Worksheet, ParamArray targetCols() As Variant) As Long
On Error GoTo Err

    If IsMissing(targetCols) Then Exit Function

    Dim lastRow As Long

    With targetWs
        Dim usedLastRow As Long

        '使用範囲最終行 取得
        With .UsedRange
            usedLastRow = .Rows(.Rows.Count).Row
        End With

        '最終行 取得
        Dim targetColsElm As Variant
        For Each targetColsElm In targetCols
            If Not IsArray(targetColsElm) Then targetColsElm = Array(targetColsElm)

            Dim targetCol As Variant
            For Each targetCol In targetColsElm
                Dim usedColRange As Range
                Set usedColRange = .Range(.Cells(1, targetCol), .Cells(usedLastRow, targetCol))

                Dim rangeIndx As Long
                For rangeIndx = usedColRange.Count To 1 Step -1
                    Dim usedCell As Range
                    Set usedCell = usedColRange.Item(rangeIndx)
                    If CStr(usedCell.Value) <> "" And usedCell.Row > lastRow Then
                        lastRow = usedCell.Row
                        Exit For
                    End If
                Next rangeIndx
            Next targetCol
        Next targetColsElm
    End With

    GetLastRow = lastRow

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetLastRow")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function GetListObject(ByVal loName As String, Optional ByRef parent As ListObjects) As ListObject
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then
        'ワークシート 判定
        Dim actSh As Object
        Set actSh = ThisWorkbook.ActiveSheet
        If TypeName(actSh) = "Worksheet" Then
            If actSh.Type = xlWorksheet Then
                Set parent = actSh.ListObjects
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    '全リストオブジェクト 確認
    Dim lo As ListObject
    For Each lo In parent
        If lo.Name = loName Then
            Set GetListObject = lo
            Exit Function
        End If
    Next lo

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetListObject")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function GetName(ByVal nameName As String, Optional ByRef parent As Names) As Name
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then Set parent = ThisWorkbook.Names

    Dim defineName As Name
    'Worksheet 判定
    If TypeName(parent.Parent) = "Worksheet" Then
        '全名前 確認
        For Each defineName In parent
            If defineName.Name = parent.Parent.Name & "!" & nameName Then
                Set GetName = defineName
                Exit Function
            End If
        Next defineName
    Else
        '全名前 確認
        For Each defineName In parent
            If defineName.Name = nameName Then
                Set GetName = defineName
                Exit Function
            End If
        Next defineName
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetPath(ByVal spec As String) As String
On Error GoTo Err

    'パス 取得
    GetPath = FSO.GetParentFolderName(spec)

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetPath")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function GetRowRange(ByRef targetWs As Worksheet, ByVal targetRow As Long, Optional ByVal firstCol As Long = 1, Optional ByVal lastCol As Variant, _
                            Optional ByVal relativeFLG As Boolean) As Range
On Error GoTo Err

    If targetRow < 1 Or targetRow > MAX_ROW Then
        Call ShowErrMsg("targetRowに指定された値が有効行数の範囲外です。" & "1以上" & MAX_ROW & "以下で入力してください。", title:="GetRowRange")
        Exit Function
    End If
    If firstCol < 1 Or firstCol > MAX_COL Then
        Call ShowErrMsg("firstColに指定された値が有効列数の範囲外です。" & "1以上" & MAX_COL & "以下で入力してください。", title:="GetRowRange")
        Exit Function
    End If

    Dim targetLastCol As Long
    If IsMissing(lastCol) Then
        '最終列 取得
        targetLastCol = GetLastCol(targetWs, targetRow)
        If firstCol > targetLastCol Then Exit Function
    Else
        If Not IsNumber(lastCol) Then
            Call ShowErrMsg("lastColに指定された値は数値ではありません。" & "1以上" & MAX_COL & "以下で入力してください。", title:="GetRowRange")
            Exit Function
        Else
            If relativeFLG Then
                '最終列 取得
                targetLastCol = firstCol + lastCol
                If targetLastCol < 1 Or targetLastCol > MAX_COL Then
                    Call ShowErrMsg("指定された値が有効行数の範囲外です。" & "firstCol + lastColが" & "1以上" & MAX_COL & "以下となるよう入力してください。", title:="GetRowRange")
                    Exit Function
                End If
            Else
                '最終列 取得
                targetLastCol = lastCol
                If targetLastCol < 1 Or targetLastCol > MAX_COL Then
                    Call ShowErrMsg("lastColに指定された値が有効行数の範囲外です。" & "1以上" & MAX_COL & "以下で入力してください。", title:="GetRowRange")
                    Exit Function
                End If
            End If
        End If
    End If

    With targetWs
        Set GetRowRange = .Range(.Cells(targetRow, firstCol), .Cells(targetRow, targetLastCol))
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetRowRange")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetSearchFileCount(ByVal searchFilePath As String, ByVal searchNames As String, Optional ByVal enFileNameType As EnumFileNameType, Optional ByVal ignoreCase As Boolean) As Long
On Error GoTo Err

    With FSO
        '引数 解析
        searchNames = Replace(searchNames, ";", ",")
        Dim searchNameArr() As String
        searchNameArr = Split(searchNames, ",")

        Dim targetFiles As Files, targetFile As File
        Set targetFiles = .GetFolder(searchFilePath).Files

        If targetFiles.Count = 0 Then Exit Function

        '検索ファイル数 カウント
        Dim searchName As Variant
        Dim fileCount As Long
        Select Case enFileNameType
            'ファイル名
            Case enFileName
                For Each targetFile In targetFiles
                    Dim fileName As String
                    fileName = .GetFileName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If fileName Like searchName Then
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
            'ベース名
            Case enBaseName
                For Each targetFile In targetFiles
                    Dim baseName As String
                    baseName = .GetBaseName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If baseName Like searchName Then
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
            '拡張子名
            Case enExtensionName
                For Each targetFile In targetFiles
                    Dim extensionName As String
                    extensionName = .GetExtensionName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If extensionName Like searchName Or (ignoreCase And (UCase(extensionName) Like UCase(searchName))) Then
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
        End Select
    End With

    GetSearchFileCount = fileCount

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSearchFileCount")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetSearchFileSpec(ByVal searchFilePath As String, ByVal searchNames As String, Optional ByVal enFileNameType As EnumFileNameType, Optional ByVal ignoreCase As Boolean) As String()
On Error GoTo Err

    With FSO
        '引数 解析
        searchNames = Replace(searchNames, ";", ",")
        Dim searchNameArr() As String
        searchNameArr = Split(searchNames, ",")

        Dim targetFiles As Files, targetFile As File
        Set targetFiles = .GetFolder(searchFilePath).Files

        If targetFiles.Count = 0 Then Exit Function

        '検索ファイルディレクトリー 取得
        Dim fileSpecArr() As String
        Dim searchName As Variant
        Dim fileCount As Long
        Select Case enFileNameType
            'ファイル名
            Case enFileName
                For Each targetFile In targetFiles
                    Dim fileName As String
                    fileName = .GetFileName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If fileName Like searchName Then
                            ReDim Preserve fileSpecArr(fileCount)
                            fileSpecArr(fileCount) = targetFile.Path
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
            'ベース名
            Case enBaseName
                For Each targetFile In targetFiles
                    Dim baseName As String
                    baseName = .GetBaseName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If baseName Like searchName Then
                            ReDim Preserve fileSpecArr(fileCount)
                            fileSpecArr(fileCount) = targetFile.Path
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
            '拡張子名
            Case enExtensionName
                For Each targetFile In targetFiles
                    Dim extensionName As String
                    extensionName = .GetExtensionName(targetFile.Name)

                    For Each searchName In searchNameArr
                        If extensionName Like searchName Or (ignoreCase And (UCase(extensionName) Like UCase(searchName))) Then
                            ReDim Preserve fileSpecArr(fileCount)
                            fileSpecArr(fileCount) = targetFile.Path
                            fileCount = fileCount + 1
                        End If
                    Next searchName
                Next targetFile
        End Select
    End With

    GetSearchFileSpec = fileSpecArr

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSearchFileSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetSelectFileSpec(ByVal filePath As String, Optional ByVal dialogTitle As String, _
                                  Optional ByVal filterDescriptions As String, Optional ByVal filterExtensions As String, Optional ByVal filterSeparateFLG As Boolean, _
                                  Optional ByVal multiSelect As Boolean) As String()
On Error GoTo Err

    'ファイルパス 存在確認
    If Not FSO.FolderExists(filePath) Then
        Call ShowErrMsg("指定されたファイルパスが存在しません。", title:="GetSelectFileSpec")
        Exit Function
    End If

    '引数 空白確認
    If dialogTitle = "" Then dialogTitle = "ファイル 選択"
    If filterDescriptions = "" Then filterDescriptions = "すべてのファイル"
    If filterExtensions = "" Then filterExtensions = "*.*"

    Dim filterDescriptionArr() As String, filterExtensionArr() As String
    filterDescriptionArr = Split(filterDescriptions, ",")
    filterExtensionArr = Split(filterExtensions, ",")

    '配列 要素数比較
    Dim filterDescriptionArrUpper As Long, filterExtensionArrUpper As Long
    filterDescriptionArrUpper = UBound(filterDescriptionArr, 1)
    filterExtensionArrUpper = UBound(filterExtensionArr, 1)
    If filterDescriptionArrUpper <> filterExtensionArrUpper Then
        Call ShowErrMsg("指定されたファイルの種類と拡張子の数が一致しません。", title:="GetSelectFileSpec")
        Exit Function
    End If

    Dim filterArrUpper As Long
    filterArrUpper = filterDescriptionArrUpper

    '引数 解析
    Dim filterArrIndx As Long
    For filterArrIndx = 0 To filterArrUpper
        '説明
        If filterDescriptionArr(filterArrIndx) = "" Then filterDescriptionArr(filterArrIndx) = "すべてのファイル"

        '拡張子
        Dim bufArr As Variant, bufArrIndx As Long
        bufArr = Split(IIf(filterExtensionArr(filterArrIndx) = "", "*", filterExtensionArr(filterArrIndx)), ";")
        For bufArrIndx = 0 To UBound(bufArr, 1)
            If bufArr(bufArrIndx) = "" Then bufArr(bufArrIndx) = "*"
        Next bufArrIndx
        filterExtensionArr(filterArrIndx) = "*." & Join(bufArr, ";*.")
    Next filterArrIndx

    'フィルター分割 確認
    If Not filterSeparateFLG Then
        filterDescriptionArr(0) = Join(filterDescriptionArr, ",")
        ReDim Preserve filterDescriptionArr(0)
        filterExtensionArr(0) = Join(filterExtensionArr, ";")
        ReDim Preserve filterExtensionArr(0)

        filterArrUpper = 0
    End If

    'ダイアログボックス 設定
    With Application.FileDialog(msoFileDialogFilePicker)
        With .Filters
            Call .Clear
            For filterArrIndx = 0 To filterArrUpper
                Call .Add(filterDescriptionArr(filterArrIndx), filterExtensionArr(filterArrIndx))
            Next filterArrIndx
        End With
        .FilterIndex = 1
        .InitialFileName = FSO.BuildPath(filePath, Application.PathSeparator)
        .Title = dialogTitle
        .AllowMultiSelect = multiSelect

        'ファイル選択状態 判定
        If .Show = -1 Then
            Dim selectedArr() As String
            If multiSelect Then
                ReDim selectedArr(.SelectedItems.Count - 1)

                Dim indx As Long
                For indx = 1 To .SelectedItems.Count
                    selectedArr(indx - 1) = .SelectedItems(indx)
                Next indx

                GetSelectFileSpec = selectedArr
            Else
                ReDim selectedArr(0)

                selectedArr(0) = .SelectedItems(1)

                GetSelectFileSpec = selectedArr
            End If
        End If
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSelectFileSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function GetSelectFolderSpec(ByVal folderPath As String, Optional ByVal dialogTitle As String) As String
On Error GoTo Err

    'フォルダーパス 存在確認
    If Not FSO.FolderExists(folderPath) Then
        Call ShowErrMsg("指定されたフォルダーパスが存在しません。", title:="GetSelectFolderSpec")
        Exit Function
    End If

    '引数 空白確認
    If dialogTitle = "" Then dialogTitle = "フォルダー 選択"

    'ダイアログボックス 設定
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = dialogTitle
        .InitialFileName = FSO.BuildPath(folderPath, Application.PathSeparator)

        'フォルダー選択状態 判定
        If .Show = -1 Then GetSelectFolderSpec = .SelectedItems(1)
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSelectFolderSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 10:22:09
'----------------------------------------------------------------------------------------------------
Public Function GetShape(ByVal shapeName As String, Optional ByRef parent As Shapes) As Shape
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then
        'ワークシート 判定
        Dim actSh As Object
        Set actSh = ThisWorkbook.ActiveSheet
        If TypeName(actSh) = "Worksheet" Then
            If actSh.Type = xlWorksheet Then
                Set parent = actSh.Shapes
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    '全シェイプ 確認
    Dim shp As Shape
    For Each shp In parent
        If shp.Name = shapeName Then
            Set GetShape = shp
            Exit Function
        End If
    Next shp

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetShape")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function GetSheet(ByVal shName As String, Optional ByRef parentWb As Workbook, Optional ByVal shType As XlSheetType, Optional ByVal codeNameFLG As Boolean) As Object
On Error GoTo Err

    '引数 既定値判定
    If parentWb Is Nothing Then Set parentWb = ThisWorkbook

    Dim parent As Sheets
    Dim sh As Object
    Dim targetName As String
    'シートタイプ 判定
    Select Case shType
        Case xlChart, xlWorksheet
            Set parent = IIf(shType = xlChart, parentWb.Charts, parentWb.Worksheets)
            '全シート 確認
            For Each sh In parent
                targetName = IIf(codeNameFLG, sh.CodeName, sh.Name)
                If targetName = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh
        Case xlDialogSheet, xlExcel4IntlMacroSheet, xlExcel4MacroSheet
            Set parent = IIf(shType = xlDialogSheet, parentWb.DialogSheets, IIf(shType = xlExcel4IntlMacroSheet, parentWb.Excel4IntlMacroSheets, parentWb.Excel4MacroSheets))
            '全シート 確認
            For Each sh In parent
                If sh.Name = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh
        Case Else
            '全シート 確認
            For Each sh In parentWb.Sheets
                If sh.Name = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh
    End Select

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 06:06:07
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Function GetVBComponent(ByVal VBCName As String, ByVal VBCType As vbext_ComponentType, Optional ByRef parent As VBComponents) As VBComponent
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then Set parent = ThisWorkbook.VBProject.VBComponents

    '全VBComponent 確認
    Dim VBC As VBComponent
    For Each VBC In parent
        If VBC.Name = VBCName And VBC.Type = VBCType Then
            Set GetVBComponent = VBC
            Exit Function
        End If
    Next VBC

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetVBComponent")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'ProgID:Scripting.FileSystemObject
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
Public Function HasAttr(ByVal spec As String, ByVal attr As FileAttribute) As Boolean
On Error GoTo Err

    With FSO
        '属性 比較(ビット演算)
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then
            'フォルダー属性 比較(ビット演算)
            HasAttr = (.GetFolder(spec).Attributes And attr) = attr
        Else
            'ファイル属性 比較(ビット演算)
            HasAttr = (.GetFile(spec).Attributes And attr) = attr
        End If
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasAttr")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasListObject(ByVal loName As String, Optional ByRef parent As ListObjects) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then
        'ワークシート 判定
        Dim actSh As Object
        Set actSh = ThisWorkbook.ActiveSheet
        If TypeName(actSh) = "Worksheet" Then
            If actSh.Type = xlWorksheet Then
                Set parent = actSh.ListObjects
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    '全リストオブジェクト 確認
    Dim lo As ListObject
    For Each lo In parent
        If lo.Name = loName Then
            HasListObject = True
            Exit Function
        End If
    Next lo

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasListObject")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasName(ByVal nameName As String, Optional ByRef parent As Names) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then Set parent = ThisWorkbook.Names

    Dim defineName As Name
    'Worksheet 判定
    If TypeName(parent.Parent) = "Worksheet" Then
        '全名前 確認
        For Each defineName In parent
            If defineName.Name = parent.Parent.Name & "!" & nameName Then
                HasName = True
                Exit Function
            End If
        Next defineName
    Else
        '全名前 確認
        For Each defineName In parent
            If defineName.Name = nameName Then
                HasName = True
                Exit Function
            End If
        Next defineName
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 10:28:43
'----------------------------------------------------------------------------------------------------
Public Function HasShape(ByVal shapeName As String, Optional ByRef parent As Shapes) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then
        'ワークシート 判定
        Dim actSh As Object
        Set actSh = ThisWorkbook.ActiveSheet
        If TypeName(actSh) = "Worksheet" Then
            If actSh.Type = xlWorksheet Then
                Set parent = actSh.Shapes
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    '全シェイプ 確認
    Dim shp As Shape
    For Each shp In parent
        If shp.Name = shapeName Then
            HasShape = True
            Exit Function
        End If
    Next shp

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasShape")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/18 01:16:05
'----------------------------------------------------------------------------------------------------
Public Function HasSheet(ByVal shName As String, Optional ByRef parentWb As Workbook, Optional ByVal shType As XlSheetType, Optional ByVal codeNameFLG As Boolean) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parentWb Is Nothing Then Set parentWb = ThisWorkbook

    Dim parent As Sheets
    Dim sh As Object
    Dim targetName As String
    'シートタイプ 判定
    Select Case shType
        Case xlChart, xlWorksheet
            Set parent = IIf(shType = xlChart, parentWb.Charts, parentWb.Worksheets)
            '全シート 確認
            For Each sh In parent
                targetName = IIf(codeNameFLG, sh.CodeName, sh.Name)
                If targetName = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh
        Case xlDialogSheet, xlExcel4IntlMacroSheet, xlExcel4MacroSheet
            Set parent = IIf(shType = xlDialogSheet, parentWb.DialogSheets, IIf(shType = xlExcel4IntlMacroSheet, parentWb.Excel4IntlMacroSheets, parentWb.Excel4MacroSheets))
            '全シート 確認
            For Each sh In parent
                If sh.Name = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh
        Case Else
            '全シート 確認
            For Each sh In parentWb.Sheets
                If sh.Name = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh
    End Select

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Function HasVBComponent(ByVal VBCName As String, ByVal VBCType As vbext_ComponentType, Optional ByRef parent As VBComponents) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then Set parent = ThisWorkbook.VBProject.VBComponents

    '全VBComponent 確認
    Dim VBC As VBComponent
    For Each VBC In parent
        If VBC.Name = VBCName And VBC.Type = VBCType Then
            HasVBComponent = True
            Exit Function
        End If
    Next VBC

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasVBComponent")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasWorkbook(ByVal wbName As String, Optional ByRef parent As Workbooks) As Boolean
On Error GoTo Err

    '引数 既定値判定
    If parent Is Nothing Then Set parent = Workbooks

    '全ワークブック 確認
    Dim wb As Workbook
    For Each wb In parent
        If wb.Name = wbName Then
            HasWorkbook = True
            Exit Function
        End If
    Next wb

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasWorkbook")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:VBScript_RegExp_55
'LIBID:{3F4DACA7-160D-11D2-A8E9-00104B365C9F}
    'ReferenceName:Microsoft VBScript Regular Expressions 5.5
    'FullPath(win32):C:\Windows\SysWOW64\vbscript.dll\3
    'FullPath(win64):C:\Windows\System32\vbscript.dll\3
    'Major.Minor:5.5
        'ProgID:VBScript.RegExp
        'CLSID:{3F4DACA4-160D-11D2-A8E9-00104B365C9F}
Public Function IsAlphabet(ByVal expression As String) As Boolean
On Error GoTo Err

    '空白 判定
    If expression = "" Then Exit Function

    '正規表現 判定
    With REG
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = "^[A-Za-zＡ-Ｚａ-ｚ]+$"
        IsAlphabet = .Test(expression)
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "IsAlphabet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 03:29:12
'----------------------------------------------------------------------------------------------------
Public Function IsBlank(ByRef expression As Variant) As Boolean
On Error GoTo Err

    'Null判定
    If IsNull(expression) Then
        IsBlank = True
        Exit Function
    'オブジェクト判定
    ElseIf IsObject(expression) Then
        '空判定
        If Not expression Is Nothing Then Exit Function
    '配列判定
    ElseIf IsArray(expression) Then
On Error GoTo ArrayErr
        '空判定
        If LBound(expression, 1) <= UBound(expression, 1) Then Exit Function
ArrayErr:
On Error GoTo Err
    Else
        '空判定
        If CStr(expression) <> "" Then Exit Function
    End If

    IsBlank = True

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "IsBlank")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:VBScript_RegExp_55
'LIBID:{3F4DACA7-160D-11D2-A8E9-00104B365C9F}
    'ReferenceName:Microsoft VBScript Regular Expressions 5.5
    'FullPath(win32):C:\Windows\SysWOW64\vbscript.dll\3
    'FullPath(win64):C:\Windows\System32\vbscript.dll\3
    'Major.Minor:5.5
        'ProgID:VBScript.RegExp
        'CLSID:{3F4DACA4-160D-11D2-A8E9-00104B365C9F}
Public Function IsNumber(ByVal expression As String) As Boolean
On Error GoTo Err

    '空白 判定
    If expression = "" Then Exit Function

    '正規表現 判定
    With REG
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = "^-?\d*\.?\d*$"
        IsNumber = .Test(expression)
    End With

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "IsNumber")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function IsOpen(ByVal fileSpec As String) As Long
On Error Resume Next

    IsOpen = 1

    'ファイル存在確認
    If Not FSO.FileExists(fileSpec) Then Exit Function

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
    If Err.Number =  75 Then
        '開閉確認(Input)
        Open fileSpec For Input As #fileNumber
        Close #fileNumber
    End If

    IsOpen = Err.Number = 55 Or Err.Number = 70

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 03:29:12
'----------------------------------------------------------------------------------------------------
Public Function IsString(ByRef expression As Variant) As Boolean
On Error Resume Next

    Dim str As String
    str = CStr(expression)

    IsString = Not Err.Number > 0

End Function
'----------------------------------------------------------------------------------------------------