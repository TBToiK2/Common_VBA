Attribute VB_Name = "Mod_Button"
Option Explicit
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub CopyRun()

    Call Frm_Main.Show(vbModeless)

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub CreateCopySheet()

    '入力画面 表示
    Dim ShName As Variant
    ShName = Application.InputBox("作成するシート名を入力してください。", "シート 新規作成", Type:=2)

    If ShName = False Then Exit Sub

    'Templateシート コピー
    Call Template.Copy(, Worksheets(Worksheets.Count))

    'シート名 変更
    If Not IsBlank(ShName) Then ActiveSheet.Name = ShName

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub VisibleDescription(ByVal ShapeName As String)

    Application.ScreenUpdating = False

    With Main

        'シート プロテクト解除
        Call .Unprotect(PASS)

        '非表示範囲 判定, 取得
        Dim DescriptionRange As Range
        Select Case ShapeName
        Case "Tri_Visible_Summary"
            Set DescriptionRange = .Range(.Rows(5), .Rows(6))
        Case "Tri_Visible_Param"
            Set DescriptionRange = .Range(.Rows(8), .Rows(31))
        Case "Tri_Visible_Copy"
            Set DescriptionRange = .Range(.Rows(33), .Rows(39))
        Case "Tri_Visible_Use"
            Set DescriptionRange = .Range(.Rows(41), .Rows(87))
        Case Else
            GoTo Skip
        End Select

        '対象範囲 表示設定
        DescriptionRange.Hidden = Not (DescriptionRange.Hidden)

        '表示アイコン 角度設定
        Dim ShapeRotation As Long
        If DescriptionRange.Hidden Then ShapeRotation = 180
        .Shapes(ShapeName).Rotation = ShapeRotation

Skip:
        'シート プロテクト
        .Cells(1, 19).Activate
        .EnableSelection = xlUnlockedCells
        Call .Protect(PASS, UserInterfaceOnly:=True, AllowFormattingCells:=True) 'セル色設定ボタン完成時にAllowFormattingCellsを削除

    End With

    Application.ScreenUpdating = True

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Sub FontColor_Click()
    MsgBox "試作中"
End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Sub InteriorColor_Click()
    MsgBox "試作中"
End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub ResetFormat()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If ActiveSheet.CodeName = "Main" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws

        'シート プロテクト解除
        Call .Unprotect(PASS)

        '共通内範囲 設定
        Dim ComRange As Range
        Dim ComValidationRange As Range
        Set ComRange = .Range(.Cells(3, 3), .Cells(8, 4))
        Set ComValidationRange = .Range(.Cells(3, 3), .Cells(7, 4))

        With ComRange

            'ロック解除
            .Locked = False

            '書式 設定
            With .Font
                .Name = "メイリオ"
                .FontStyle = ""
                .Size = 10
            End With
            .HorizontalAlignment = xlHAlignCenter
            With .Borders
                .Color = vbBlack
                .Weight = xlThin
                .LineStyle = xlContinuous
            End With

            '結合 設定
            Call .Merge(True)

            '入力規則, 条件付き書式 設定
            Call .Validation.Delete
            Call .FormatConditions.Delete

            With ComValidationRange

                Call .Validation.Add(xlValidateList, xlValidAlertStop, xlEqual, "する, しない")

                Dim ComAddCondition As FormatCondition
                Set ComAddCondition = .FormatConditions.Add(xlBlanksCondition)
                ComAddCondition.Interior.Color = vbRed
            End With
        End With

        'ファイル別内範囲 設定
        Dim FileRange As Range
        Dim FileConditionRange As Range
        Dim FileNameRange As Range, FileRowColRange As Range
        Set FileRange = .Range(.Cells(13, 3), .Cells(19, 4))
        Set FileConditionRange = .Range(.Cells(13, 3), .Cells(18, 4))
        Set FileNameRange = .Range(.Cells(13, 3), .Cells(13, 4))
        Set FileRowColRange = .Range(.Cells(15, 3), .Cells(18, 4))

        With FileRange

            'ロック解除
            .Locked = False

            '書式 設定
            With .Font
                .Name = "メイリオ"
                .FontStyle = ""
                .Size = 10
            End With
            .HorizontalAlignment = xlHAlignGeneral
            With .Borders
                .Color = vbBlack
                .Weight = xlThin
                .LineStyle = xlContinuous
                .Item(xlEdgeBottom).Weight = xlMedium
                .Item(xlInsideVertical).LineStyle = xlDash
            End With

            '結合 設定
            Call .UnMerge

            '条件付き書式 設定
            Call .Validation.Delete
            Call .FormatConditions.Delete

            With FileConditionRange
                Dim FileAddCondition As FormatCondition
                Set FileAddCondition = .FormatConditions.Add(xlBlanksCondition)
                FileAddCondition.Interior.Color = vbRed
            End With
        End With
        FileNameRange.WrapText = True
        FileRowColRange.HorizontalAlignment = xlHAlignCenter

        'コピー項目表内範囲 設定
        Dim CopyItemDataEnd As Range
        Dim CopyRange As Range, CopyAllRange As Range
        Dim CopyHeaderRange As Range
        Dim CopyInChargeRange As Range, CopyDiscriptionRange As Range, CopyDuplicationRange As Range
        Set CopyItemDataEnd = .Cells(.Rows.Count, 2).End(xlUp)
        Set CopyRange = .Range(.Cells(21, 1), .Cells(CopyItemDataEnd.Row, 5))
        Set CopyAllRange = .Range(.Cells(21, 1), .Cells(.Rows.Count, 5))
        Set CopyHeaderRange = .Range(.Cells(21, 1), .Cells(21, 5))
        Set CopyInChargeRange = .Range(.Cells(22, 4), .Cells(CopyItemDataEnd.Row, 4))
        Set CopyDiscriptionRange = .Range(.Cells(22, 5), .Cells(CopyItemDataEnd.Row, 5))
        Set CopyDuplicationRange = .Range(.Cells(22, 6), .Cells(.Rows.Count, 6))

        '書式 クリア
        CopyAllRange.Borders.LineStyle = xlLineStyleNone

        With CopyRange

            'ロック解除
            .Locked = False

            '書式 設定
            With .Font
                .Name = "メイリオ"
                .FontStyle = ""
                .Size = 10
            End With
            .HorizontalAlignment = xlHAlignGeneral
            With .Borders
                .Color = vbBlack
                .Weight = xlThin
                .LineStyle = xlContinuous
                .Item(xlEdgeRight).Weight = xlMedium
                .Item(xlEdgeBottom).Weight = xlMedium
                .Item(xlEdgeLeft).Weight = xlMedium
            End With
            Call .Rows.AutoFit
        End With
        CopyDiscriptionRange.WrapText = True
        With CopyHeaderRange
            .HorizontalAlignment = xlHAlignCenter
            .Font.Bold = True
        End With
        CopyInChargeRange.HorizontalAlignment = xlHAlignCenter

        '重複項目 再設定
        '値 消去
        With CopyDuplicationRange
            Call .Clear
            .Locked = False
        End With

        '関数 設定
        If CopyItemDataEnd.Row > CopyHeaderRange.Row Then

            Dim Address As String, CountIf As String
            Address = CopyItemDataEnd.Address
            CountIf = "COUNTIF($B$22:" & Address & ", $B22)"

            With .Cells(22, 6)

                .Formula = "=IF(" & CountIf & "> 1, ""["" & $B22 & ""] × "" &" & CountIf & ", """")"
                If CopyItemDataEnd.Row - CopyHeaderRange.Row > 1 Then
                    Call .AutoFill(.Resize(CopyItemDataEnd.Row - CopyHeaderRange.Row))
                End If
            End With
        End If

        'シート プロテクト
        Call .Protect(PASS, UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True)

    End With

    Application.DisplayAlerts = False
    Application.ScreenUpdating = True

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Sub ResetChangedFormat(Target As Range)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If ActiveSheet.CodeName = "Main" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws

        'シート プロテクト解除
        Call .Unprotect(PASS)

        Dim ComRange As Range, FileRange As Range
        Set ComRange = .Range(.Cells(3, 3), .Cells(8, 4))
        Set FileRange = .Range(.Cells(13, 3), .Cells(19, 4))

        '共通内範囲 存在確認
        Dim IntersectComRange As Range
        Set IntersectComRange = Intersect(ComRange, Target)
        If Not IntersectComRange Is Nothing Then

            Dim ComValidationRange As Range
            Dim ComFormatCopyRange As Range, ComMarkRange As Range
            Set ComValidationRange = .Range(.Cells(3, 3), .Cells(7, 4))
            Set ComFormatCopyRange = .Range(.Cells(4, 3), .Cells(4, 3))
            Set ComMarkRange = .Range(.Cells(8, 3), .Cells(8, 4))

            With IntersectComRange

                '書式 設定
                .Font.Size = 10
                .HorizontalAlignment = xlHAlignCenter
                With .Borders
                    .Color = vbBlack
                    .Weight = xlThin
                    .LineStyle = xlContinuous
                End With

                '結合 設定
                If Not .MergeCells Or IsNull(.MergeCells) Then
                    If .Columns.Count = 1 Then
                        Call .Resize(, 2).Merge(True)
                    Else
                        Call .Merge(True)
                    End If
                End If

                '入力規則内範囲 存在確認
                Dim IntersectComValidationRange As Range
                Set IntersectComValidationRange = Intersect(IntersectComRange, ComValidationRange)
                If Not IntersectComValidationRange Is Nothing Then

                    '入力規則 設定
                    With IntersectComValidationRange.Validation
                        Call .Delete
                        Call .Add(xlValidateList, xlValidAlertStop, xlEqual, "する, しない")
                    End With
                    
                    '条件付き書式 設定
                    With ComValidationRange.FormatConditions

                        Call .Delete

                        Dim ComAddCondition As FormatCondition
                        Set ComAddCondition = .Add(xlBlanksCondition)
                        ComAddCondition.Interior.Color = vbRed
                    End With
                End If

                '書式コピー内範囲 存在確認
                If Not Intersect(IntersectComRange, ComFormatCopyRange) Is Nothing Then

                    'コピーセル文字, 背景色項目 活性設定
                    Dim HiddenFLG As Boolean
                    If ConvertTrueFalse(ComFormatCopyRange.Value, HiddenFLG) Then
                        ws.Shapes("Rct_Hidden").Visible = HiddenFLG
                    Else
                        ws.Shapes("Rct_Hidden").Visible = False
                    End If
                End If

                'コピー変更行マーク内範囲 存在確認
                Dim IntersectComMarkRange As Range
                Set IntersectComMarkRange = Intersect(IntersectComRange, ComMarkRange)
                If Not IntersectComMarkRange Is Nothing Then
                    With IntersectComMarkRange

                        '入力規則 設定
                        Call .Validation.Delete

                        '条件付き書式 設定
                        Call .FormatConditions.Delete
                    End With
                End If
            End With
        End If

        'ファイル別内範囲 存在確認
        Dim IntersectFileRange As Range
        Set IntersectFileRange = Intersect(FileRange, Target)
        If Not IntersectFileRange Is Nothing Then

            Dim FileConditionRange As Range
            Dim FileNameRange As Range, FileRowColRange As Range, FilePassRange As Range
            Dim FileFRange As Range, FileTRange As Range
            Set FileConditionRange = .Range(.Cells(13, 3), .Cells(18, 4))
            Set FileNameRange = .Range(.Cells(13, 3), .Cells(13, 4))
            Set FileRowColRange = .Range(.Cells(15, 3), .Cells(18, 4))
            Set FilePassRange = .Range(.Cells(19, 3), .Cells(19, 4))
            Set FileFRange = .Range(.Cells(13, 3), .Cells(19, 3))
            Set FileTRange = .Range(.Cells(13, 4), .Cells(19, 4))

            With IntersectFileRange

                '書式 設定
                .Font.Size = 10
                .HorizontalAlignment = xlHAlignGeneral
                With .Borders
                    .Color = vbBlack
                    .Weight = xlThin
                    .LineStyle = xlContinuous
                End With

                '結合 設定
                Call .UnMerge

                '入力規則 設定
                .Validation.Delete

                '条件付き書式内範囲 存在確認
                Dim IntersectFileConditionRange As Range
                Set IntersectFileConditionRange = Intersect(IntersectFileRange, FileConditionRange)
                If Not IntersectFileConditionRange Is Nothing Then

                    ' 条件付き書式 設定
                    With FileConditionRange.FormatConditions

                        Call .Delete

                        Dim FileAddCondition As FormatCondition
                        Set FileAddCondition = .Add(xlBlanksCondition)
                        FileAddCondition.Interior.Color = vbRed
                    End With
                End If

                'ファイル名内範囲 存在確認
                Dim IntersectFileNameRange As Range
                Set IntersectFileNameRange = Intersect(IntersectFileRange, FileNameRange)
                If Not IntersectFileNameRange Is Nothing Then

                    '書式 設定
                    IntersectFileNameRange.WrapText = True
                End If

                '行列内範囲 存在確認
                Dim IntersectFileRowColRange As Range
                Set IntersectFileRowColRange = Intersect(IntersectFileRange, FileRowColRange)
                If Not IntersectFileRowColRange Is Nothing Then

                    '書式 設定
                    IntersectFileRowColRange.HorizontalAlignment = xlHAlignCenter
                End If

                'パスワード内範囲 存在確認
                Dim IntersectFilePassRange As Range
                Set IntersectFilePassRange = Intersect(IntersectFileRange, FilePassRange)
                If Not IntersectFilePassRange Is Nothing Then
                    With IntersectFilePassRange

                        '条件付き書式 設定
                        Call .FormatConditions.Delete

                        '書式 設定
                        .Borders(xlEdgeBottom).Weight = xlMedium
                    End With
                End If

                'コピー元パラメータ内範囲 存在確認
                Dim IntersectFileFRange As Range
                Set IntersectFileFRange = Intersect(IntersectFileRange, FileFRange)
                If Not IntersectFileFRange Is Nothing Then

                    '書式 設定
                    IntersectFileFRange.Borders(xlEdgeRight).LineStyle = xlDash
                End If

                'コピー先パラメータ内範囲 存在確認
                Dim IntersectFileTRange As Range
                Set IntersectFileTRange = Intersect(IntersectFileRange, FileTRange)
                If Not IntersectFileTRange Is Nothing Then

                    '書式 設定
                    IntersectFileTRange.Borders(xlEdgeLeft).LineStyle = xlDash
                End If
            End With
        End If

        'シート プロテクト
        Call .Protect(PASS, UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True)

    End With

    Application.DisplayAlerts = False
    Application.ScreenUpdating = True

End Sub
'----------------------------------------------------------------------------------------------------
