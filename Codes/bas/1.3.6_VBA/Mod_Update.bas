Attribute VB_Name = "Mod_Update"
'1.3.6_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2021/07/26 10:48:35
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
Public Sub AddAttr(ByVal spec As String, ByVal attr As FileAttribute)
On Error GoTo Err

    With FSO

        '属性 比較(ビット演算)
        Dim afterAttr As FileAttribute
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then

            Dim attrFolder As Folder
            Set attrFolder = .GetFolder(spec)
            'フォルダー属性 比較(ビット演算)
            afterAttr = attrFolder.Attributes Or attr
            attrFolder.Attributes = afterAttr

        Else

            Dim attrFile As File
            Set attrFile = .GetFile(spec)
            'ファイル属性 比較(ビット演算)
            afterAttr = .GetFile(spec).Attributes Or attr
            attrFile.Attributes = afterAttr

        End If

    End With

    Exit Sub

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "AddAttr")

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/05 10:56:29
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Sub CopyVBComponent(ByRef source As VBProject, ByRef destination As VBProject, ByVal VBCType As vbext_ComponentType, _
                           ByVal sourceVBCName As String, Optional ByVal destinationVBCName As String, Optional ByVal overwrite As Boolean)
On Error GoTo Err

    If source Is destination Then
        Call ShowErrMsg("コピー元とコピー先のVBProjectオブジェクトが同じです。", title:="CopyVBComponent")
        Exit Sub
    End If

    'コピー先VBComponents名 空白確認
    If destinationVBCName = "" Then destinationVBCName = sourceVBCName

    'コピー元VBComponents 存在確認
    If Not HasVBComponent(sourceVBCName, VBCType, source.VBComponents) Then
        Call ShowErrMsg("コピー元に指定したVBComponent名が存在しません。", title:="CopyVBComponent")
        Exit Sub
    End If

    'コピー先VBComponents 存在確認
    Dim destinationVBC As VBComponent
    If HasVBComponent(destinationVBCName, VBCType, destination.VBComponents) Then
        If Not overwrite Then
            Call ShowErrMsg("指定したVBComponent名は既にコピー先に存在しています。", title:="CopyVBComponent")
            Exit Sub
        End If

        Set destinationVBC = destination.VBComponents(destinationVBCName)
    Else
        If VBCType = vbext_ct_ActiveXDesigner Or VBCType = vbext_ct_Document Then
            Call ShowErrMsg("指定した型のVBComponentオブジェクトは上書きでのみコピー可能です。", title:="CopyVBComponent")
            Exit Sub
        End If

        Set destinationVBC = destination.VBComponents.Add(VBCType)
        destinationVBC.Name = destinationVBCName
    End If

    'VBComponent CodeModuleコピー
    Dim moduleCode As String
    With source.VBComponents(sourceVBCName).CodeModule
        moduleCode = .Lines(1, .CountOfLines)
    End With
    With destinationVBC.CodeModule
        Call .DeleteLines(1, .CountOfLines)
        Call .AddFromString(moduleCode)
    End With

    Exit Sub

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "CopyVBComponent")

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 03:29:12
'----------------------------------------------------------------------------------------------------
Public Function ConcatColumnsValues(ByRef targetWs As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal separateFLG As Boolean, ParamArray targetCols() As Variant) As String()
On Error GoTo Err

    If IsMissing(targetCols) Then Exit Function

    Dim rowUpper As Long, colUpper As Long
    rowUpper = lastRow - firstRow
    If separateFLG Then colUpper = UBound(targetCols, 1)

    Dim targetValueArr() As String
    ReDim targetValueArr(0 To rowUpper, 0 To colUpper)

    Dim targetRow As Long
    For targetRow = firstRow To lastRow
        Dim targetValueArrIndx As Long, targetValue As String
        targetValueArrIndx = 0
        targetValue = ""

        Dim paramArrIndx As Long
        For paramArrIndx = 0 To UBound(targetCols, 1)
            If IsArray(targetCols(paramArrIndx)) Then
                Dim targetCol As Variant
                For Each targetCol In targetCols(paramArrIndx)
                    targetValue = targetValue & targetWs.Cells(targetRow, targetCol).Value
                Next targetCol
            Else
                targetValue = targetValue & targetWs.Cells(targetRow, targetCols(paramArrIndx)).Value
            End If

            If separateFLG And targetValue <> "" Then
                targetValueArr(targetRow - firstRow, targetValueArrIndx) = targetValue
                targetValueArrIndx = targetValueArrIndx + 1
                targetValue = ""
            End If
        Next paramArrIndx

        If targetValue <> "" Then
            targetValueArr(targetRow - firstRow, 0) = targetValue
        End If
    Next targetRow

    ConcatColumnsValues = targetValueArr

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ConcatColumnsValues")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 03:29:12
'----------------------------------------------------------------------------------------------------
Public Function ConcatRowsValues(ByRef targetWs As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long, ByVal separateFLG As Boolean, ParamArray targetRows() As Variant) As String()
On Error GoTo Err

    If IsMissing(targetRows) Then Exit Function

    Dim rowUpper As Long, colUpper As Long
    If separateFLG Then rowUpper = UBound(targetRows, 1)
    colUpper = lastCol - firstCol

    Dim targetValueArr() As String
    ReDim targetValueArr(0 To colUpper, 0 To rowUpper)

    Dim targetCol As Long
    For targetCol = firstCol To lastCol
        Dim targetValueArrIndx As Long, targetValue As String
        targetValueArrIndx = 0
        targetValue = ""

        Dim paramArrIndx As Long
        For paramArrIndx = 0 To UBound(targetRows, 1)
            If IsArray(targetRows(paramArrIndx)) Then
                Dim targetRow As Variant
                For Each targetRow In targetRows(paramArrIndx)
                    targetValue = targetValue & targetWs.Cells(targetRow, targetCol).Value
                Next targetRow
            Else
                targetValue = targetValue & targetWs.Cells(targetRows(paramArrIndx), targetCol).Value
            End If

            If separateFLG And targetValue <> "" Then
                targetValueArr(targetCol - firstCol, targetValueArrIndx) = targetValue
                targetValueArrIndx = targetValueArrIndx + 1
                targetValue = ""
            End If
        Next paramArrIndx

        If targetValue <> "" Then
            targetValueArr(targetCol - firstCol, 0) = targetValue
        End If
    Next targetCol

    ConcatRowsValues = targetValueArr

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ConcatRowsValues")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function ConvColAddrRef(ByVal colAddr As String, ByVal toReferenceStyle As XlReferenceStyle) As String
On Error Resume Next

    '参照形式 判定
    Dim colAddrRef As String
    If toReferenceStyle = xlA1 Then
        colAddrRef = Split(Cells(1, CLng(colAddr)).Address(ReferenceStyle:=toReferenceStyle), "$")(1)
    Else
        colAddrRef = Range(colAddr & "1").Column
    End If

    If colAddrRef <> "" Then
        ConvColAddrRef = colAddrRef
    Else
        Call ShowErrMsg("指定された値が正しくありません。", title:="ConvColAddrRef")
    End If

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/21 02:08:32
'----------------------------------------------------------------------------------------------------
Public Function ConvTrueFalse(ByVal expression As String) As Long
On Error GoTo Err

    '真偽判定用配列 作成
    Dim trueArr() As Variant, falseArr() As Variant
    trueArr = Array("True", "T", "Yes", "Y", "はい", "する")
    falseArr = Array("False", "F", "No", "N", "いいえ", "しない")

    Dim t As Variant, f As Variant
    'True 判定
    For Each t In trueArr
        If StrComp(expression, t, vbTextCompare) = 0 Then
            ConvTrueFalse = True
            Exit Function
        End If
    Next t
    'False 判定
    For Each f In falseArr
        If StrComp(expression, f, vbTextCompare) = 0 Then
            ConvTrueFalse = False
            Exit Function
        End If
    Next f

    ConvTrueFalse = 1

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ConvTrueFalse")

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
        'ProgID:Scripting.Dictionary
        'CLSID:{EE09B103-97E0-11CF-978F-00A02463E06F}
Public Function CopySheet(ByRef copySh As Object, Optional ByRef before As Variant, Optional ByRef after As Variant) As Object
On Error GoTo Err

    Dim sheetsFLG As Boolean

    'シート判定
    Select Case TypeName(copySh)
        'beforeシート指定
        Case "Worksheet", "Chart", "DialogSheet"
        Case "Sheets"
            sheetsFLG = True
        Case "Nothing"
            Call ShowErrMsg("copyShオブジェクトが空です。", title:="CopySheet")
            Exit Function
        Case Else
            Call ShowErrMsg("copyShにはシートオブジェクトを指定してください。", title:="CopySheet")
            Exit Function
    End Select

    Dim wb As Workbook
    'before, after未指定
    If IsMissing(before) And IsMissing(after) Then
        'シート 新規ブックコピー
        Call copySh.Copy
        'SheetsFLG 判定
        If sheetsFLG Then
            Set CopySheet = ActiveWorkbook.Sheets
        Else
            Set CopySheet = ActiveWorkbook.ActiveSheet
        End If

        Exit Function

    'after未指定
    ElseIf IsMissing(after) Then
        Select Case TypeName(before)
            'beforeシート指定
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = before.Parent
            Case "Nothing"
                Call ShowErrMsg("beforeオブジェクトが空です。", title:="CopySheet")
                Exit Function
            Case Else
                Call ShowErrMsg("beforeにはシートオブジェクトを指定してください。", title:="CopySheet")
                Exit Function
        End Select

    'before未指定
    ElseIf IsMissing(before) Then
        Select Case TypeName(after)
            'afterシート指定
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = after.Parent
            Case "Nothing"
                Call ShowErrMsg("afterオブジェクトが空です。", title:="CopySheet")
                Exit Function
            Case Else
                Call ShowErrMsg("afterにはシートオブジェクトを指定してください。", title:="CopySheet")
                Exit Function
        End Select

    'before, after指定
    Else
        Call ShowErrMsg("beforeとafterは同時に指定できません。", title:="CopySheet")
        Exit Function
    End If

    'コピー先ワークブック 判定
    Dim shs As Sheets
    If wb.Name = ThisWorkbook.Name Then
        Set shs = ThisWorkbook.Sheets
    Else
        Set shs = wb.Sheets
    End If

    'コピー前シート名, インデックス ディクショナリー格納
    Dim shDIC As New Dictionary
    Dim befSh As Object
    For Each befSh In shs
        Call shDIC.Add(befSh.Name, befSh.Index)
    Next befSh

    'シート コピー
    If IsMissing(after) Then
        Call copySh.Copy(Before:=before)
    Else
        Call copySh.Copy(After:=after)
    End If

    'コピーシートインデックス 配列格納
    Dim shIndxArr() As Variant
    shIndxArr = Array()
    Dim aftSh As Object
    For Each aftSh In shs
        If Not shDIC.Exists(aftSh.Name) Then
            ReDim Preserve shIndxArr(UBound(shIndxArr) + 1)
            shIndxArr(UBound(shIndxArr)) = aftSh.Index
        End If
    Next aftSh

    '配列要素数 判定
    If UBound(shIndxArr) > -1 Then
        'SheetsFLG 判定
        If sheetsFLG Then
            Set CopySheet = wb.Sheets(shIndxArr)
        Else
            Set CopySheet = wb.Sheets(shIndxArr)(1)
        End If
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "CopySheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/07 17:00:00
'----------------------------------------------------------------------------------------------------
Public Function Max(ParamArray expressions() As Variant) As Variant
On Error GoTo Err

    If IsMissing(expressions) Then
        Call Err.Raise(450)
    End If

    Dim numericArgFLG As Boolean

    Dim paramArrIndx As Long
    For paramArrIndx = 0 To UBound(expressions, 1)
        If Not IsArray(expressions(paramArrIndx)) Then expressions(paramArrIndx) = Array(expressions(paramArrIndx))

        Dim expression As Variant
        For Each expression In expressions(paramArrIndx)
            '日付 シリアル値変換
            If IsDate(expression) Then expression = CDbl(expression)
            '数値比較
            If IsNumeric(expression) And Not IsEmpty(expression) Then
                Max = IIf(IsEmpty(Max), expression, IIf(Max > expression, Max, expression))
                numericArgFLG = True
            End If
        Next expression
    Next paramArrIndx

    If Not numericArgFLG Then
        Call Err.Raise(450)
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "Max")
    Max = Null

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/07 17:00:00
'----------------------------------------------------------------------------------------------------
Public Function Min(ParamArray expressions() As Variant) As Variant
On Error GoTo Err

    If IsMissing(expressions) Then
        Call Err.Raise(450)
    End If

    Dim numericArgFLG As Boolean

    Dim paramArrIndx As Long
    For paramArrIndx = 0 To UBound(expressions, 1)
        If Not IsArray(expressions(paramArrIndx)) Then expressions(paramArrIndx) = Array(expressions(paramArrIndx))

        Dim expression As Variant
        For Each expression In expressions(paramArrIndx)
            '日付 シリアル値変換
            If IsDate(expression) Then expression = CDbl(expression)
            '数値比較
            If IsNumeric(expression) And Not IsEmpty(expression) Then
                Min = IIf(IsEmpty(Min), expression, IIf(Min < expression, Min, expression))
                numericArgFLG = True
            End If
        Next expression
    Next paramArrIndx

    If Not numericArgFLG Then
        Call Err.Raise(450)
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "Min")
    Min = Null

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
        'ProgID:Scripting.Dictionary
        'CLSID:{EE09B103-97E0-11CF-978F-00A02463E06F}
Public Function MoveSheet(ByRef moveSh As Object, Optional ByRef before As Variant, Optional ByRef after As Variant) As Object
On Error GoTo Err

    Dim sheetsFLG As Boolean

    'シート判定
    Select Case TypeName(moveSh)
        'beforeシート指定
        Case "Worksheet", "Chart", "DialogSheet"
        Case "Sheets"
            sheetsFLG = True
        Case "Nothing"
            Call ShowErrMsg("moveShオブジェクトが空です。", title:="MoveSheet")
            Exit Function
        Case Else
            Call ShowErrMsg("moveShにはシートオブジェクトを指定してください。", title:="MoveSheet")
            Exit Function
    End Select

    Dim wb As Workbook
    'before, after未指定
    If IsMissing(before) And IsMissing(after) Then
        'シート 新規ブック移動
        Call moveSh.Move
        'SheetsFLG 判定
        If sheetsFLG Then
            Set MoveSheet = ActiveWorkbook.Sheets
        Else
            Set MoveSheet = ActiveWorkbook.ActiveSheet
        End If

        Exit Function

    'after未指定
    ElseIf IsMissing(after) Then
        Select Case TypeName(before)
            'beforeシート指定
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = before.Parent
            Case "Nothing"
                Call ShowErrMsg("beforeオブジェクトが空です。", title:="MoveSheet")
                Exit Function
            Case Else
                Call ShowErrMsg("beforeにはシートオブジェクトを指定してください。", title:="MoveSheet")
                Exit Function
        End Select

    'before未指定
    ElseIf IsMissing(before) Then
        Select Case TypeName(after)
            'afterシート指定
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = after.Parent
            Case "Nothing"
                Call ShowErrMsg("afterオブジェクトが空です。", title:="MoveSheet")
                Exit Function
            Case Else
                Call ShowErrMsg("afterにはシートオブジェクトを指定してください。", title:="MoveSheet")
                Exit Function
        End Select

    'before, after指定
    Else
        Call ShowErrMsg("beforeとafterは同時に指定できません。", title:="MoveSheet")
        Exit Function
    End If

    '移動先ワークブック 判定
    Dim shs As Sheets
    If wb.Name = ThisWorkbook.Name Then
        Set shs = ThisWorkbook.Sheets
    Else
        Set shs = wb.Sheets
    End If

    '移動前シート名, インデックス ディクショナリー格納
    Dim shDIC As New Dictionary
    Dim befSh As Object
    For Each befSh In shs
        Call shDIC.Add(befSh.Name, befSh.Index)
    Next befSh

    'シート 移動
    If IsMissing(after) Then
        Call moveSh.Move(Before:=before)
    Else
        Call moveSh.Move(After:=after)
    End If

    '移動先ワークブック 判定
    If wb.Name = ThisWorkbook.Name Then
        Set MoveSheet = moveSh
        Exit Function
    End If

    '移動シートインデックス 配列格納
    Dim shIndxArr() As Variant
    shIndxArr = Array()
    Dim aftSh As Object
    For Each aftSh In shs
        If Not shDIC.Exists(aftSh.Name) Then
            ReDim Preserve shIndxArr(UBound(shIndxArr) + 1)
            shIndxArr(UBound(shIndxArr)) = aftSh.Index
        End If
    Next aftSh

    '配列要素数 判定
    If UBound(shIndxArr) > -1 Then
        'SheetsFLG 判定
        If sheetsFLG Then
            Set MoveSheet = wb.Sheets(shIndxArr)
        Else
            Set MoveSheet = wb.Sheets(shIndxArr)(1)
        End If
    End If

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "MoveSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/27 03:14:45
'----------------------------------------------------------------------------------------------------
Public Function ReplaceCRLF(ByVal expression As String) As String
On Error Resume Next

    'Carriage Return, Line Feed 削除
    ReplaceCRLF = Replace(Replace(expression, vbCr, vbNullString), vbLf, vbNullString)

End Function
'----------------------------------------------------------------------------------------------------