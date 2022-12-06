Attribute VB_Name = "Mod_Create"
'1.1.4_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2021/11/06 13:48:40
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 02:51:33
'----------------------------------------------------------------------------------------------------
Public Function ArrayByType(ByVal arrType As VbVarType, ParamArray arr() As Variant) As Variant
On Error Resume Next

    '配列 初期判定
    If UBound(arr) = -1 Then
        ArrayByType = Array()
        Exit Function
    End If

    Dim arrLB As Long, arrUB As Long
    arrLB = LBound(arr)
    arrUB = UBound(arr)

    Dim cnt As Long, element As Variant
    Select Case arrType
        '0
        Case vbEmpty
            Dim emptyArr() As Variant
            ReDim emptyArr(arrLB To arrUB)
            ArrayByType = emptyArr
        '1
        Case vbNull
            Dim nullArr() As Variant
            ReDim nullArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                nullArr(cnt) = Null
            Next cnt
            ArrayByType = nullArr
        '2
        Case vbInteger
            Dim intArr() As Integer
            ReDim intArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                intArr(cnt) = CInt(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = intArr
        '3
        Case vbLong
            Dim lngArr() As Long
            ReDim lngArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                lngArr(cnt) = CLng(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = lngArr
        '4
        Case vbSingle
            Dim sngArr() As Single
            ReDim sngArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                sngArr(cnt) = CSng(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = sngArr
        '5
        Case vbDouble
            Dim dblArr() As Double
            ReDim dblArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                dblArr(cnt) = CDbl(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = dblArr
        '6
        Case vbCurrency
            Dim curArr() As Currency
            ReDim curArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                curArr(cnt) = CCur(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = curArr
        '7
        Case vbDate
            Dim dateArr() As Date
            ReDim dateArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                dateArr(cnt) = CDate(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = dateArr
        '8
        Case vbString
            Dim strArr() As String
            ReDim strArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                strArr(cnt) = CStr(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = strArr
        '9
        Case vbObject
            Dim objArr() As Object
            ReDim objArrArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                Set objArrArr(cnt) = arr(cnt)
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = objArrArr
        '11
        Case vbBoolean
            Dim boolArr() As Boolean
            ReDim boolArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                boolArr(cnt) = CBool(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = boolArr
        '12
        Case vbVariant
            ArrayByType = arr
        '14
        Case vbDecimal
            Dim decArr() As Variant
            ReDim decArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                decArr(cnt) = CDec(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = decArr
        '17
        Case vbByte
            Dim byteArr() As Byte
            ReDim byteArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                byteArr(cnt) = CByte(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = byteArr

        #If Win64 Then

        '20
        Case vbLongLong
            Dim lnglngArr() As LongLong
            ReDim lnglngArr(arrLB To arrUB)
            For cnt = arrLB To arrUB
                lnglngArr(cnt) = CLngLng(arr(cnt))
            Next cnt
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = lnglngArr

        #End If

        Case Else
            Call ShowErrMsg("指定されたデータ型は無効です。", title:="ConvTrueFalse")
            Exit Function
    End Select

    Exit Function

'エラー処理
Err_Convert:

    Call ShowErrMsg("指定されたデータ型ではない要素が含まれています。", title:="ConvTrueFalse")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
        'ProgID:Scripting.FileSystemObject
Public Function BuildPath(ParamArray path() As Variant) As String
On Error Resume Next

    With FSO

        Dim maxElement As Long
        Dim element As Long
        Dim pathParam As String
        '第一引数 配列判定
        If IsArray(path(0)) Then

            Dim pathArr() As Variant
            pathArr = path(0)

            'パラメータ数 判定
            maxElement = UBound(pathArr, 1)
            Select Case maxElement
                Case Is = -1
                    Call ShowErrMsg("指定された配列に値が一つも存在しません。", title:="BuildPath")

                Case Is = 0
                    pathParam = CStr(pathArr(0))
                    'エラー 判定
                    If Err.Number > 0 Then GoTo Err_Array
                    'パス 作成
                    BuildPath = .BuildPath(pathParam, vbNullString)

                Case Else
                    For element = 0 To maxElement
                        pathParam = CStr(pathArr(element))
                        'エラー 判定
                        If Err.Number > 0 Then GoTo Err_Array
                        'パス 作成
                        BuildPath = .BuildPath(BuildPath, pathParam)
                    Next element

            End Select

        Else

            'パラメータ数 判定
            maxElement = UBound(path, 1)
            Select Case maxElement
                Case Is = 0
                    'パス 作成
                    BuildPath = .BuildPath(path(0), vbNullString)

                Case Else
                    For element = 0 To maxElement
                        pathParam = CStr(path(element))
                        'エラー 判定
                        If Err.Number > 0 Then GoTo Err_Array
                        'パス 作成
                        BuildPath = .BuildPath(BuildPath, pathParam)
                    Next element

            End Select

        End If

    End With

    Exit Function

'エラー処理
Err_Array:

    Call ShowErrMsg("指定された配列内に文字列に変換できない要素が含まれています。", title:="BuildPath")

    BuildPath = vbNullString

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
'参照設定
'LibraryName:Scripting
'LIBID:{420B2830-E718-11CF-893D-00A0C9054228}
    'ReferenceName:Microsoft Scripting Runtime
    'FullPath(win32):C:\Windows\SysWOW64\scrrun.dll
    'FullPath(win64):C:\Windows\System32\scrrun.dll
    'Major.Minor:1.0
        'CLSID:{0D43FE01-F093-11CF-8940-00A0C9054228}
        'ProgID:Scripting.FileSystemObject
Public Function CreateBackupFile(ByVal fileSpec As String) As Boolean
On Error GoTo Err

    With FSO

        'ファイル存在 確認
        If Not .FileExists(fileSpec) Then
            Call ShowErrMsg("指定されたファイルディレクトリは存在しません。", title:="CreateBackupFile")
            Exit Function
        End If

        'ファイル名, 拡張子 取得
        Dim baseName As String, extensionName As String
        baseName = .GetBaseName(fileSpec)
        extensionName = .GetExtensionName(fileSpec)

    End With

    Call FileCopy(fileSpec, ThisWorkbook.Path & "\" & baseName & "_" & Format(Now, "yyyymmddhhmmss") & "." & extensionName)

    CreateBackupFile = True

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg("ファイルのバックアップに失敗しました。" & vbCrLf & Err.Description, Err.Number, "CreateBackupFile")

End Function
'----------------------------------------------------------------------------------------------------