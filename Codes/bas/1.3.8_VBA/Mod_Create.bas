Attribute VB_Name = "Mod_Create"
'1.3.8_VBA
Option Explicit
Option Private Module
'----------------------------------------------------------------------------------------------------
'2025/03/18 04:27:44
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/13 12:51:10
'----------------------------------------------------------------------------------------------------
Public Function ArrayByType(ByVal arrType As VbVarType, ParamArray elements() As Variant) As Variant
On Error Resume Next

    If IsMissing(elements) Then Exit Function

    Dim arrLB As Long, arrUB As Long
    arrLB = LBound(elements, 1)
    arrUB = UBound(elements, 1)

    Dim arrIndx As Long, element As Variant
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
            For arrIndx = arrLB To arrUB
                nullArr(arrIndx) = Null
            Next arrIndx
            ArrayByType = nullArr
        '2
        Case vbInteger
            Dim intArr() As Integer
            ReDim intArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                intArr(arrIndx) = CInt(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = intArr
        '3
        Case vbLong
            Dim lngArr() As Long
            ReDim lngArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                lngArr(arrIndx) = CLng(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = lngArr
        '4
        Case vbSingle
            Dim sngArr() As Single
            ReDim sngArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                sngArr(arrIndx) = CSng(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = sngArr
        '5
        Case vbDouble
            Dim dblArr() As Double
            ReDim dblArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                dblArr(arrIndx) = CDbl(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = dblArr
        '6
        Case vbCurrency
            Dim curArr() As Currency
            ReDim curArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                curArr(arrIndx) = CCur(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = curArr
        '7
        Case vbDate
            Dim dateArr() As Date
            ReDim dateArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                dateArr(arrIndx) = CDate(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = dateArr
        '8
        Case vbString
            Dim strArr() As String
            ReDim strArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                strArr(arrIndx) = CStr(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = strArr
        '9
        Case vbObject
            Dim objArr() As Object
            ReDim objArrArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                Set objArrArr(arrIndx) = elements(arrIndx)
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = objArrArr
        '11
        Case vbBoolean
            Dim boolArr() As Boolean
            ReDim boolArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                boolArr(arrIndx) = CBool(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = boolArr
        '12
        Case vbVariant
            ArrayByType = elements
        '14
        Case vbDecimal
            Dim decArr() As Variant
            ReDim decArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                decArr(arrIndx) = CDec(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = decArr
        '17
        Case vbByte
            Dim byteArr() As Byte
            ReDim byteArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                byteArr(arrIndx) = CByte(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = byteArr

        #If Win64 Then

        '20
        Case vbLongLong
            Dim lnglngArr() As LongLong
            ReDim lnglngArr(arrLB To arrUB)
            For arrIndx = arrLB To arrUB
                lnglngArr(arrIndx) = CLngLng(elements(arrIndx))
            Next arrIndx
            If Err.Number = 13 Then GoTo Err_Convert
            ArrayByType = lnglngArr

        #End If

        Case Else
            Call ShowErrMsg("指定されたデータ型は無効です。", title:="ArrayByType")
            Exit Function
    End Select

    Exit Function

'エラー処理
Err_Convert:

    Call ShowErrMsg("指定されたデータ型ではない要素が含まれています。", title:="ArrayByType")

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
Public Function BuildPath(ParamArray paths() As Variant) As String
On Error GoTo Err

    If IsMissing(paths) Then Exit Function

    Dim pathsElm As Variant
    For Each pathsElm In paths
        If Not IsArray(pathsElm) Then pathsElm = Array(pathsElm)

        Dim path As Variant
        For Each path In pathsElm
            BuildPath = FSO.BuildPath(BuildPath, CStr(path))
        Next path
    Next pathsElm

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "BuildPath")
    BuildPath = ""

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
Public Function CreateBackupFile(ByVal sourceFileSpec As String, Optional ByVal destinationFilePath As String) As Boolean
On Error GoTo Err

    If destinationFilePath = "" Then destinationFilePath = ThisWorkbook.Path

    With FSO
        'ファイル存在 確認
        If Not .FileExists(sourceFileSpec) Then
            Call ShowErrMsg("指定されたファイルは存在しません。", title:="CreateBackupFile")
            Exit Function
        End If

        'ファイル名, 拡張子 取得
        Dim baseName As String, extensionName As String
        baseName = .GetBaseName(sourceFileSpec)
        extensionName = .GetExtensionName(sourceFileSpec)
    End With

    Call FileCopy(sourceFileSpec, destinationFilePath & "\" & baseName & "_" & Format(Now, "yyyymmddhhmmss") & "." & extensionName)

    CreateBackupFile = True

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg("ファイルのバックアップに失敗しました。" & vbCrLf & Err.Description, Err.Number, "CreateBackupFile")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/03/07 00:03:06
'----------------------------------------------------------------------------------------------------
Public Function CreateCollection(ParamArray collItems()) As Collection
On Error GoTo Err

    If IsMissing(collItems) Then Exit Function

    Dim coll As New Collection
    Dim collItem As Variant
    For Each collItem In collItems
        Call coll.Add(collItem)
    Next collItem

    Set CreateCollection = coll

    Exit Function

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "CreateCollection")

End Function
'----------------------------------------------------------------------------------------------------