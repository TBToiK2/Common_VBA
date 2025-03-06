Attribute VB_Name = "Mod_Read"
'1.3.3a_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2022/03/12 18:57:57
'----------------------------------------------------------------------------------------------------
Public Enum EnumFileNameType
    enFileName = 0
    enBaseName = 1
    enExtensionName = 2
End Enum
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 12:51:33
'----------------------------------------------------------------------------------------------------
Public Function ArrayValueExists(ByVal targetArr As Variant, ByVal searchValue As Variant) As Boolean
On Error GoTo Err

    '�^�[�Q�b�g���� �z�񔻒�
    If Not IsArray(targetArr) Then Exit Function

On Error Resume Next

    '�^�[�Q�b�g���� �󔒔z�񔻒�
    If UBound(targetArr) = -1 Then Exit Function

On Error GoTo Err

    '�������� �z��, �I�u�W�F�N�g����
    If IsArray(searchValue) Or IsObject(searchValue) Then Exit Function

    Dim elmValue As Variant
    '�S�l �m�F
    For Each elmValue In targetArr
        '�z��, �I�u�W�F�N�g ����
        If Not IsArray(elmValue) And Not IsObject(elmValue) Then
            If searchValue = elmValue Then
                ArrayValueExists = True
                Exit Function
            End If
        End If
    Next elmValue

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ArrayValueExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�t�@�C�� ���݊m�F
    FileExists = FSO.FileExists(fileSpec)

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "FileExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�t�H���_�[ ���݊m�F
    FolderExists = FSO.FolderExists(folderSpec)

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "FolderExists")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/28 15:47:52
'----------------------------------------------------------------------------------------------------
Public Function GetColRange(ByRef ws As Worksheet, ByVal targetCol As Long, Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Variant, _
                            Optional ByVal relativeFLG As Boolean) As Range
On Error GoTo Err

    If targetCol < 1 Or targetCol > MAX_COL Then
        Call ShowErrMsg("targetCol�Ɏw�肳�ꂽ�l���L���񐔂͈̔͊O�ł��B" & "1�ȏ�" & MAX_COL & "�ȉ��œ��͂��Ă��������B", title:="GetColRange")
        Exit Function
    End If
    If firstRow < 1 Or firstRow > MAX_ROW Then
        Call ShowErrMsg("firstRow�Ɏw�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "1�ȏ�" & MAX_ROW & "�ȉ��œ��͂��Ă��������B", title:="GetColRange")
        Exit Function
    End If

    Dim targetLastRow As Long
    If IsMissing(lastRow) Then

        '�ŏI�s �擾
        targetLastRow = GetLastRow(ws, targetCol)
        If firstRow > targetLastRow Then Exit Function

    Else

        If Not IsNumber(lastRow) Then
            Call ShowErrMsg("lastRow�Ɏw�肳�ꂽ�l�͐��l�ł͂���܂���B" & "1�ȏ�" & MAX_ROW & "�ȉ��œ��͂��Ă��������B", title:="GetColRange")
            Exit Function
        Else
            If relativeFLG Then
                '�ŏI�s �擾
                targetLastRow = firstRow + lastRow
                If targetLastRow < 1 Or targetLastRow > MAX_ROW Then
                    Call ShowErrMsg("�w�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "firstRow + lastRow��" & "1�ȏ�" & MAX_ROW & "�ȉ��ƂȂ�悤���͂��Ă��������B", title:="GetColRange")
                    Exit Function
                End If
            Else
                '�ŏI�s �擾
                targetLastRow = lastRow
                If targetLastRow < 1 Or targetLastRow > MAX_ROW Then
                    Call ShowErrMsg("lastRow�Ɏw�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "1�ȏ�" & MAX_ROW & "�ȉ��œ��͂��Ă��������B", title:="GetColRange")
                    Exit Function
                End If
            End If
        End If

    End If

    With ws
        Set GetColRange = .Range(.Cells(firstRow, targetCol), .Cells(targetLastRow, targetCol))
    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetColRange")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

        '�t�@�C���� �擾
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

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetFileName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/28 01:24:56
'----------------------------------------------------------------------------------------------------
Public Function GetLastCol(ByRef ws As Worksheet, ByVal targetRow As Long) As Long
On Error GoTo Err

    If targetRow < 1 Or targetRow > MAX_ROW Then
        Call ShowErrMsg("�w�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "1�ȏ�" & MAX_ROW & "�ȉ��œ��͂��Ă��������B", title:="GetLastCol")
        Exit Function
    End If

    With ws

        '�g�p�͈͍ŏI�� �擾
        With .UsedRange
            Dim usedLastCol As Long
            usedLastCol = .Columns(.Columns.Count).Column
        End With

        '�g�p�͈͔z�� �擾
        Dim usedRangeArr() As Variant
        With .Range(.Cells(targetRow, 1), .Cells(targetRow, usedLastCol))
            '�z�� ����
            If IsArray(.Value) Then
                usedRangeArr = .Value
            Else
                ReDim usedRangeArr(1 To 1, 1 To 1)
                usedRangeArr(1, 1) = .Value
            End If
        End With

        '�ŏI�� �擾
        Dim lastCol As Long
        For lastCol = UBound(usedRangeArr, 2) To 1 Step -1
            If CStr(usedRangeArr(1, lastCol)) <> "" Then
                GetLastCol = lastCol
                Exit Function
            End If
        Next lastCol

    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetLastCol")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/28 01:24:56
'----------------------------------------------------------------------------------------------------
Public Function GetLastRow(ByRef ws As Worksheet, ByVal targetCol As Long) As Long
On Error GoTo Err

    If targetCol < 1 Or targetCol > MAX_COL Then
        Call ShowErrMsg("�w�肳�ꂽ�l���L���񐔂͈̔͊O�ł��B" & "1�ȏ�" & MAX_COL & "�ȉ��œ��͂��Ă��������B", title:="GetLastRow")
        Exit Function
    End If

    With ws

        '�g�p�͈͍ŏI�s �擾
        With .UsedRange
            Dim usedLastRow As Long
            usedLastRow = .Rows(.Rows.Count).Row
        End With

        '�g�p�͈͔z�� �擾
        Dim usedRangeArr() As Variant
        With .Range(.Cells(1, targetCol), .Cells(usedLastRow, targetCol))
            '�z�� ����
            If IsArray(.Value) Then
                usedRangeArr = .Value
            Else
                ReDim usedRangeArr(1 To 1, 1 To 1)
                usedRangeArr(1, 1) = .Value
            End If
        End With

        '�ŏI�s �擾
        Dim lastRow As Long
        For lastRow = UBound(usedRangeArr, 1) To 1 Step -1
            If CStr(usedRangeArr(lastRow, 1)) <> "" Then
                GetLastRow = lastRow
                Exit Function
            End If
        Next lastRow

    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetLastRow")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function GetListObject(ByVal loName As String, Optional ByRef parent As ListObjects) As ListObject
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then
        '���[�N�V�[�g ����
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

    '�S���X�g�I�u�W�F�N�g �m�F
    Dim lo As ListObject
    For Each lo In parent
        If lo.Name = loName Then
            Set GetListObject = lo
            Exit Function
        End If
    Next lo

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetListObject")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function GetName(ByVal nameName As String, Optional ByRef parent As Names) As Name
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = ThisWorkbook.Names

    Dim defineName As Name
    'Worksheet ����
    If TypeName(parent.Parent) = "Worksheet" Then
        '�S���O �m�F
        For Each defineName In parent
            If defineName.Name = parent.Parent.Name & "!" & nameName Then
                Set GetName = defineName
                Exit Function
            End If
        Next defineName
    Else
        '�S���O �m�F
        For Each defineName In parent
            If defineName.Name = nameName Then
                Set GetName = defineName
                Exit Function
            End If
        Next defineName
    End If

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�p�X �擾
    GetPath = FSO.GetParentFolderName(spec)

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetPath")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/28 15:47:52
'----------------------------------------------------------------------------------------------------
Public Function GetRowRange(ByRef ws As Worksheet, ByVal targetRow As Long, Optional ByVal firstCol As Long = 1, Optional ByVal lastCol As Variant, _
                            Optional ByVal relativeFLG As Boolean) As Range
On Error GoTo Err

    If targetRow < 1 Or targetRow > MAX_ROW Then
        Call ShowErrMsg("targetRow�Ɏw�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "1�ȏ�" & MAX_ROW & "�ȉ��œ��͂��Ă��������B", title:="GetRowRange")
        Exit Function
    End If
    If firstCol < 1 Or firstCol > MAX_COL Then
        Call ShowErrMsg("firstCol�Ɏw�肳�ꂽ�l���L���񐔂͈̔͊O�ł��B" & "1�ȏ�" & MAX_COL & "�ȉ��œ��͂��Ă��������B", title:="GetRowRange")
        Exit Function
    End If

    Dim targetLastCol As Long
    If IsMissing(lastCol) Then

        '�ŏI�� �擾
        targetLastCol = GetLastCol(ws, targetRow)
        If firstCol > targetLastCol Then Exit Function

    Else

        If Not IsNumber(lastCol) Then
            Call ShowErrMsg("lastCol�Ɏw�肳�ꂽ�l�͐��l�ł͂���܂���B" & "1�ȏ�" & MAX_COL & "�ȉ��œ��͂��Ă��������B", title:="GetRowRange")
            Exit Function
        Else
            If relativeFLG Then
                '�ŏI�� �擾
                targetLastCol = firstCol + lastCol
                If targetLastCol < 1 Or targetLastCol > MAX_COL Then
                    Call ShowErrMsg("�w�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "firstCol + lastCol��" & "1�ȏ�" & MAX_COL & "�ȉ��ƂȂ�悤���͂��Ă��������B", title:="GetRowRange")
                    Exit Function
                End If
            Else
                '�ŏI�� �擾
                targetLastCol = lastCol
                If targetLastCol < 1 Or targetLastCol > MAX_COL Then
                    Call ShowErrMsg("lastCol�Ɏw�肳�ꂽ�l���L���s���͈̔͊O�ł��B" & "1�ȏ�" & MAX_COL & "�ȉ��œ��͂��Ă��������B", title:="GetRowRange")
                    Exit Function
                End If
            End If
        End If

    End If

    With ws
        Set GetRowRange = .Range(.Cells(targetRow, firstCol), .Cells(targetRow, targetLastCol))
    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetRowRange")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

        '���� ���
        searchNames = Replace(searchNames, ";", ",")
        Dim searchNameArr() As String
        searchNameArr = Split(searchNames, ",")

        Dim targetFiles As Files, targetFile As File
        Set targetFiles = .GetFolder(searchFilePath).Files

        If targetFiles.Count = 0 Then Exit Function

        '�����t�@�C���� �J�E���g
        Dim searchName As Variant
        Dim fileCount As Long
        Select Case enFileNameType
            '�t�@�C����
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

            '�x�[�X��
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

            '�g���q��
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

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSearchFileCount")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

        '���� ���
        searchNames = Replace(searchNames, ";", ",")
        Dim searchNameArr() As String
        searchNameArr = Split(searchNames, ",")

        Dim targetFiles As Files, targetFile As File
        Set targetFiles = .GetFolder(searchFilePath).Files

        If targetFiles.Count = 0 Then Exit Function

        '�����t�@�C���f�B���N�g���[ �擾
        Dim fileSpecArr() As String
        Dim searchName As Variant
        Dim fileCount As Long
        Select Case enFileNameType
            '�t�@�C����
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

            '�x�[�X��
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

            '�g���q��
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

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSearchFileSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�t�@�C���p�X ���݊m�F
    If Not FSO.FolderExists(filePath) Then
        Call ShowErrMsg("�w�肳�ꂽ�t�@�C���p�X�����݂��܂���B", title:="GetSelectFileSpec")
        Exit Function
    End If

    '���� �󔒊m�F
    If dialogTitle = "" Then dialogTitle = "�t�@�C�� �I��"
    If filterDescriptions = "" Then filterDescriptions = "���ׂẴt�@�C��"
    If filterExtensions = "" Then filterExtensions = "*.*"

    Dim filterDescriptionArr() As String, filterExtensionArr() As String
    filterDescriptionArr = Split(filterDescriptions, ",")
    filterExtensionArr = Split(filterExtensions, ",")

    '�z�� �v�f����r
    Dim filterDescriptionArrUpper As Long, filterExtensionArrUpper As Long
    filterDescriptionArrUpper = UBound(filterDescriptionArr, 1)
    filterExtensionArrUpper = UBound(filterExtensionArr, 1)
    If filterDescriptionArrUpper <> filterExtensionArrUpper Then
        Call ShowErrMsg("�w�肳�ꂽ�t�@�C���̎�ނƊg���q�̐�����v���܂���B", title:="GetSelectFileSpec")
        Exit Function
    End If

    Dim filterArrUpper As Long
    filterArrUpper = filterDescriptionArrUpper

    '���� ���
    Dim filterArrIndx As Long
    For filterArrIndx = 0 To filterArrUpper
        '����
        If filterDescriptionArr(filterArrIndx) = "" Then filterDescriptionArr(filterArrIndx) = "���ׂẴt�@�C��"

        '�g���q
        Dim bufArr As Variant, bufArrIndx As Long
        bufArr = Split(IIf(filterExtensionArr(filterArrIndx) = "", "*", filterExtensionArr(filterArrIndx)), ";")
        For bufArrIndx = 0 To UBound(bufArr, 1)
            If bufArr(bufArrIndx) = "" Then bufArr(bufArrIndx) = "*"
        Next bufArrIndx
        filterExtensionArr(filterArrIndx) = "*." & Join(bufArr, ";*.")
    Next filterArrIndx

    '�t�B���^�[���� �m�F
    If Not filterSeparateFLG Then
        filterDescriptionArr(0) = Join(filterDescriptionArr, ",")
        ReDim Preserve filterDescriptionArr(0)
        filterExtensionArr(0) = Join(filterExtensionArr, ";")
        ReDim Preserve filterExtensionArr(0)

        filterArrUpper = 0
    End If

    '�_�C�A���O�{�b�N�X �ݒ�
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

        '�t�@�C���I����� ����
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

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSelectFileSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�t�H���_�[�p�X ���݊m�F
    If Not FSO.FolderExists(folderPath) Then
        Call ShowErrMsg("�w�肳�ꂽ�t�H���_�[�p�X�����݂��܂���B", title:="GetSelectFolderSpec")
        Exit Function
    End If

    '���� �󔒊m�F
    If dialogTitle = "" Then dialogTitle = "�t�H���_�[ �I��"

    '�_�C�A���O�{�b�N�X �ݒ�
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = dialogTitle
        .InitialFileName = FSO.BuildPath(folderPath, Application.PathSeparator)

        '�t�H���_�[�I����� ����
        If .Show = -1 Then GetSelectFolderSpec = .SelectedItems(1)
    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSelectFolderSpec")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 10:22:09
'----------------------------------------------------------------------------------------------------
Public Function GetShape(ByVal shapeName As String, Optional ByRef parent As Shapes) As Shape
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then
        '���[�N�V�[�g ����
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

    '�S�V�F�C�v �m�F
    Dim shp As Shape
    For Each shp In parent
        If shp.Name = shapeName Then
            Set GetShape = shp
            Exit Function
        End If
    Next shp

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetShape")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function GetSheet(ByVal shName As String, Optional ByRef parentWb As Workbook, Optional ByVal shType As XlSheetType, Optional ByVal codeNameFLG As Boolean) As Object
On Error GoTo Err

    '���� ����l����
    If parentWb Is Nothing Then Set parentWb = ThisWorkbook

    Dim parent As Sheets
    Dim sh As Object
    Dim targetName As String
    '�V�[�g�^�C�v ����
    Select Case shType
        Case xlChart, xlWorksheet
            Set parent = IIf(shType = xlChart, parentWb.Charts, parentWb.Worksheets)
            '�S�V�[�g �m�F
            For Each sh In parent
                targetName = IIf(codeNameFLG, sh.CodeName, sh.Name)
                If targetName = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh

        Case xlDialogSheet, xlExcel4IntlMacroSheet, xlExcel4MacroSheet
            Set parent = IIf(shType = xlDialogSheet, parentWb.DialogSheets, IIf(shType = xlExcel4IntlMacroSheet, parentWb.Excel4IntlMacroSheets, parentWb.Excel4MacroSheets))
            '�S�V�[�g �m�F
            For Each sh In parent
                If sh.Name = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh

        Case Else
            '�S�V�[�g �m�F
            For Each sh In parentWb.Sheets
                If sh.Name = shName Then
                    Set GetSheet = sh
                    Exit Function
                End If
            Next sh

    End Select

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 06:06:07
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Function GetVBComponent(ByVal VBCName As String, ByVal VBCType As vbext_ComponentType, Optional ByRef parent As VBComponents) As VBComponent
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = ThisWorkbook.VBProject.VBComponents

    '�SVBComponent �m�F
    Dim VBC As VBComponent
    For Each VBC In parent
        If VBC.Name = VBCName And VBC.Type = VBCType Then
            Set GetVBComponent = VBC
            Exit Function
        End If
    Next VBC

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "GetVBComponent")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

        '���� ��r(�r�b�g���Z)
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then

            '�t�H���_�[���� ��r(�r�b�g���Z)
            HasAttr = (.GetFolder(spec).Attributes And attr) = attr

        Else
            '�t�@�C������ ��r(�r�b�g���Z)
            HasAttr = (.GetFile(spec).Attributes And attr) = attr

        End If

    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasAttr")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasListObject(ByVal loName As String, Optional ByRef parent As ListObjects) As Boolean
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then
        '���[�N�V�[�g ����
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

    '�S���X�g�I�u�W�F�N�g �m�F
    Dim lo As ListObject
    For Each lo In parent
        If lo.Name = loName Then
            HasListObject = True
            Exit Function
        End If
    Next lo

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasListObject")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasName(ByVal nameName As String, Optional ByRef parent As Names) As Boolean
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = ThisWorkbook.Names

    Dim defineName As Name
    'Worksheet ����
    If TypeName(parent.Parent) = "Worksheet" Then
        '�S���O �m�F
        For Each defineName In parent
            If defineName.Name = parent.Parent.Name & "!" & nameName Then
                HasName = True
                Exit Function
            End If
        Next defineName
    Else
        '�S���O �m�F
        For Each defineName In parent
            If defineName.Name = nameName Then
                HasName = True
                Exit Function
            End If
        Next defineName
    End If

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasName")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 10:28:43
'----------------------------------------------------------------------------------------------------
Public Function HasShape(ByVal shapeName As String, Optional ByRef parent As Shapes) As Boolean
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then
        '���[�N�V�[�g ����
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

    '�S�V�F�C�v �m�F
    Dim shp As Shape
    For Each shp In parent
        If shp.Name = shapeName Then
            HasShape = True
            Exit Function
        End If
    Next shp

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasShape")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasSheet(ByVal shName As String, Optional ByRef parentWb As Workbook, Optional ByVal shType As XlSheetType, Optional ByVal codeNameFLG As Boolean) As Boolean
On Error GoTo Err

    '���� ����l����
    If parentWb Is Nothing Then Set parentWb = ThisWorkbook

    Dim parent As Sheets
    Dim sh As Object
    Dim targetName As String
    '�V�[�g�^�C�v ����
    Select Case shType
        Case xlChart, xlWorksheet
            Set parent = IIf(shType = xlChart, parentWb.Charts, parentWb.Worksheets)
            '�S�V�[�g �m�F
            For Each sh In parent
                targetName = IIf(codeNameFLG, sh.CodeName, sh.Name)
                If targetName = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh

        Case xlDialogSheet, xlExcel4IntlMacroSheet, xlExcel4MacroSheet
            Set parent = IIf(shType = xlDialogSheet, parentWb.DialogSheets, IIf(shType = xlExcel4IntlMacroSheet, parentWb.Excel4IntlMacroSheets, parentWb.Excel4MacroSheets))
            '�S�V�[�g �m�F
            For Each sh In parent
                If sh.Name = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh

        Case Else
            '�S�V�[�g �m�F
            For Each sh In parentWb.Sheets
                If sh.Name = shName Then
                    HasSheet = True
                    Exit Function
                End If
            Next sh

    End Select

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Function HasVBComponent(ByVal VBCName As String, ByVal VBCType As vbext_ComponentType, Optional ByRef parent As VBComponents) As Boolean
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = ThisWorkbook.VBProject.VBComponents

    '�SVBComponent �m�F
    Dim VBC As VBComponent
    For Each VBC In parent
        If VBC.Name = VBCName And VBC.Type = VBCType Then
            HasVBComponent = True
            Exit Function
        End If
    Next VBC

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasVBComponent")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function HasWorkbook(ByVal wbName As String, Optional ByRef parent As Workbooks) As Boolean
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = Workbooks

    '�S���[�N�u�b�N �m�F
    Dim wb As Workbook
    For Each wb In parent
        If wb.Name = wbName Then
            HasWorkbook = True
            Exit Function
        End If
    Next wb

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "HasWorkbook")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�� ����
    If expression = "" Then Exit Function

    '���K�\�� ����
    With REG
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = "^[A-Za-z�`-�y��-��]+$"
        IsAlphabet = .Test(expression)
    End With

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "IsAlphabet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function IsBlank(ByVal expression As Variant) As Boolean
On Error GoTo Err

    'Null����
    If IsNull(expression) Then
        IsBlank = True
        Exit Function
    '�I�u�W�F�N�g����
    ElseIf IsObject(expression) Then
        '�󔻒�
        If Not expression Is Nothing Then Exit Function
    '�z�񔻒�
    ElseIf IsArray(expression) Then
On Error GoTo ArrayErr
        '�󔻒�
        expression = UBound(expression)
        Exit Function
ArrayErr:
On Error GoTo Err
    Else
        '�󔻒�
        If CStr(expression) <> "" Then Exit Function
    End If

    IsBlank = True

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "IsBlank")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2024/11/28 16:42:20
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
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

    '�� ����
    If expression = "" Then Exit Function

    '���K�\�� ����
    With REG
        .Global = False
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = "^-?\d*\.?\d*$"
        IsNumber = .Test(expression)
    End With

    Exit Function

'�G���[����
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

    '�t�@�C�����݊m�F
    If Not FSO.FileExists(fileSpec) Then Exit Function

    Dim fileNumber As Long
    '�t�@�C���i���o�[1-255
    fileNumber = FreeFile(0)
    '�t�@�C���i���o�[256-512
    If fileNumber = 0 Then fileNumber = FreeFile(1)

    '�t�@�C���i���o�[ �S�g�p��
    If Err.Number = 67 Then Exit Function

    '�J�m�F(Append)
    Open fileSpec For Append As #fileNumber
    Close #fileNumber

    '�ǂݎ���p����
    If Err.Number =  75 Then
        '�J�m�F(Input)
        Open fileSpec For Input As #fileNumber
        Close #fileNumber
    End If

    IsOpen = Err.Number = 55 Or Err.Number = 70

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function IsString(ByVal expression As Variant) As Boolean
On Error Resume Next

    Dim str As String
    str = CStr(expression)

    IsString = Not Err.Number > 0

End Function
'----------------------------------------------------------------------------------------------------