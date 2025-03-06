Attribute VB_Name = "Mod_Update"
'1.3.3a_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2021/07/26 10:48:35
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
Public Sub AddAttr(ByVal spec As String, ByVal attr As FileAttribute)
On Error GoTo Err

    With FSO

        '���� ��r(�r�b�g���Z)
        Dim afterAttr As FileAttribute
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then

            Dim attrFolder As Folder
            Set attrFolder = .GetFolder(spec)
            '�t�H���_�[���� ��r(�r�b�g���Z)
            afterAttr = attrFolder.Attributes Or attr
            attrFolder.Attributes = afterAttr

        Else

            Dim attrFile As File
            Set attrFile = .GetFile(spec)
            '�t�@�C������ ��r(�r�b�g���Z)
            afterAttr = .GetFile(spec).Attributes Or attr
            attrFile.Attributes = afterAttr

        End If

    End With

    Exit Sub

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "AddAttr")

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/05 10:56:29
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Sub CopyVBComponent(ByRef source As VBProject, ByRef destination As VBProject, ByVal VBCType As vbext_ComponentType, _
                           ByVal sourceVBCName As String, Optional ByVal destinationVBCName As String, Optional ByVal overwrite As Boolean)
On Error GoTo Err

    If source Is destination Then
        Call ShowErrMsg("�R�s�[���ƃR�s�[���VBProject�I�u�W�F�N�g�������ł��B", title:="CopyVBComponent")
        Exit Sub
    End If

    '�R�s�[��VBComponents�� �󔒊m�F
    If destinationVBCName = "" Then destinationVBCName = sourceVBCName

    '�R�s�[��VBComponents ���݊m�F
    If Not HasVBComponent(sourceVBCName, VBCType, source.VBComponents) Then
        Call ShowErrMsg("�R�s�[���Ɏw�肵��VBComponent�������݂��܂���B", title:="CopyVBComponent")
        Exit Sub
    End If

    '�R�s�[��VBComponents ���݊m�F
    Dim destinationVBC As VBComponent
    If HasVBComponent(destinationVBCName, VBCType, destination.VBComponents) Then
        If Not overwrite Then
            Call ShowErrMsg("�w�肵��VBComponent���͊��ɃR�s�[��ɑ��݂��Ă��܂��B", title:="CopyVBComponent")
            Exit Sub
        End If

        Set destinationVBC = destination.VBComponents(destinationVBCName)
    Else
        If VBCType = vbext_ct_ActiveXDesigner Or VBCType = vbext_ct_Document Then
            Call ShowErrMsg("�w�肵���^��VBComponent�I�u�W�F�N�g�͏㏑���ł̂݃R�s�[�\�ł��B", title:="CopyVBComponent")
            Exit Sub
        End If

        Set destinationVBC = destination.VBComponents.Add(VBCType)
        destinationVBC.Name = destinationVBCName
    End If

    'VBComponent CodeModule�R�s�[
    Dim moduleCode As String
    With source.VBComponents(sourceVBCName).CodeModule
        moduleCode = .Lines(1, .CountOfLines)
    End With
    With destinationVBC.CodeModule
        Call .DeleteLines(1, .CountOfLines)
        Call .AddFromString(moduleCode)
    End With

    Exit Sub

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "CopyVBComponent")

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Function ConvColAddrRef(ByVal colAddr As String, ByVal toReferenceStyle As XlReferenceStyle) As String
On Error Resume Next

    '�Q�ƌ`�� ����
    Dim colAddrRef As String
    If toReferenceStyle = xlA1 Then
        colAddrRef = Split(Cells(1, CLng(colAddr)).Address(ReferenceStyle:=toReferenceStyle), "$")(1)
    Else
        colAddrRef = Range(colAddr & "1").Column
    End If

    If colAddrRef <> "" Then
        ConvColAddrRef = colAddrRef
    Else
        Call ShowErrMsg("�w�肳�ꂽ�l������������܂���B", title:="ConvColAddrRef")
    End If

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/21 02:08:32
'----------------------------------------------------------------------------------------------------
Public Function ConvTrueFalse(ByVal expression As String) As Long
On Error GoTo Err

    '�^�U����p�z�� �쐬
    Dim trueArr() As Variant, falseArr() As Variant
    trueArr = Array("True", "T", "Yes", "Y", "�͂�", "����")
    falseArr = Array("False", "F", "No", "N", "������", "���Ȃ�")

    Dim t As Variant, f As Variant
    'True ����
    For Each t In trueArr
        If StrComp(expression, t, vbTextCompare) = 0 Then
            ConvTrueFalse = True
            Exit Function
        End If
    Next t
    'False ����
    For Each f In falseArr
        If StrComp(expression, f, vbTextCompare) = 0 Then
            ConvTrueFalse = False
            Exit Function
        End If
    Next f

    ConvTrueFalse = 1

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "ConvTrueFalse")

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
        'ProgID:Scripting.Dictionary
        'CLSID:{EE09B103-97E0-11CF-978F-00A02463E06F}
Public Function CopySheet(ByRef copySh As Object, Optional ByRef before As Variant, Optional ByRef after As Variant) As Object
On Error GoTo Err

    Dim sheetsFLG As Boolean

    '�V�[�g����
    Select Case TypeName(copySh)
        'before�V�[�g�w��
        Case "Worksheet", "Chart", "DialogSheet"
        Case "Sheets"
            sheetsFLG = True
        Case "Nothing"
            Call ShowErrMsg("copySh�I�u�W�F�N�g����ł��B", title:="CopySheet")
            Exit Function
        Case Else
            Call ShowErrMsg("copySh�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="CopySheet")
            Exit Function
    End Select

    Dim wb As Workbook
    'before, after���w��
    If IsMissing(before) And IsMissing(after) Then
        '�V�[�g �V�K�u�b�N�R�s�[
        Call copySh.Copy
        'SheetsFLG ����
        If sheetsFLG Then
            Set CopySheet = ActiveWorkbook.Sheets
        Else
            Set CopySheet = ActiveWorkbook.ActiveSheet
        End If

        Exit Function

    'after���w��
    ElseIf IsMissing(after) Then
        Select Case TypeName(before)
            'before�V�[�g�w��
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = before.Parent
            Case "Nothing"
                Call ShowErrMsg("before�I�u�W�F�N�g����ł��B", title:="CopySheet")
                Exit Function
            Case Else
                Call ShowErrMsg("before�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="CopySheet")
                Exit Function
        End Select

    'before���w��
    ElseIf IsMissing(before) Then
        Select Case TypeName(after)
            'after�V�[�g�w��
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = after.Parent
            Case "Nothing"
                Call ShowErrMsg("after�I�u�W�F�N�g����ł��B", title:="CopySheet")
                Exit Function
            Case Else
                Call ShowErrMsg("after�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="CopySheet")
                Exit Function
        End Select

    'before, after�w��
    Else
        Call ShowErrMsg("before��after�͓����Ɏw��ł��܂���B", title:="CopySheet")
        Exit Function
    End If

    '�R�s�[�惏�[�N�u�b�N ����
    Dim shs As Sheets
    If wb.Name = ThisWorkbook.Name Then
        Set shs = ThisWorkbook.Sheets
    Else
        Set shs = wb.Sheets
    End If

    '�R�s�[�O�V�[�g��, �C���f�b�N�X �f�B�N�V���i���[�i�[
    Dim shDIC As New Dictionary
    Dim befSh As Object
    For Each befSh In shs
        Call shDIC.Add(befSh.Name, befSh.Index)
    Next befSh

    '�V�[�g �R�s�[
    If IsMissing(after) Then
        Call copySh.Copy(Before:=before)
    Else
        Call copySh.Copy(After:=after)
    End If

    '�R�s�[�V�[�g�C���f�b�N�X �z��i�[
    Dim shIndxArr() As Variant
    shIndxArr = Array()
    Dim aftSh As Object
    For Each aftSh In shs
        If Not shDIC.Exists(aftSh.Name) Then
            ReDim Preserve shIndxArr(UBound(shIndxArr) + 1)
            shIndxArr(UBound(shIndxArr)) = aftSh.Index
        End If
    Next aftSh

    '�z��v�f�� ����
    If UBound(shIndxArr) > -1 Then
        'SheetsFLG ����
        If sheetsFLG Then
            Set CopySheet = wb.Sheets(shIndxArr)
        Else
            Set CopySheet = wb.Sheets(shIndxArr)(1)
        End If
    End If

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "CopySheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/18 15:41:51
'----------------------------------------------------------------------------------------------------
Public Function Max(ParamArray expressions() As Variant) As Variant
On Error GoTo Err

    If IsMissing(expressions) Then
        Call Err.Raise(450)
    End If

    Dim expression As Variant
    For Each expression In expressions
        '���t �V���A���l�ϊ�
        If IsDate(expression) Then expression = CDbl(expression)
        '���l��r
        If IsNumeric(expression) And Not IsEmpty(expression) Then
            Max = IIf(IsEmpty(Max), expression, IIf(Max > expression, Max, expression))
        End If
    Next expression

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "Max")
    Max = Null

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/18 15:41:51
'----------------------------------------------------------------------------------------------------
Public Function Min(ParamArray expressions() As Variant) As Variant
On Error GoTo Err

    If IsMissing(expressions) Then
        Call Err.Raise(450)
    End If

    Dim expression As Variant
    For Each expression In expressions
        '���t �V���A���l�ϊ�
        If IsDate(expression) Then expression = CDbl(expression)
        '���l��r
        If IsNumeric(expression) And Not IsEmpty(expression) Then
            Min = IIf(IsEmpty(Min), expression, IIf(Min < expression, Min, expression))
        End If
    Next expression

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "Min")
    Min = Null

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
        'ProgID:Scripting.Dictionary
        'CLSID:{EE09B103-97E0-11CF-978F-00A02463E06F}
Public Function MoveSheet(ByRef moveSh As Object, Optional ByRef before As Variant, Optional ByRef after As Variant) As Object
On Error GoTo Err

    Dim sheetsFLG As Boolean

    '�V�[�g����
    Select Case TypeName(moveSh)
        'before�V�[�g�w��
        Case "Worksheet", "Chart", "DialogSheet"
        Case "Sheets"
            sheetsFLG = True
        Case "Nothing"
            Call ShowErrMsg("moveSh�I�u�W�F�N�g����ł��B", title:="MoveSheet")
            Exit Function
        Case Else
            Call ShowErrMsg("moveSh�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="MoveSheet")
            Exit Function
    End Select

    Dim wb As Workbook
    'before, after���w��
    If IsMissing(before) And IsMissing(after) Then
        '�V�[�g �V�K�u�b�N�ړ�
        Call moveSh.Move
        'SheetsFLG ����
        If sheetsFLG Then
            Set MoveSheet = ActiveWorkbook.Sheets
        Else
            Set MoveSheet = ActiveWorkbook.ActiveSheet
        End If

        Exit Function

    'after���w��
    ElseIf IsMissing(after) Then
        Select Case TypeName(before)
            'before�V�[�g�w��
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = before.Parent
            Case "Nothing"
                Call ShowErrMsg("before�I�u�W�F�N�g����ł��B", title:="MoveSheet")
                Exit Function
            Case Else
                Call ShowErrMsg("before�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="MoveSheet")
                Exit Function
        End Select

    'before���w��
    ElseIf IsMissing(before) Then
        Select Case TypeName(after)
            'after�V�[�g�w��
            Case "Worksheet", "Chart", "DialogSheet"
                Set wb = after.Parent
            Case "Nothing"
                Call ShowErrMsg("after�I�u�W�F�N�g����ł��B", title:="MoveSheet")
                Exit Function
            Case Else
                Call ShowErrMsg("after�ɂ̓V�[�g�I�u�W�F�N�g���w�肵�Ă��������B", title:="MoveSheet")
                Exit Function
        End Select

    'before, after�w��
    Else
        Call ShowErrMsg("before��after�͓����Ɏw��ł��܂���B", title:="MoveSheet")
        Exit Function
    End If

    '�ړ��惏�[�N�u�b�N ����
    Dim shs As Sheets
    If wb.Name = ThisWorkbook.Name Then
        Set shs = ThisWorkbook.Sheets
    Else
        Set shs = wb.Sheets
    End If

    '�ړ��O�V�[�g��, �C���f�b�N�X �f�B�N�V���i���[�i�[
    Dim shDIC As New Dictionary
    Dim befSh As Object
    For Each befSh In shs
        Call shDIC.Add(befSh.Name, befSh.Index)
    Next befSh

    '�V�[�g �ړ�
    If IsMissing(after) Then
        Call moveSh.Move(Before:=before)
    Else
        Call moveSh.Move(After:=after)
    End If

    '�ړ��惏�[�N�u�b�N ����
    If wb.Name = ThisWorkbook.Name Then
        Set MoveSheet = moveSh
        Exit Function
    End If

    '�ړ��V�[�g�C���f�b�N�X �z��i�[
    Dim shIndxArr() As Variant
    shIndxArr = Array()
    Dim aftSh As Object
    For Each aftSh In shs
        If Not shDIC.Exists(aftSh.Name) Then
            ReDim Preserve shIndxArr(UBound(shIndxArr) + 1)
            shIndxArr(UBound(shIndxArr)) = aftSh.Index
        End If
    Next aftSh

    '�z��v�f�� ����
    If UBound(shIndxArr) > -1 Then
        'SheetsFLG ����
        If sheetsFLG Then
            Set MoveSheet = wb.Sheets(shIndxArr)
        Else
            Set MoveSheet = wb.Sheets(shIndxArr)(1)
        End If
    End If

    Exit Function

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "MoveSheet")

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/27 03:14:45
'----------------------------------------------------------------------------------------------------
Public Function ReplaceCRLF(ByVal expression As String) As String
On Error Resume Next

    'Carriage Return, Line Feed �폜
    ReplaceCRLF = Replace(Replace(expression, vbCr, vbNullString), vbLf, vbNullString)

End Function
'----------------------------------------------------------------------------------------------------