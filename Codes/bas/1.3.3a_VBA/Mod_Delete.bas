Attribute VB_Name = "Mod_Delete"
'1.3.3a_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2021/11/06 13:48:05
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
Public Sub DeleteAttr(ByVal spec As String, ByVal attr As FileAttribute)
On Error GoTo Err

    With FSO

        '���� ��r(�r�b�g���Z)
        Dim afterAttr As FileAttribute
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then

            Dim attrFolder As Folder
            Set attrFolder = .GetFolder(spec)
            '�t�H���_�[���� ��r(�r�b�g���Z)
            afterAttr = attrFolder.Attributes Xor attr
            attrFolder.Attributes = afterAttr

        Else

            Dim attrFile As File
            Set attrFile = .GetFile(spec)
            '�t�@�C������ ��r(�r�b�g���Z)
            afterAttr = .GetFile(spec).Attributes Xor attr
            attrFile.Attributes = afterAttr

        End If

    End With

    Exit Sub

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "DeleteAttr")

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2025/01/17 06:21:49
'----------------------------------------------------------------------------------------------------
'�Q�Ɛݒ�
'LibraryName:VBIDE
'LIBID:{0002E157-0000-0000-C000-000000000046}
    'ReferenceName:Microsoft Visual Basic for Applications Extensibility 5.3
    'FullPath(win32):C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
    'Major.Minor:5.3
Public Sub DeleteCodeModule(ByVal VBCName As String, ByVal VBCType As vbext_ComponentType, Optional ByRef parent As VBComponents)
On Error GoTo Err

    '���� ����l����
    If parent Is Nothing Then Set parent = ThisWorkbook.VBProject.VBComponents

    '�SVBComponent �m�F
    Dim VBC As VBComponent
    For Each VBC In parent
        If VBC.Name = VBCName And VBC.Type = VBCType Then
            '�R�[�h���W���[�� �S�폜
            Dim VBCCM As CodeModule
            Set VBCCM = VBC.CodeModule
            Call VBCCM.DeleteLines(1, VBCCM.CountOfLines)

            Exit Sub
        End If
    Next VBC

    Exit Sub

'�G���[����
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "DeleteCodeModule")

End Sub
'----------------------------------------------------------------------------------------------------