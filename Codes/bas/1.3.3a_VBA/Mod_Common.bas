Attribute VB_Name = "Mod_Common"
'1.3.3a_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2023/02/22 17:10:31
'----------------------------------------------------------------------------------------------------
Public Const MIN_ROW = 1
Public Const MAX_ROW = 1048576
Public Const MIN_COL = 1
Public Const MAX_COL = 16384
Public FSO As New FileSystemObject
Public REG As New RegExp
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 15:31:21
'----------------------------------------------------------------------------------------------------
Public Sub AfterProcess(Optional ByVal calculation As XlCalculation, Optional ByRef excelApp As Excel.Application, Optional ByVal isWbOpening As Boolean)

    '���� ����l����
    If excelApp Is Nothing Then Set excelApp = Excel.Application

    '�e�v���Z�X �ĊJ
    With excelApp
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        If Not isWbOpening Then
            .CutCopyMode = False
            .StatusBar = False
        End If
        If .Workbooks.Count > 0 Then
            .Calculation = calculation
        End If
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/12/06 15:31:21
'----------------------------------------------------------------------------------------------------
Public Sub BeforeProcess(Optional ByRef calculation As XlCalculation, Optional ByRef excelApp As Excel.Application, Optional ByVal isWbOpening As Boolean)

    '���� ����l����
    If excelApp Is Nothing Then Set excelApp = Excel.Application

    '�e�v���Z�X ��~
    With excelApp
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        If Not isWbOpening Then
            .StatusBar = False
        End If
        If .Workbooks.Count > 0 Then
            calculation = .Calculation
            .Calculation = xlCalculationManual
        End If
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/07/27 00:58:07
'----------------------------------------------------------------------------------------------------
Public Sub ShowAllData(ByRef ws As Worksheet)
On Error Resume Next

    With ws
        '�t�B���^�[ ����
        If .FilterMode Then Call .ShowAllData

        '�s�� �S�\��
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/03/17 09:59:52
'----------------------------------------------------------------------------------------------------
Public Sub ShowErrMsg(ByVal errDescription As String, Optional ByVal errNumber As Long, Optional ByVal title As String)
On Error Resume Next

    '���b�Z�[�W�v�����v�g �ݒ�
    Dim prompt As String
    prompt = "�G���[���e:[" & vbCrLf & errDescription & vbCrLf & "]"
    If errNumber <> 0 Then Prompt = "�G���[�ԍ�:[" & errNumber & "]" & vbCrLf & prompt

    '�^�C�g�� �ݒ�
    If title <> "" Then title = ":" & title

    '���b�Z�[�W �\��
    Call MsgBox(prompt, vbOKOnly + vbCritical, "�G���[" & title)

    Err.Clear

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/29 04:11:11
'----------------------------------------------------------------------------------------------------
Public Sub ShowInfoMsg(ByVal prompt As String, Optional ByVal title As String)
On Error Resume Next

    '�^�C�g�� �ݒ�
    If title <> "" Then title = ":" & title

    '���b�Z�[�W �\��
    Call MsgBox(prompt, vbOKOnly + vbInformation, "���" & title)

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'2022/06/29 04:11:11
'----------------------------------------------------------------------------------------------------
Public Function ShowQuestionMsg(ByVal prompt As String, Optional ByVal title As String) As VbMsgBoxResult
On Error Resume Next

    '�^�C�g�� �ݒ�
    If title <> "" Then title = ":" & title

    '���b�Z�[�W �\��
    ShowQuestionMsg = MsgBox(prompt, vbOKCancel + vbQuestion, "�m�F" & title)

End Function
'----------------------------------------------------------------------------------------------------