VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Workbook_Open()

    Application.ScreenUpdating = False

    '�N���� �����ݒ�
    With Main

        '�V�[�g �v���e�N�g����
        Call .Unprotect(PASS)

        .Activate
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False

        '�Ώ۔͈� �\��
        .Range(.Rows(1), .Rows(87)).Hidden = False
        
        '�\���A�C�R�� �p�x�ݒ�
        .Shapes("Tri_Visible_Summary").Rotation = 0
        .Shapes("Tri_Visible_Param").Rotation = 0
        .Shapes("Tri_Visible_Copy").Rotation = 0
        .Shapes("Tri_Visible_Use").Rotation = 0

        '�V�[�g �v���e�N�g
        .Cells(1, 19).Activate
        .EnableSelection = xlUnlockedCells
        Call .Protect(PASS, UserInterfaceOnly:=True)

    End With

    Application.ScreenUpdating = True

End Sub
'----------------------------------------------------------------------------------------------------