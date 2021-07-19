VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)

    Call ResetChangedFormat(Target)

End Sub
'----------------------------------------------------------------------------------------------------
