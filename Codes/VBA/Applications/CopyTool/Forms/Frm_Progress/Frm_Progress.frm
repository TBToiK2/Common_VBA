VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Progress 
   Caption         =   "処理中"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "Frm_Progress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

  'プログレスバー 初期化
  Lbl_ProgressBar.Width = 0

End Sub
'----------------------------------------------------------------------------------------------------
