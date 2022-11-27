Attribute VB_Name = "Mod_Delete"
'1.0.4_VBA
Option Explicit
'----------------------------------------------------------------------------------------------------
'2021/11/06 13:48:05
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
Public Sub DeleteAttr(ByVal spec As String, ByVal attr As FileAttribute)
On Error GoTo Err

    With FSO

        '属性 比較(ビット演算)
        Dim afterAttr As FileAttribute
        If (GetAttr(spec) And vbDirectory) = vbDirectory Then

            Dim attrFolder As Folder
            Set attrFolder = .GetFolder(spec)
            'フォルダ属性 比較(ビット演算)
            afterAttr = attrFolder.Attributes Xor attr
            attrFolder.Attributes = afterAttr

        Else

            Dim attrFile As File
            Set attrFile = .GetFile(spec)
            'ファイル属性 比較(ビット演算)
            afterAttr = .GetFile(spec).Attributes Xor attr
            attrFile.Attributes = afterAttr

        End If

    End With

    Exit Sub

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number, "DeleteAttr")

End Sub
'----------------------------------------------------------------------------------------------------