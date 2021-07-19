VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Main 
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14565
   OleObjectBlob   =   "Frm_Main.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------
'確認メッセージ
Private Const INF_BACK_ON  As String = "バックグラウンドで処理を行います。よろしいですか？"
Private Const INF_BACK_OFF As String = "処理を行います。" & vbCrLf & _
                                        "「セキュリティ警告」ダイアログが表示された場合は、必ず" & vbCrLf & _
                                        "「マクロを有効にする」を選択してください。" & vbCrLf & _
                                        "※「マクロを無効にする」を選択した場合、処理が止まります。" & vbCrLf & _
                                        "よろしいですか？"

'共通パラメータ
Private g_ExcelApp      As Application 'コピー用Excelオブジェクト
Private g_ParamSh       As Worksheet   'パラメータ設定シート
Private g_ParamShName   As String      'パラメータ設定シート名
Private g_BackgroundFLG As Boolean     'バックグラウンド処理
Private g_LogFLG        As Boolean     '更新ログ作成
Private g_FormatFLG     As Boolean     '書式コピー
Private g_BlankFLG      As Boolean     '空白コピー
Private g_AutoSaveFLG   As Boolean     'コピー先オートセーブ
Private g_BackupFLG     As Boolean     'コピー先バックアップ
Private g_ChangeMarkCol As Long        'コピー先変更行マーク
Private g_FontColor     As Long        'コピーセル文字色
Private g_InteriorColor As Long        'コピーセル背景色

'ファイル別パラメータ (F→コピー元, T→コピー先)
Private g_FWb        As Workbook, g_TWb     As Workbook 'コピー先ワークブック
Private g_FFileSpec  As String, g_TFileSpec As String   'ファイルパス
Private g_FFileName  As String, g_TFileName As String   'ファイル名
Private g_FShName    As String, g_TShName   As String   'シート名
Private g_FPswd      As String, g_TPswd     As String   'パスワード
Private g_FItemIDRow As Long, g_TItemIDRow  As Long     '項目ID行
Private g_FStartCol  As Long, g_TStartCol   As Long     '項目開始列
Private g_FDataIDCol As Long, g_TDataIDCol  As Long     'データID列
Private g_FStartRow  As Long, g_TStartRow   As Long     'データ開始行

'コピー対象テーブル
Private g_CopyItemArr() As Variant           'コピー対象項目配列
Private Const COPYCOL_START_ROW As Long = 22 'コピー対象項目先頭行
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Const GWL_STYLE = -16
Private Const WS_SYSMENU = &H80000

#If VBA7 Then

    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32.dll" Alias "SetWindowLongPtrA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    
    Private Declare PtrSafe Function DrawMenuBar Lib "user32.dll" (ByVal hWnd As LongPtr) As Long

#Else

    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, _
         ByVal dwNewLong As Long) As Long
    
    Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
    
    Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal hWnd As Long) As Long
#End If
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()

    '閉じるボタン 消去
    #If VBA7 Then
        Dim hWnd As LongPtr
        Dim Wnd_STYLE As LongPtr
    #Else
        Dim hWnd As Long
        Dim Wnd_STYLE As Long
    #End If

    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)

    #If VBA7 Then
        Call SetWindowLongPtr(hWnd, GWL_STYLE, Wnd_STYLE)
    #Else
        Call SetWindowLong(hWnd, GWL_STYLE, Wnd_STYLE)
    #End If

    DrawMenuBar hWnd

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo Err

    'フォームキャプション 設定
    With Main
        Frm_Main.Caption = .Cells(1, 1).Value & " " & .Cells(1, 19).Value
    End With

    Dim ThisWorksheets As Sheets
    Set ThisWorksheets = ThisWorkbook.Worksheets
    'シート有無 確認
    If Not HasWorksheet(ThisWorksheets, "Main", True) Then
        Call ShowErrMsg("コピーツール内に実行シートが存在しません。")
        End
    End If

    'コンボボックス 値追加
    Dim ws As Worksheet
    For Each ws In ThisWorksheets
        If ws.CodeName <> "Main" And ws.CodeName <> "Template" Then
            Cmb_ShName.AddItem ws.Name
        End If
    Next ws

    'バックグラウンドモード 初期化
    Ckb_Background.Value = False

    Exit Sub

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

    'フォーム 終了
    End

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Cmb_ShName_Change()
On Error Resume Next

    With ThisWorkbook.Sheets(Cmb_ShName.Value)
        'ファイルパス 表示
        Lbl_FPath.Caption = .Range("C13")
        Lbl_TPath.Caption = .Range("D13")
    End With

    '選択項目 保存
    g_ParamShName = Cmb_ShName.Value

    'OKボタン フォーカス
    Btn_OK.SetFocus

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Btn_ChangeFPath_Click()
On Error Resume Next

    'シート 選択判定
    Dim Path As String
    If IsBlank(g_ParamShName) Then
        Call ShowErrMsg("対象シートが選択されていません。" & vbCrLf & "シートを選択してください。")
        Exit Sub
    Else
        Path = GetPath(Lbl_FPath.Caption)
        If Not FolderExists(Path) Then Path = ThisWorkbook.Path
    End If

    'カレントドライブ・ディレクトリ 初期化
    If Left(Path, 2) = "\\" Then
        CreateObject("WScript.Shell").CurrentDirectory = Path
    Else
        Call ChDrive(Path)
        Call ChDir(Path)
    End If

    If Err.Number > 0 Then
        Path = ThisWorkbook.Path
        'カレントドライブ・ディレクトリ 初期化
        If Left(Path, 2) = "\\" Then
            CreateObject("WScript.Shell").CurrentDirectory = Path
        Else
            Call ChDrive(Path)
            Call ChDir(Path)
        End If
    End If

    '選択ファイル名 取得
    Dim FileSpec As String
    FileSpec = GetSelectFileSpec()

    'ファイル 選択判定
    If Not IsBlank(FileSpec) Then
        '選択ファイル名 反映
        Lbl_FPath.Caption = FileSpec
        ThisWorkbook.Sheets(g_ParamShName).Range("C13").Value = FileSpec
    End If

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Btn_ChangeTPath_Click()
On Error Resume Next

    'シート 選択判定
    Dim Path As String
    If IsBlank(g_ParamShName) Then
        Call ShowErrMsg("対象シートが選択されていません。" & vbCrLf & "シートを選択してください。")
        Exit Sub
    Else
        Path = GetPath(Lbl_TPath.Caption)
        If Not FolderExists(Path) Then Path = ThisWorkbook.Path
    End If

    'カレントドライブ・ディレクトリ 初期化
    If Left(Path, 2) = "\\" Then
        CreateObject("WScript.Shell").CurrentDirectory = Path
    Else
        Call ChDrive(Path)
        Call ChDir(Path)
    End If

    If Err.Number > 0 Then
        Path = ThisWorkbook.Path
        'カレントドライブ・ディレクトリ 初期化
        If Left(Path, 2) = "\\" Then
            CreateObject("WScript.Shell").CurrentDirectory = Path
        Else
            Call ChDrive(Path)
            Call ChDir(Path)
        End If
    End If

    '選択ファイル名 取得
    Dim FileSpec As String
    FileSpec = GetSelectFileSpec()

    'ファイル 選択判定
    If Not IsBlank(FileSpec) Then
        '選択ファイル名 反映
        Lbl_TPath.Caption = FileSpec
        ThisWorkbook.Sheets(g_ParamShName).Range("D13").Value = FileSpec
    End If

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Btn_OK_Click()
On Error GoTo Err

    'シート 選択判定
    If IsBlank(g_ParamShName) Then
        Call ShowErrMsg("対象シートが選択されていません。" & vbCrLf & "シートを選択してください。")
        Exit Sub
    End If

    'バックグラウンド処理 確認
    g_BackgroundFLG = Ckb_Background.Value
    If g_BackgroundFLG Then
        If ShowInfoMsg(INF_BACK_ON) = vbCancel Then Exit Sub
    Else
        If ShowInfoMsg(INF_BACK_OFF) = vbCancel Then Exit Sub
    End If


Dim t As Double
t = Timer


    'パラメータ 確認
    If Not CheckParamData() Then Exit Sub

    '使用ファイルが読み取り専用で開いている場合は、一旦クローズ
    '他プロセスで以降の処理を実行した際、ファイルが重複する為
    Dim wb As Workbook
    For Each wb In Workbooks
        With wb
            If .ReadOnly And (.FullName = g_FFileSpec Or .FullName = g_TFileSpec) Then
                Call .Close(SaveChanges:=False)
            End If
        End With
    Next

    'ファイル 開閉確認
    'コピー元
    If IsOpen(g_FFileSpec) Then
        ShowErrMsg ("コピー元 ファイルが使用中です。終了してください。")
        Exit Sub
    End If
    'コピー先
    If IsOpen(g_TFileSpec) Then
        ShowErrMsg ("コピー先 ファイルが使用中です。終了してください。")
        Exit Sub
    End If

    'コピー先ファイル 属性比較
    If CompareAttr(g_TFileSpec, vbReadOnly) Then
        ShowErrMsg ("コピー先 ファイルの属性が読み取り専用です。終了してください。")
        Exit Sub
    End If

    'コピー先ファイル バックアップ
    If g_BackupFLG Then
        If Not BackupFile(g_TFileSpec) Then Exit Sub
    End If

    'バックグラウンド処理 判定
    If g_BackgroundFLG Then

        'Excelオブジェクト 取得
        Set g_ExcelApp = New Excel.Application

        With g_ExcelApp
            .DisplayAlerts = False

            'マクロ警告ダイアログ 値保存
            Dim AutoSecurity As MsoAutomationSecurity
            AutoSecurity = .AutomationSecurity
            'マクロ警告ダイアログ 値変更
            .AutomationSecurity = msoAutomationSecurityForceDisable

            'ファイル オープン
            'コピー元
            Call .Workbooks.Open(g_FFileSpec, ReadOnly:=True, Notify:=False, Password:=g_FPswd)
            Set g_FWb = .ActiveWorkbook
            'コピー先
            Call .Workbooks.Open(g_TFileSpec, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, Password:=g_TPswd)
            Set g_TWb = .ActiveWorkbook

            'マクロ警告ダイアログ 値復元
            .AutomationSecurity = AutoSecurity

            .DisplayAlerts = True
        End With

    Else

        'Excelオブジェクト 取得
        Set g_ExcelApp = Excel.Application

        '同一ファイル名 存在確認
        'コピー元
        If HasWorkbook(g_ExcelApp.Workbooks, g_FFileName) Then
            Call ShowErrMsg("コピー元と同名のファイルが開かれています。終了して下さい。")
            Exit Sub
        End If
        'コピー先
        If HasWorkbook(g_ExcelApp.Workbooks, g_TFileName) Then
            Call ShowErrMsg("コピー先と同名のファイルが開かれています。終了して下さい。")
            Exit Sub
        End If

        With g_ExcelApp
            .DisplayAlerts = False

            'ファイル オープン
            'コピー元
            Call .Workbooks.Open(g_FFileSpec, ReadOnly:=True, Notify:=False, Password:=g_FPswd)
            Set g_FWb = .ActiveWorkbook
            'コピー先
             Call .Workbooks.Open(g_TFileSpec, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, Password:=g_TPswd)
            Set g_TWb = .ActiveWorkbook

            .DisplayAlerts = True
        End With
    End If

    Dim BeforeCalc As Long
    '各プロセス 停止
    Call BeforeProcess(BeforeCalc, g_ExcelApp)

    'シート 存在確認
    'コピー元
    If Not HasWorksheet(g_FWb.Sheets, g_FShName) Then
        Call ShowErrMsg("コピー元 ファイル内に指定されたシートが存在しません｡" & vbCrLf & "シート名: " & g_FShName)
        Exit Sub
    End If
    'コピー先
    If Not HasWorksheet(g_TWb.Sheets, g_TShName) Then
        Call ShowErrMsg("コピー先 ファイル内に指定されたシートが存在しません｡" & vbCrLf & "シート名: " & g_TShName)
        Exit Sub
    End If

    'データ コピー
    Dim CopiedCount As Long
    If Not DataCopy(CopiedCount) Then
        Call ShowErrMsg("データのコピーに失敗しました。")
        GoTo ErrAfter
    End If

    '自動保存 判定
    If g_AutoSaveFLG Then Call g_TWb.Save


Debug.Print "Btn_OK_Click:" & (Timer - t) & "秒"


    'メッセージ 表示
    Call MsgBox("データのコピーが終了しました｡" & vbCrLf & "処理件数: " & CopiedCount & "件です。", vbOKOnly)

    'コピー先ファイル 表示
    g_ExcelApp.Visible = True

    '各プロセス 再開
    Call AfterProcess(BeforeCalc, g_ExcelApp)

    Call CloseWorkbook
    Call Unload(Me)

    Exit Sub

'エラー処理
Err:

    Call ShowErrMsg(Err.Description, Err.Number)

ErrAfter:

    '各プロセス 再開
    Call AfterProcess(BeforeCalc, g_ExcelApp)

    Call CloseWorkbook(True)
    Call Unload(Me)

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub Btn_Cancel_Click()
On Error Resume Next

    '処理終了
    Unload Me

End Sub
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function CheckParamData() As Boolean
On Error GoTo Err


Dim t As Double
t = Timer


    'パラメータアドレス 設定
    Dim LogAddress As String, FormatAddress As String, BlankAddress As String, AutoSaveAddress As String, BackupAddress As String
    Dim ChangeMarkAddress As String, FontAddress As String, InteriorAddress As String
    Dim FFileSpecAddress As String, TFileSpecAddress As String, FShNameAddress As String, TShNameAddress As String
    Dim FItemIDAddress As String, TItemIDAddress As String, FStartColAddress As String, TStartColAddress As String
    Dim FDataIDAddress As String, TDataIDAddress As String, FStartRowAddress As String, TStartRowAddress As String
    Dim FPassAddress As String, TPassAddress As String
    LogAddress = "$C$3":        FormatAddress = "$C$4":     BlankAddress = "$C$5":      AutoSaveAddress = "$C$6": BackupAddress = "$C$7"
    ChangeMarkAddress = "$C$8": FontAddress = "$C$9":       InteriorAddress = "$C$10"
    FFileSpecAddress = "$C$13": TFileSpecAddress = "$D$13": FShNameAddress = "$C$14":   TShNameAddress = "$D$14"
    FItemIDAddress = "$C$15":   TItemIDAddress = "$D$15":   FStartColAddress = "$C$16": TStartColAddress = "$D$16"
    FDataIDAddress = "$C$17":   TDataIDAddress = "$D$17":   FStartRowAddress = "$C$18": TStartRowAddress = "$D$18"
    FPassAddress = "$C$19":     TPassAddress = "$D$19"

    Set g_ParamSh = ThisWorkbook.Sheets(g_ParamShName)

    'パラメータ 確認
    With g_ParamSh

        Dim ErrParam As String
        '更新ログ作成
        If Not ConvertTrueFalse(.Range(LogAddress).Value, g_LogFLG) Then ErrParam = ErrParam & vbCrLf & "更新ログ作成"
        '書式コピー
        If Not ConvertTrueFalse(.Range(FormatAddress).Value, g_FormatFLG) Then ErrParam = ErrParam & vbCrLf & "書式コピー"
        '空白コピー
        If Not ConvertTrueFalse(.Range(BlankAddress).Value, g_BlankFLG) Then ErrParam = ErrParam & vbCrLf & "空白コピー"
        'コピー先オートセーブ
        If Not ConvertTrueFalse(.Range(AutoSaveAddress).Value, g_AutoSaveFLG) Then ErrParam = ErrParam & vbCrLf & "コピー先オートセーブ"
        'コピー先バックアップ
        If Not ConvertTrueFalse(.Range(BackupAddress).Value, g_BackupFLG) Then ErrParam = ErrParam & vbCrLf & "コピー先バックアップ"

        'コピー先変更行マーク
        If Not IsBlank(.Range(ChangeMarkAddress).Value) Then
            If Not ConvertColumnReference(xlR1C1, .Range(ChangeMarkAddress).Value, g_ChangeMarkCol) Then
                ErrParam = ErrParam & vbCrLf & "コピー先変更マーク行"
            End If
        End If

        'コピーセル文字色
        g_FontColor = .Range(FontAddress).Font.Color
        'コピーセル背景色
        g_InteriorColor = .Range(InteriorAddress).Interior.Color

        'ファイルパス
        g_FFileSpec = .Range(FFileSpecAddress).Value
        g_TFileSpec = .Range(TFileSpecAddress).Value
        If IsBlank(g_FFileSpec) Then ErrParam = ErrParam & vbCrLf & "コピー元 ファイルパス"
        If IsBlank(g_TFileSpec) Then ErrParam = ErrParam & vbCrLf & "コピー先 ファイルパス"

        'ファイル名
        g_FFileName = GetFileName(g_FFileSpec)
        g_TFileName = GetFileName(g_TFileSpec)
        If IsBlank(g_FFileName) Then ErrParam = ErrParam & vbCrLf & "コピー元 ファイル名"
        If IsBlank(g_TFileName) Then ErrParam = ErrParam & vbCrLf & "コピー先 ファイル名"

        'シート名
        g_FShName = .Range(FShNameAddress).Value
        g_TShName = .Range(TShNameAddress).Value
        If IsBlank(g_FShName) Then ErrParam = ErrParam & vbCrLf & "コピー元 シート名"
        If IsBlank(g_TShName) Then ErrParam = ErrParam & vbCrLf & "コピー先 シート名"

        '項目ID行
        g_FItemIDRow = Int(Val(.Range(FItemIDAddress).Value))
        g_TItemIDRow = Int(Val(.Range(TItemIDAddress).Value))
        If g_FItemIDRow < MIN_ROW Or g_FItemIDRow > MAX_ROW Then ErrParam = ErrParam & vbCrLf & "コピー元 項目ID行"
        If g_TItemIDRow < MIN_ROW Or g_TItemIDRow > MAX_ROW Then ErrParam = ErrParam & vbCrLf & "コピー先 項目ID行"

        '項目開始列
        If Not ConvertColumnReference(ByVal xlR1C1, .Range(FStartColAddress).Value, g_FStartCol) Then
            ErrParam = ErrParam & vbCrLf & "コピー元 項目開始列"
        End If
        If Not ConvertColumnReference(ByVal xlR1C1, .Range(TStartColAddress).Value, g_TStartCol) Then
            ErrParam = ErrParam & vbCrLf & "コピー先 項目開始列"
        End If

        'データID列
        If Not ConvertColumnReference(ByVal xlR1C1, .Range(FDataIDAddress).Value, g_FDataIDCol) Then
            ErrParam = ErrParam & vbCrLf & "コピー元 データID列"
        End If
        If Not ConvertColumnReference(ByVal xlR1C1, .Range(TDataIDAddress).Value, g_TDataIDCol) Then
            ErrParam = ErrParam & vbCrLf & "コピー先 データID列"
        End If

        'データ開始行
        g_FStartRow = Int(Val(.Range(FStartRowAddress).Value))
        g_TStartRow = Int(Val(.Range(TStartRowAddress).Value))
        If g_FStartRow < MIN_ROW Or g_FStartRow > MAX_ROW Then ErrParam = ErrParam & vbCrLf & "コピー元 データ開始行"
        If g_TStartRow < MIN_ROW Or g_TStartRow > MAX_ROW Then ErrParam = ErrParam & vbCrLf & "コピー先 データ開始行"

        'パスワード
        g_FPswd = .Range("C19").Value
        g_TPswd = .Range("C19").Value

        If Not IsBlank(ErrParam) Then
            Call ShowErrMsg("以下のパラメータが不正です。設定しなおしてください。" & ErrParam)
            GoTo Err_Config
        End If

        'コピー対象項目ID 空白確認
        If .Cells(Rows.Count, 2).End(xlUp).Row < COPYCOL_START_ROW Then
            Call ShowErrMsg("項目IDが1つも設定されていません。設定してください。")
            GoTo Err_Config
        End If

        'コピー対象項目配列 取得
        '1列目：項目名
        '2列目：項目ID
        '3列目：別項目ID
        g_CopyItemArr = .Range(.Cells(COPYCOL_START_ROW, 1), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 3)).Value

    End With

    Dim CopyColIndex As Long
    For CopyColIndex = 1 To UBound(g_CopyItemArr, 1)

        Dim FCopyCol As Long
        If IsBlank(g_CopyItemArr(CopyColIndex, 2)) Then
            Call ShowErrMsg("以下の項目名の項目IDが設定されていません。設定してください。" & vbCrLf & g_CopyItemArr(CopyColIndex, 1))
            GoTo Err_Config
        End If

    Next CopyColIndex

    'ファイル 存在確認
    If Not FileExists(g_FFileSpec) Then
        Call ShowErrMsg("コピー元 ファイルが存在しません｡")
        Exit Function
    End If
    If Not FileExists(g_FFileSpec) Then
        Call ShowErrMsg("コピー先 ファイルが存在しません｡")
        Exit Function
    End If


Debug.Print "CheckParamData:" & (Timer - t) & "秒"


    CheckParamData = True

    Exit Function

'エラー処理
Err_Config:

    g_ParamSh.Activate

    Exit Function

Err:

    Call ShowErrMsg(Err.Number, Err.Description)

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Function DataCopy(ByRef CopiedCount As Long) As Boolean
On Error GoTo Err


Dim t As Double
t = Timer


    '「処理中」フォーム 表示
    Dim BarMaxWidth As Single
    With Frm_Progress
        Call .Show(vbModeless)
        BarMaxWidth = .Fra_Progress.Width
        .Lbl_Description.Caption = "コピー中……"
    End With

    'シート 取得
    Dim FSheet As Worksheet, TSheet As Worksheet
    Set FSheet = g_FWb.Sheets(g_FShName)
    Set TSheet = g_TWb.Sheets(g_TShName)

    'コピーシート 全表示
    Call ShowAllData(FSheet)
    With TSheet
        If .FilterMode Then Call .ShowAllData
    End With

    '項目ID最大列数, データID最大行数, 項目ID範囲, データID範囲 取得
    Dim FEndCol As Long, TEndCol As Long
    Dim FEndRow As Long, TEndRow As Long
    Dim FItemIDArr() As Variant, TItemIDArr() As Variant
    Dim FDataIDArr() As Variant, TDataIDArr() As Variant
    'コピー元
    With FSheet
        FEndCol = .Cells(g_FItemIDRow, .Columns.Count).End(xlToLeft).Column
        FEndRow = .Cells(.Rows.Count, g_FDataIDCol).End(xlUp).Row
        '配列 判定
        With .Range(.Cells(g_FItemIDRow, g_FStartCol), .Cells(g_FItemIDRow, FEndCol))
            If IsArray(.Value) Then
                FItemIDArr = .Value
            Else
                ReDim FItemIDArr(1 To 1, 1 To 1)
                FItemIDArr(1, 1) = .Value
            End If
        End With
        '配列 判定
        With .Range(.Cells(g_FStartRow, g_FDataIDCol), .Cells(FEndRow, g_FDataIDCol))
            If IsArray(.Value) Then
                FDataIDArr = .Value
            Else
                ReDim FDataIDArr(1 To 1, 1 To 1)
                FDataIDArr(1, 1) = .Value
            End If
        End With
    End With
    'コピー先
    With TSheet
        TEndCol = .UsedRange.Columns.Count
        TEndRow = .UsedRange.Rows.Count
        '配列 判定
        With .Range(.Cells(g_TItemIDRow, g_TStartCol), .Cells(g_TItemIDRow, TEndCol))
            If IsArray(.Value) Then
                TItemIDArr = .Value
            Else
                ReDim TItemIDArr(1 To 1, 1 To 1)
                TItemIDArr(1, 1) = .Value
            End If
        End With
        '配列 判定
        With .Range(.Cells(g_TStartRow, g_TDataIDCol), .Cells(TEndRow, g_TDataIDCol))
            If IsArray(.Value) Then
                TDataIDArr = .Value
            Else
                ReDim TDataIDArr(1 To 1, 1 To 1)
                TDataIDArr(1, 1) = .Value
            End If
        End With
    End With

    '配列2次元要素数 追加
    '4列目：コピー元項目ID各要素番号
    '5列目：コピー先項目ID各要素番号
    ReDim Preserve g_CopyItemArr(1 To UBound(g_CopyItemArr, 1), 1 To UBound(g_CopyItemArr, 2) + 2)

    'コピー対象項目配列 ループ
    Dim CopyItemIDCount As Long
    For CopyItemIDCount = 1 To UBound(g_CopyItemArr, 1)

        'コピー対象項目要素 取得
        Dim CopyItemID As Variant, AnotherCopyItemID As Variant
        CopyItemID = g_CopyItemArr(CopyItemIDCount, 2)
        AnotherCopyItemID = g_CopyItemArr(CopyItemIDCount, 3)

        Dim FItemIDCol As Variant, TItemIDCol As Variant
        Dim FItemIDCount As Long, TItemIDCount As Long
        'コピー元
        FItemIDCol = Application.Match(CopyItemID, FItemIDArr, 0)
        If Not IsError(FItemIDCol) Then
            g_CopyItemArr(CopyItemIDCount, 4) = FItemIDCol
            FItemIDCount = FItemIDCount + 1
        End If
        'コピー先
        If IsBlank(AnotherCopyItemID) Then
            TItemIDCol = Application.Match(CopyItemID, TItemIDArr, 0)
        Else
            TItemIDCol = Application.Match(AnotherCopyItemID, TItemIDArr, 0)
        End If
        If Not IsError(TItemIDCol) Then
            g_CopyItemArr(CopyItemIDCount, 5) = TItemIDCol
            TItemIDCount = TItemIDCount + 1
        End If

    Next

    If FItemIDCount * TItemIDCount = 0 Then
        If FItemIDCount = 0 Then Call ShowErrMsg("コピー元ファイルにコピー対象項目が存在しません。")
        If TItemIDCount = 0 Then Call ShowErrMsg("コピー先ファイルにコピー対象項目が存在しません。")
        Exit Function
    End If

    'コピーデータ保持用配列
    Dim CopyData() As String, CopyAllData() As Variant

    ReDim CopyAllData(0)

    'コピー元対象行 ループ
    Dim FDataRow As Long
    For FDataRow = g_FStartRow To FEndRow

        ReDim CopyData(0)

        '対象項目ID 取得
        Dim DataID As Variant
        DataID = FDataIDArr(FDataRow - (g_FStartRow - 1), 1)

        'コピー先対象行 取得(相対)
        Dim RelDataRow As Variant
        RelDataRow = Application.Match(DataID, TDataIDArr, 0)

        'コピー先対象行 有無判定
        If Not IsError(RelDataRow) Then

            'コピー先対象行 取得(絶対)
            Dim TDataRow As Long
            TDataRow = CLng(RelDataRow) + (g_TStartRow - 1)

            'コピー対象行配列 取得
            'コピー元
            Dim FDataRowArr() As Variant, TDataRowArr() As Variant, TDataRowFormulaArr() As Variant
            With FSheet
                With .Range(.Cells(FDataRow, g_FStartCol), .Cells(FDataRow, FEndCol))
                    '配列 判定
                    If IsArray(.Value) Then
                        FDataRowArr = .Value
                    Else
                        ReDim FDataRowArr(1 To 1, 1 To 1)
                        FDataRowArr(1, 1) = .Value
                    End If
                End With
            End With
            'コピー先
            With TSheet
                With .Range(.Cells(TDataRow, g_TStartCol), .Cells(TDataRow, TEndCol))
                    '配列 判定
                    If IsArray(.Value) Then
                        TDataRowArr = .Value

                        On Error GoTo Err_Formula

                        TDataRowFormulaArr = .Formula

                    Else
                        ReDim TDataRowArr(1 To 1, 1 To 1)
                        ReDim TDataRowFormulaArr(1 To 1, 1 To 1)
                        TDataRowArr(1, 1) = .Value

                        On Error GoTo Err_Formula

                        TDataRowFormulaArr(1, 1) = .Formula

                    End If
                End With

                GoTo Skip

Err_Formula:

                'Formula文字数制限(8221文字)該当セル 処理
                ReDim TDataRowFormulaArr(1 To 1, 1 To TEndCol - g_TStartCol + 1)
                Dim Cell As Range
                For Each Cell In .Range(.Cells(TDataRow, g_TStartCol), .Cells(TDataRow, TEndCol))
                    With Cell
                        If .HasFormula Then
                            TDataRowFormulaArr(1, .Column - g_TStartCol + 1) = .Formula
                        Else
                            TDataRowFormulaArr(1, .Column - g_TStartCol + 1) = .Value
                        End If
                    End With
                Next Cell

            End With

Skip:

            On Error GoTo Err

            'コピー対象項目配列 ループ
            Dim CopyItemCount As Long
            For CopyItemCount = 1 To UBound(g_CopyItemArr, 1)

                'コピー対象項目要素 取得
                Dim CopyItemFNo As Long, CopyItemTNo As Long
                CopyItemFNo = g_CopyItemArr(CopyItemCount, 4)
                CopyItemTNo = g_CopyItemArr(CopyItemCount, 5)

                'コピー対象項目 有無判定
                If CopyItemFNo * CopyItemTNo = 0 Then GoTo Continue

                'コピー対象セル値 取得
                Dim FValue As Variant, TValue As Variant
                FValue = FDataRowArr(1, CopyItemFNo)
                TValue = TDataRowArr(1, CopyItemTNo)

                'セル値 エラー判定
                'コピー元
                If IsError(FValue) Then GoTo Continue
                'コピー先
                If IsError(TValue) Then GoTo Continue

                'コピー元空白時処理 判定
                If Not (Not g_BlankFLG And IsBlank(Trim(FValue))) Then

                    'セル値 比較
                    If CStr(FValue) <> CStr(TValue) Then

                        '項目ID 配列追加
                        If IsBlank(CopyData(0)) Then CopyData(0) = DataID

                        'コピー対象セル列記号 取得
                        Dim FDataCol As String, TDataCol As String
                        If Not ConvertColumnReference(xlA1, FDataCol, CopyItemFNo + (g_FStartCol - 1)) Then
                            Exit Function
                        End If
                        If Not ConvertColumnReference(xlA1, TDataCol, CopyItemTNo + (g_TStartCol - 1)) Then
                            Exit Function
                        End If

                        If g_LogFLG Then
                            'コピーデータ 配列格納
                            '項目名[アドレス], コピー元セル値, コピー先セル値 配列追加
                            Dim AddCopyDataCount As Long
                            AddCopyDataCount = UBound(CopyData, 1) + 3
                            ReDim Preserve CopyData(AddCopyDataCount)
                            CopyData(AddCopyDataCount - 2) = g_CopyItemArr(CopyItemCount, 1) & "[" & FDataCol & FDataRow & ", " & TDataCol & TDataRow & "]"
                            CopyData(AddCopyDataCount - 1) = FValue
                            CopyData(AddCopyDataCount) = TValue
                        End If

                        '変更値 配列反映
                        TDataRowFormulaArr(1, CopyItemTNo) = FValue

                        'コピーセル 取得
                        Dim FCell As Range, TCell As Range
                        'コピー元
                        Set FCell = FSheet.Range(FDataCol & FDataRow)
                        'コピー先
                        Set TCell = TSheet.Range(TDataCol & TDataRow)

                        '書式コピーフラグ 判定
                        If g_FormatFLG Then
                            Call FCell.Copy(Destination:=TCell)
                        Else
                            '文字色 変更
                            TCell.Font.Color = g_FontColor
                            '背景色 変更
                            TCell.Interior.Color = g_InteriorColor
                        End If
                    End If
                End If

Continue:
            Next CopyItemCount

            If Not IsBlank(CopyData(0)) Then

                With TSheet
                    'コピー先対象行 反映
                    .Range(.Cells(TDataRow, g_TStartCol), .Cells(TDataRow, TEndCol)).Value = TDataRowFormulaArr
                    '処理日付 反映
                    If g_ChangeMarkCol > 0 Then .Cells(TDataRow, g_ChangeMarkCol) = Date
                End With

                'ログ用変更データ 配列格納
                If g_LogFLG And UBound(CopyData, 1) > 0 Then
                    CopyAllData(UBound(CopyAllData, 1)) = CopyData
                    ReDim Preserve CopyAllData(UBound(CopyAllData, 1) + 1)
                End If

                '処理件数 カウント
                CopiedCount = CopiedCount + 1
            End If
        End If

        'フォーム 更新
        Dim FCopyRatio As Long, TCopyRatio As Long
        TCopyRatio = (FDataRow - (g_FStartRow - 1)) * 100 \ (FEndRow - (g_FStartRow - 1))
        If TCopyRatio <> FCopyRatio Then
            Frm_Progress.Lbl_ProgressBar.Width = BarMaxWidth * TCopyRatio / 100
            Call Frm_Progress.Repaint
            FCopyRatio = TCopyRatio
        End If

        DoEvents

    Next FDataRow

    'コピー先シート アクティベート
    TSheet.Activate

    'ログ出力 判定
    If g_LogFLG And CopiedCount > 0 Then
    
        '「ログ出力中」フォーム 表示
        Frm_Progress.Lbl_Description.Caption = "ログ出力中……"

        'ログ用シート 作成
        Dim LogSh As Worksheet
        Set LogSh = ThisWorkbook.Sheets.Add

        'ログ 出力
        Dim CopyMaxCount
        Dim LogCount As Long
        For LogCount = 0 To UBound(CopyAllData, 1) - 1

            'ログ出力最大値 取得
            Dim MaxCopyData As Long
            MaxCopyData = UBound(CopyAllData(LogCount), 1)
            CopyMaxCount = WorksheetFunction.Max(CopyMaxCount, MaxCopyData)

            '2行目以降 ログ出力
            With LogSh
                .Range(.Cells(2 + LogCount, 1), .Cells(2 + LogCount, MaxCopyData)).Value = CopyAllData(LogCount)
            End With

            'フォーム 更新
            Dim FLogRatio As Long, TLogRatio As Long
            TLogRatio = LogCount * 100 \ UBound(CopyAllData, 1)
            If TLogRatio <> FLogRatio Then
                Frm_Progress.Lbl_ProgressBar.Width = BarMaxWidth * TLogRatio / 100
                Call Frm_Progress.Repaint
                FLogRatio = TLogRatio
            End If

        Next LogCount

        'ログ見出し 出力
        Dim LogTitle() As String
        ReDim Preserve LogTitle(CopyMaxCount) As String
        LogTitle(0) = "データID"
        Dim LogTitleCount
        For LogTitleCount = 1 To UBound(LogTitle, 1) Step 3
            LogTitle(LogTitleCount) = "項目名, アドレス[元, 先]"
            LogTitle(LogTitleCount + 1) = "コピー元セル値"
            LogTitle(LogTitleCount + 2) = "コピー先セル値"
        Next LogTitleCount
        With LogSh
            .Range(.Cells(1, 1), .Cells(1, 1 + UBound(LogTitle, 1))).Value = LogTitle
        End With

        'セル幅, 罫線 設定
        With LogSh.UsedRange
            Call .Columns.AutoFit
            .Borders.LineStyle = True
        End With

        'ログとして新規ブックへ
        Call LogSh.Copy
        '元シート削除
        Application.DisplayAlerts = False
        On Error Resume Next
        Call LogSh.Delete
        On Error GoTo 0
        Application.DisplayAlerts = False

    End If


Debug.Print "DataCopy:" & (Timer - t) & "秒"


    'フォーム 終了
    Call Unload(Frm_Progress)

    DataCopy = True

    Exit Function

'エラー処理
Err:

    'フォーム 終了
    Call Unload(Frm_Progress)

    Call ShowErrMsg(Err.Description, Err.Number)

End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Private Sub CloseWorkbook(Optional ByVal ErrFLG As Boolean)
On Error Resume Next

    'ファイル クローズ
    'コピー元
    If Not IsBlank(g_FWb) Then
        Call g_FWb.Close(SaveChanges:=False)
    End If
    If ErrFLG Then
        'コピー先
        If Not IsBlank(g_TWb) Then
            Call g_TWb.Close(SaveChanges:=False)
        End If

        'バックグラウンド処理 判定
        If g_BackgroundFLG Then
            If Not IsBlank(g_ExcelApp) Then Call g_ExcelApp.Quit
        End If
    End If

End Sub
'----------------------------------------------------------------------------------------------------
