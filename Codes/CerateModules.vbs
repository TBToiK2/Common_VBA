Option Explicit

Dim FSO, WSH, ADOSRead, ADOSWrite
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WSH = CreateObject("WScript.Shell")
Set ADOSRead = CreateObject("ADODB.Stream")
Set ADOSWrite = CreateObject("ADODB.Stream")

Dim REG
Set REG = CreateObject("VBScript.RegExp")
With REG
    .Global = False
    .IgnoreCase = False
    .Multiline = False
    .Pattern = "^(release|hotfix)/(.*)$"
End With

With ADOSRead
    .Mode = 3
    .Type = 2
    .Charset = "UTF-8"
    .LineSeparator = -1
End With

With ADOSWrite
    .Mode = 3
    .Type = 2
    .Charset = "Shift_JIS"
    .LineSeparator = -1
End With

Dim dateTime, fullDateTime, dateTimeFmt
dateTime = Now
fullDateTime = DateValue(dateTime) & " " & Right("0" & TimeValue(dateTime), 8)
dateTimeFmt = Replace(Replace(Replace(fullDateTime, "/", ""), ":", ""), " ", "")

With FSO

    Dim thisFilePath, driveName
    thisFilePath = .GetParentFolderName(WScript.ScriptFullName)
    driveName = .GetDriveName(thisFilePath)

    If RunCmd("rev-parse --git-dir") = "" Then
        Call MsgBox(thisFilePath & "はGitリポジトリではありません。")
        WScript.Quit
    End If

    Dim currentBranch
    currentBranch = Replace(Replace(RunCmd("rev-parse --abbrev-ref HEAD"), vbLf, ""), vbCr, "")
    Dim REGMatch
    Set REGMatch = REG.Execute(currentBranch)
    If REGMatch.Count = 0 Then
        Call MsgBox("現ブランチはreleaseまたはhotfixブランチではありません。")
        WScript.Quit
    End If

    Dim ver
    ver = REGMatch(0).SubMatches(1)
    If ver = "" Then
        Call MsgBox("リリース名が設定されていません。")
        WScript.Quit
    End If

    Dim basFolderSpec
    basFolderSpec = .BuildPath(thisFilePath, "bas")
    If Not .FolderExists(basFolderSpec) Then
        Call .CreateFolder(basFolderSpec)
    End If

    Dim saveToFolderSpec
    saveToFolderSpec = .BuildPath(basFolderSpec, ver)
    If Not .FolderExists(saveToFolderSpec) Then
        Call .CreateFolder(saveToFolderSpec)
        Dim newCreateFLG
        newCreateFLG = True
    End If

    Dim moduleFolderSpec
    moduleFolderSpec = .BuildPath(thisFilePath, "Modules")
    If Not .FolderExists(moduleFolderSpec) Then
        WScript.Quit
    End If

    Dim moduleFolder
    Set moduleFolder = .GetFolder(moduleFolderSpec)

    Dim fol, cnt
    For Each fol In moduleFolder.SubFolders

        Dim folSpec, folName
        folSpec = fol.Path
        folName = fol.Name

        If Left(folName, 4) = "Mod_" Then

            cnt = 0

            Dim saveToFileSpec
            saveToFileSpec = .BuildPath(saveToFolderSpec, folName & ".bas")
            With ADOSWrite
                Call .Open
                    Call .WriteText("Attribute VB_Name = """ & folName & """", 1)
                    Call .WriteText("'" & ver, 1)
                    Call .SaveToFile(saveToFileSpec, 2)
                Call .Close
            End With

            Call MergeTextFiles(folSpec, "Declarations")
            Call MergeTextFiles(folSpec, "Sub")
            Call MergeTextFiles(folSpec, "Function")

        End If

    Next

    If cnt = 0 Then
        If newCreateFLG Then
            Call .DeleteFolder(saveToFolderSpec)
        End If
        Call MsgBox("結合対象ファイルが存在しません。")
    End If

End With

Function RunCmd(gitCmd)

    Dim WSHEX
    Set WSHEX = WSH.Exec("cmd.exe /C " & driveName & " & git -C """ & thisFilePath & """ " & gitCmd)

    If WSHEX.Status = 2 Then
        Call MsgBox("コマンドの実行に失敗しました。")
        WScript.Quit
    End If

    Do While WSHEX.Status = 0
        Call WScript.Sleep(100)
    Loop

    RunCmd = WSHEX.StdOut.ReadAll

End Function

Sub MergeTextFiles(moduleFolderSpec, trgetFolderName)
On Error Resume Next

    Dim targetFolderSpec
    targetFolderSpec = FSO.BuildPath(moduleFolderSpec, trgetFolderName)

    If Not FSO.FolderExists(targetFolderSpec) Then Exit Sub

    Dim Reg
    Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Global = True
        .IgnoreCase = False
        .MultiLine = False
        .Pattern = "\r\n|\r|\n"
    End With

    Dim ADOR
    Set ADOR = CreateObject("ADODB.Recordset")
    With ADOR
        Call .Fields.Append("No", 20)
        Call .Fields.Append("FileName", 200, 256)
        Call .Open
            Dim f
            For Each f In FSO.GetFolder(targetFolderSpec).Files
                '隠しファイル属性 判定
                If (f.Attributes And 2) <> 2 And FSO.GetExtensionName(f.Name) = "txt" Then
                    Call .AddNew
                    If Right(f.Name, 17) = "_Declarations.txt" Then
                        .Fields(0) = 1
                    ElseIf Right(f.Name, 12) = "_Declare.txt" Then
                        .Fields(0) = 2
                    Else
                        .Fields(0) = 3
                    End If
                    .Fields(1) = f.Path
                    Call .Update
                End If
            Next

            .Sort = "[No]ASC, [FileName] ASC"

            Call .MoveFirst
            Do Until .EOF
                With ADOSRead
                    Call .Open
                        Call .LoadFromFile(ADOR.Fields(1))
                            Dim allTxt
                            allTxt = Reg.Replace(.ReadText(-1), vbCrLf)
                    Call .Close
                End With
                With ADOSWrite
                    Call .Open
                        Call .LoadFromFile(saveToFileSpec)
                            .Position = .Size
                            If cnt > 0 Then Call .WriteText(vbCrLf & vbCrLf, 0)
                            Call .WriteText(allTxt, 0)
                        Call .SaveToFile(saveToFileSpec, 2)
                    Call .Close
                End With
                cnt = cnt + 1

                Call .MoveNext
            Loop

        Call .Close
    End With

    Set ADOR = Nothing

End Sub