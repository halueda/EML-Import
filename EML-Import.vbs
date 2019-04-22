'// eml ファイルを OUTLOOK に取り込むスクリプト
'//
'// 仕様：スクリプトを置いてあるフォルダにある .eml ファイルを対象
'//       サブフォルダ配下の .eml も対象とする
'//       OUTLOOK の「受信トレイ」のフォルダ「インポート」に取り込む
'//       取り込んだ .eml ファイルは削除する
'//       コマンドプロンプトから実行した場合は進捗状況を表示する
'//
'// 実行方法：スクリプトをダブルクリック、またはコマンドプロンプトから
'//           > cscript スクリプト名.vbs  で実行
'//


'//===================================================================
'// FileSystemObject
Const ForReading = 1    '// ファイルを読み取り専用として開きます。このファイルには書き込むことができません。
Const ForWriting = 2    '// ファイルを書き込み専用として開きます。
Const ForAppending = 8  '//ファイルを開き、ファイルの最後に追加して書き込みます。


'//===================================================================
'// オブジェクト準備
Dim FSO
Set FSO = WScript.CreateObject("Scripting.FileSystemObject")

Dim WSH
Set WSH = WScript.CreateObject("WScript.Shell")

Dim OutlookApp 
Set OutlookApp = WScript.CreateObject("Outlook.Application") 


'//===================================================================
'// OUTLOOK側インポートフォルダ設定
Const olFolderInbox = 6 
Dim fldImport 
Set fldImport = OutlookApp.Session.GetDefaultFolder(olFolderInbox) 
Set fldImport = fldImport.Folders("インポート")
'fldImport.Display 


'//===================================================================
'// ログ準備 - ログはスクリプトと同じ場所に作成
'//          - ログファイル名は スクリプト名_年月日.log
Dim oLog, fn
fn = FSO.getParentFolderName(WScript.ScriptFullName) & "\" & _
     FSO.GetBaseName(WScript.ScriptFullName) & "_" & _
     Replace(Left(Now(),10), "/", "") & ".log"
If FSO.FileExists(fn) = False then
    Set oLog = FSO.CreateTextFile(fn)
Else
    Set oLog = FSO.OpenTextFile(fn, ForAppending, True)
End If


'//===================================================================
'// 開始
log "START:" & FSO.GetFolder(".").Name
Call LoopFolder( FSO.GetFolder("."), fldImport )


'//===================================================================
'// 終了
OutlookApp.ActiveExplorer.Close 
log "インポートは終了しました。"
oLog.Close
Set oLog = Nothing


'//===================================================================
'// emlファイル取り込み（サブフォルダも対象）
Sub LoopFolder(objFolder, fldImport)
    Dim objSubFolder
    Dim objFile

    log "LOOP: " & objFolder

    '// ファイルを登録
    For Each objFile In objFolder.files
        '// 拡張子が .eml ならインポート処理
        If LCase(Right(objFile.Name,4)) = ".eml" Then
           OpenEml objFile, fldImport
        End If 
    Next

    '// フォルダがあれば再帰
    For Each objSubFolder In objFolder.SubFolders
       Dim mySubFolder
       '// fldImport の下に objSubFolder.Name というサブフォルダを作って、そこでやる
       Dim folderName
       folderName = FSO.GetFileName(objSubFolder.Name)
       On Error Resume Next 
       fldImport.Folders.Add folderName
       On Error GOTO 0
       Set mySubFolder =fldImport.Folders(folderName)
        LoopFolder objSubFolder, mySubFolder
    Next
End Sub


'//===================================================================
'// eml ファイルを開いてインポート
Sub OpenEml( emlFile, fldImport ) 
    '// エラー無視
    On Error Resume Next 
    log  "OPEN EML: " & emlFile.Name

    '// メールが開いていたら閉じる 
    While Not OutlookApp.ActiveInspector Is Nothing 
        OutlookApp.ActiveInspector.Close 
        WScript.Sleep 500 
    Wend 
    On Error GOTO 0

    '// eml ファイルを Outlook で開くコマンドを実行 
    WSH.Run "outlook /eml """ & FSO.getParentFolderName(emlFile) & _
        "\" & emlFile.Name & """" 

    '// Outlook 起動待ち 
    While OutlookApp.ActiveInspector Is Nothing 
        WScript.Sleep 500 
    Wend 

    '// メールフォルダ移動 
    OutlookApp.ActiveInspector.CurrentItem.Move fldImport 

    '// 取り込んだファイルは削除(エラーが発生していなければ）
    If Err.Number = 0 Then
        emlFile.Delete
    End If
End Sub 


'//===================================================================
'// ログ出力
Sub log(strMsg)
    '// エラー無視
    On Error Resume Next

    '// ログファイルに出力
    oLog.WriteLine(Now() & " " & strMsg)

    '// CSCRIPT なら ECHOで表示
    If LCase(Right(WScript.FullName, Len("cscript.exe"))) = "cscript.exe" Then
        WScript.Echo Now() & " " & strMsg
    End If
    On Error GOTO 0

End Sub
