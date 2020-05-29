Option Explicit 

'変数宣言
Dim wshShell    'WshShellオブジェクト
Dim fso         'FileSystemObjectオブジェクト
Dim str         'コピー用文字列バッファ
Dim i

'変数初期化
Set wshShell = WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
str = ""

If WScript.Arguments.Count < 1 Then
    Msgbox "ファイルまたはフォルダが選択されていません"
    WScript.Quit
Else
    'コマンドライン引数からファイル名を取得し1つずつバッファに格納
    For i = 0 To WScript.Arguments.Count - 1
        If i > 0 Then
            str = str & vbNewLine
        End If
        str = str & fso.GetFileName(WScript.Arguments(i))
    Next

End If

'文字列バッファの内容をクリップボードに送る
wshShell.Exec("clip").StdIn.Write str

'後処理
Set wshShell = Nothing
Set fso = Nothing
WScript.Quit