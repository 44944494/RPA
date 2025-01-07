On Error Resume Next ' エラー処理を開始

' Excelアプリケーションを作成
Dim excelApp
Set excelApp = CreateObject("Excel.Application")

' エラーが発生した場合に終了する処理
If Err.Number <> 0 Then
    WScript.Echo "Excelアプリケーションを作成できませんでした。エラーコード: " & Err.Number
    WScript.Quit
End If

' 非表示でExcelを開く
excelApp.Visible = False

' ブックを開く
Dim workbookPath
workbookPath = "C:\Users\diabl\Downloads\【完了】新規登録申請書 大量用Ver3.2_csvファイル出力用.xlsm"

Dim workbook
Set workbook = excelApp.Workbooks.Open(workbookPath)

' エラーが発生した場合に終了する処理
If Err.Number <> 0 Then
    WScript.Echo "ブックを開けませんでした。パスを確認してください。エラーコード: " & Err.Number
    excelApp.Quit
    Set excelApp = Nothing
    WScript.Quit
End If

' マクロを実行（標準モジュールにある場合）
On Error Resume Next
excelApp.Run "CopyMappedColumns"

If Err.Number <> 0 Then
    WScript.Echo "マクロを実行できませんでした。エラーコード: " & Err.Number
Else
    WScript.Echo "マクロの実行が完了しました。"
End If

' 必要に応じてブックを保存して閉じる
workbook.Close False

' Excelアプリケーションを終了
excelApp.Quit

' COMオブジェクトを解放
Set workbook = Nothing
Set excelApp = Nothing

WScript.Echo "処理が終了しました。"
