Attribute VB_Name = "Module1"
Sub BM_SEARCH()

Dim ser As search
Set ser = New search
Dim cp As Integer

'第一引数にテキスト、第二引数にパターン
cp = ser.search_main("abbbbaaaassss", "aaa")

End Sub
