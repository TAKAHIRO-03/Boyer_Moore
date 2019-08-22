Attribute VB_Name = "Module1"
Sub BM_SEARCH()

Dim ser As Search
Set ser = New Search
Dim cp As Integer

'第一引数にテキスト、第二引数にパターン
cp = ser.search_main("うぉぉぉ", "aaa")

End Sub
