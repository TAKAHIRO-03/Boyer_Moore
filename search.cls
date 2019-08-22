VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '?ｿｽ?ｿｽﾎ変撰ｿｽ?ｿｽ骭ｾ?ｿｽ?ｿｽ?ｿｽﾄね！?ｿｽ?ｿｽ?ｿｽﾄ意厄ｿｽ
Private Const BM_TABLE_SIZE = 98305
'?ｿｽA?ｿｽX?ｿｽL?ｿｽ[?ｿｽR?ｿｽ[?ｿｽh?ｿｽﾍ茨ｿｽ:-32768 ?ｿｽ` 65536
'?ｿｽA?ｿｽX?ｿｽL?ｿｽ[?ｿｽR?ｿｽ[?ｿｽh?ｿｽﾍ囲（?ｿｽﾏ奇ｿｽ?ｿｽ?ｿｽj0 ?ｿｽ` 98305
Private text() As String
Private pattern() As String
Private cp As Integer
Private Sub Class_Initialize() ' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ

    Debug.Print "Serach?ｿｽJ?ｿｽn"
    
End Sub
Private Sub Class_Terminate() ' ?ｿｽI?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ

    Erase text
    Erase pattern
    cp = 0
    Debug.Print "Serach?ｿｽJ?ｿｽn?ｿｽI?ｿｽ?ｿｽ"
    
End Sub
Public Function search_main(ByVal txt As String, ByVal patt As String) As Integer

    text = split(txt)
    pattern = split(patt)
    cp = BM_SEARCH(text, pattern) '?ｿｽ?ｿｽ?ｿｽ?ｿｽ

    If cp = 0 Then
        Debug.Print "Not found"
    Else
        Debug.Print "found!!!! " + CStr(cp) + "?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾚゑｿｽ?ｿｽ?ｿｽ"; CStr(cp + UBound(pattern)) + "?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾚの間にゑｿｽ?ｿｽ?ｿｽﾜゑｿｽ?ｿｽB"
    End If

    search_main = cp
    
End Function
Private Function split(ByVal target As String) As String()

    Dim txt As String
    Dim arr() As String
    Dim i As Long
    
    If target = "" Then
       MsgBox "?ｿｽe?ｿｽL?ｿｽX?ｿｽg?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾍゑｿｽ?ｿｽ?ｿｽﾄゑｿｽ?ｿｽﾜゑｿｽ?ｿｽ?ｿｽB"
       End
    Else
       txt = target
    End If
    
    ReDim arr(Len(txt) - 1)
    For i = 0 To UBound(arr)
        arr(i) = Mid(txt, i + 1, 1)
    Next i
    
    split = arr
    
End Function
Private Function bm_table_init(ByRef table() As Long, ByRef pattern() As String, ByRef ptn_len As Integer)

    Dim ptn_sp() As String
    Dim i As Long
    ptn_len = UBound(pattern) + 1 '?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾌ抵ｿｽ?ｿｽ?ｿｽ
    
    
    ReDim table(BM_TABLE_SIZE)
    
    For i = 0 To BM_TABLE_SIZE
        table(i) = ptn_len
    Next i
    
    For i = 0 To ptn_len - 1
        table(con_num(pattern(i))) = ptn_len - i - 1
    Next i
    
    '?ｿｽf?ｿｽo?ｿｽb?ｿｽO?ｿｽp
    Debug.Print "?ｿｽf?ｿｽt?ｿｽH?ｿｽ?ｿｽ?ｿｽg?ｿｽ?ｿｽSkip?ｿｽ?ｿｽ" + CStr(ptn_len)
    
    For i = 0 To BM_TABLE_SIZE
        If Not table(i) = ptn_len Then
            Debug.Print "Char = " + CStr(con_str(i))
            Debug.Print "table = " + CStr(i)
            Debug.Print "step = " + CStr(table(i))
        End If
    Next i

End Function
Private Function BM_SEARCH(ByRef text() As String, ByRef pattern() As String) As Integer

    Dim table() As Long
    Dim pat As String
    Dim txt As String
    Dim txt_len As Integer
    Dim ptn_len As Integer
    Dim i As Integer '?ｿｽe?ｿｽL?ｿｽX?ｿｽg?ｿｽﾌ費ｿｽr?ｿｽﾊ置
    Dim j As Integer '?ｿｽp?ｿｽ^?ｿｽ[?ｿｽ?ｿｽ?ｿｽﾌ費ｿｽr?ｿｽﾊ置
    
    txt_len = UBound(text) + 1
    ptn_len = UBound(pattern) + 1
    ReDim table(BM_TABLE_SIZE)
    
    Call bm_table_init(table, pattern, ptn_len)
    
    i = ptn_len - 1
    j = i
    
    Do While ((i < txt_len) And (j >= 0))
    
    If Not text(i) = pattern(j) Then
        i = i + table(con_num(text(i)))
        j = ptn_len - 1
    Else
        i = i - 1
        j = j - 1
    End If
    
    Loop
    
    If j < 0 Then
        BM_SEARCH = i + 2
    Else
        BM_SEARCH = 0
    End If

End Function
Private Function con_num(ByVal txt As String) As Long

    Dim number As Long
    
    If Asc(txt) < 0 Then
        number = ((Asc(txt)) * -1) + 65536
        con_num = number
    Else
        number = Asc(txt)
        con_num = number
    End If

End Function
Private Function con_str(ByVal num As Long) As String
    Dim str As String
    
    If 65536 < num Then
        str = Chr((num - 65536) * -1)
        con_str = str
    Else
        str = Chr(num)
        con_str = str
    End If

End Function
