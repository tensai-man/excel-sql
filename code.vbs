Private Sub CommandButton1_Click()
Dim product_nums As Variant
Dim product_num, table_name, row_num As String
Dim i As Long

'リストを取得 リストとテーブル名と列を得る
table_name = Me.txt_filename.Text
row_num = Me.txt_num.Text
product_num = Me.txt_target.Text
'改行区切りで配列に入れる
product_nums = Split(product_num, vbCrLf)
'foreach処理
accu = ""
For i = LBound(product_nums) To UBound(product_nums)
'Chr(34) ダブル　Chr(39)　シングル
'union all select * from CSVTABLE('item.csv') table where "2" ='ringo'


If product_nums(i) <> "" Then
accu = accu + "union all select * from CSVTABLE('" & table_name & "') where " & Chr(34) & row_num & Chr(34) & "=" & Chr(39) & product_nums(i) & Chr(39) & " " & vbCrLf
End If

Next i
accu = Mid(accu, 11)
Me.txt_result.Text = accu
End Sub