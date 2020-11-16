

Private Sub btnAdd_Click()

    Dim MaxRow As Integer
    
    MaxRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    '受注ID
    Range("A" & MaxRow).Value = txtOrdId.Value
    
    '日付
    'If txtYear.Value And txtMonth.Value And txtDay.Value Then
    Range("B" & MaxRow).Value = _
    txtYear.Value & "/" & _
    txtMonth.Value & "/" & _
    txtDay.Value
    'End If
    
    '受注元
    Range("C" & MaxRow).Value = cmbComp.Value
    
    '商品ID
    Range("D" & MaxRow).Value = txtProdId.Value
    
    '商品名
    Range("E" & MaxRow).Value = txtProdName.Value
    
    '単価
    Range("F" & MaxRow).Value = txtPrice.Value
    
    '数量
    Range("G" & MaxRow).Value = txtNum.Value
    
    '金額
    Range("H" & MaxRow).Value = txtTotal.Value
    
    
    '配送
    If chkDeliv.Value Then
        Range("I" & MaxRow).Value = "済"
    Else
         Range("I" & MaxRow).Value = ""
    End If
    
    
    '請求
    If chkInv.Value Then
        Range("J" & MaxRow).Value = "済"
    Else
         Range("J" & MaxRow).Value = ""
    End If
    
    '入金
    If chkPay.Value Then
        Range("K" & MaxRow).Value = "済"
    Else
         Range("K" & MaxRow).Value = ""
    End If
    
    'フォーカスを最初に戻す
    txtOrdId.SetFocus
    
    'フォームの初期化処理を改めて呼び出す(フォームを閉じなくても次に行ける)
    Call UserForm_Initialize
    
    
End Sub

Private Sub calcTotal()
'関数宣言
    

    '単価や数量が空白の時のエラー処理が必要（条件分岐）
    'IsNumeric()で数値かどうかを判定（数値でないなら計算をしないようにする）
    If IsNumeric(txtPrice.Value) And IsNumeric(txtNum.Value) Then
        '単価×数量の値を金額のテキストボックスに代入する
        txtTotal.Value = txtPrice.Value * txtNum.Value
    Else
         txtTotal.Value = ""
    End If

End Sub

Private Sub txtNum_Change()

     'txtPrice_Change()の内容と同様
     Call calcTotal

End Sub

Private Sub txtPrice_Change()

     'txtNum_Change()の内容と同様
     Call calcTotal

End Sub

Private Sub txtProdId_Change()
    '商品リストの最終行を取得
    Dim ProdMaxRow As Integer
    ProdMaxRow = Worksheets("リスト").Cells(Rows.Count, 3).End(xlUp).Row

    '商品IDを入力したら、自動的にそれに紐づいた商品名がテキストボックスに挿入されるようにする
    
    'try
    On Error GoTo vlookError
        txtProdName.Value = _
        WorksheetFunction.VLookup(CInt(txtProdId.Value), Worksheets("リスト").Range("C3:E" & ProdMaxRow), 2, False)
    Exit Sub
    
    'catch
vlookError:
    txtProdId.Value = ""
    
End Sub

Private Sub UserForm_Initialize()
'ユーザーフォームのイニシャライズ時の処理

    '最終行の取得
    Dim MaxRow As Integer
    'このMaxRowは最後に +1 はいらない
    MaxRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '受注IDを自動で最新番号（A列の最終行の値＋１）にしておく
    If MaxRow > 4 Then
        txtOrdId.Value = Range("A" & MaxRow).Value + 1
    End If
    
    '年月日を個別で取得して挿入
    txtYear.Value = Year(Date)
    txtMonth.Value = Month(Date)
    txtDay.Value = Day(Date)
    
    '会社名を「リスト」から取得してコンボボックスに加えていく
    
    Dim CompMaxRow As Integer
    '会社リストの最終行を取得
    CompMaxRow = Worksheets("リスト").Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Integer
    
    
    For i = 3 To CompMaxRow
        cmbComp.AddItem Worksheets("リスト").Range("A" & i).Value
    Next i

    
End Sub
