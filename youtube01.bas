Attribute VB_Name = "youtube01"
'/**
'*
'* Youtube => Yarn Channel ExcelVBAへの挑戦状
'*
'* チャンネル登録はこちらから
'* https://www.youtube.com/channel/UCLH9TzszRQZcr9B1a7L_EvQ/?sub_confirmation=1
'*
'**/

Sub 個数表作成()
    
'すでに作成済みの場合作成不可
    If Cells(1, 1) <> "" Then
        MsgBox "すでに票が存在しています。", vbOKOnly
        Exit Sub
    End If
        
'Rnd() {ランダムに 0以上 1未満を出力}

'(Int(Rnd() * 5) + 1) [ランダムに 1 ～ 5 を出力]

'商品リスト作成
    Dim ProductList(5) As String
    
    ProductList(1) = "りんご"
    ProductList(2) = "みかん"
    ProductList(3) = "イチゴ"
    ProductList(4) = "バナナ"
    ProductList(5) = "パイナップル"

'表カテゴリ作成
    Cells(1, 1) = "管理番号"
    Cells(1, 2) = "商品"
    Cells(1, 3) = "個数"

'ランダムな表を作成
    Dim i As Long
    
    For i = 2 To 101
    
'Forループの現在数値をIDにする。
        Cells(i, 1) = i - 1
        
'上の表 [StoreList] からランダムに表示
        Cells(i, 2) = ProductList((Int(Rnd() * 5) + 1))
        
'個数をランダムに表示
        Cells(i, 3) = (Int(Rnd() * 5) + 1)

'i が 100になるまでループ
    Next i
    
'列幅の自動調整
    Columns("A").EntireColumn.AutoFit
    Columns("B").EntireColumn.AutoFit
    Columns("C").EntireColumn.AutoFit
    
'リストのテーブル化
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$101"), , xlYes).Name = "テーブル1"
    ActiveSheet.ListObjects("テーブル1").TableStyle = "TableStyleLight9"

End Sub

Sub 単価表作成()
    
    If Cells(1, 5) <> "" Then
        MsgBox "すでに票が存在しています。", vbOKOnly
        Exit Sub
    End If
    
'商品リスト作成
    Dim ProductList(5) As String
    
    ProductList(1) = "りんご"
    ProductList(2) = "みかん"
    ProductList(3) = "イチゴ"
    ProductList(4) = "バナナ"
    ProductList(5) = "パイナップル"
    
'価格リスト
    Dim PriceList(5) As Long
    
    PriceList(1) = 150
    PriceList(2) = 80
    PriceList(3) = 600
    PriceList(4) = 180
    PriceList(5) = 1000
    

'表カテゴリ作成
    Cells(1, 5) = "管理番号"
    Cells(1, 6) = "商品"
    Cells(1, 7) = "単価"
    Cells(1, 8) = "売上個数"
    Cells(1, 9) = "売上"
    
'表を作成
    For i = 2 To 6

'Forループの現在数値をIDにする。
        Cells(i, 5) = i - 1
        
'上の表 [ProductList] から表示
        Cells(i, 6) = ProductList(i - 1)
        
'上の表 [PriceList] から表示
        Cells(i, 7) = PriceList(i - 1)
        
'売上個数取得関数を入力
        Cells(i, 8) = "=集計(""" & ProductList(i - 1) & """,テーブル1[個数])"

'a売上計算関数を入力
        Cells(i, 9) = "=" & Cells(i, 8) * PriceList(i - 1)
        
'i が 100になるまでループ
    Next i
    
'列幅の自動調整
    Columns("E").EntireColumn.AutoFit
    Columns("F").EntireColumn.AutoFit
    Columns("G").EntireColumn.AutoFit
    Columns("H").EntireColumn.AutoFit
        
'リストのテーブル化
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$E$1:$I$6"), , xlYes).Name = "テーブル2"
    ActiveSheet.ListObjects("テーブル2").TableStyle = "TableStyleLight9"
    
End Sub

Function 集計(ProductName As String, RangeAria As Range) As Long
    
    Dim NumCount As Long
    NumCount = 2
    
    For Each r In RangeAria
        
        If ProductName = Cells(NumCount, 2) Then
        
            sumPrice = sumPrice + r
        
        End If
        
        NumCount = NumCount + 1
        
    Next r
    
    集計 = sumPrice
    
End Function

