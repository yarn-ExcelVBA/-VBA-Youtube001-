Attribute VB_Name = "youtube01"
'/**
'*
'* Youtube => Yarn Channel ExcelVBA�ւ̒����
'*
'* �`�����l���o�^�͂����炩��
'* https://www.youtube.com/channel/UCLH9TzszRQZcr9B1a7L_EvQ/?sub_confirmation=1
'*
'**/

Sub ���\�쐬()
    
'���łɍ쐬�ς݂̏ꍇ�쐬�s��
    If Cells(1, 1) <> "" Then
        MsgBox "���łɕ[�����݂��Ă��܂��B", vbOKOnly
        Exit Sub
    End If
        
'Rnd() {�����_���� 0�ȏ� 1�������o��}

'(Int(Rnd() * 5) + 1) [�����_���� 1 �` 5 ���o��]

'���i���X�g�쐬
    Dim ProductList(5) As String
    
    ProductList(1) = "���"
    ProductList(2) = "�݂���"
    ProductList(3) = "�C�`�S"
    ProductList(4) = "�o�i�i"
    ProductList(5) = "�p�C�i�b�v��"

'�\�J�e�S���쐬
    Cells(1, 1) = "�Ǘ��ԍ�"
    Cells(1, 2) = "���i"
    Cells(1, 3) = "��"

'�����_���ȕ\���쐬
    Dim i As Long
    
    For i = 2 To 101
    
'For���[�v�̌��ݐ��l��ID�ɂ���B
        Cells(i, 1) = i - 1
        
'��̕\ [StoreList] ���烉���_���ɕ\��
        Cells(i, 2) = ProductList((Int(Rnd() * 5) + 1))
        
'���������_���ɕ\��
        Cells(i, 3) = (Int(Rnd() * 5) + 1)

'i �� 100�ɂȂ�܂Ń��[�v
    Next i
    
'�񕝂̎�������
    Columns("A").EntireColumn.AutoFit
    Columns("B").EntireColumn.AutoFit
    Columns("C").EntireColumn.AutoFit
    
'���X�g�̃e�[�u����
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$101"), , xlYes).Name = "�e�[�u��1"
    ActiveSheet.ListObjects("�e�[�u��1").TableStyle = "TableStyleLight9"

End Sub

Sub �P���\�쐬()
    
    If Cells(1, 5) <> "" Then
        MsgBox "���łɕ[�����݂��Ă��܂��B", vbOKOnly
        Exit Sub
    End If
    
'���i���X�g�쐬
    Dim ProductList(5) As String
    
    ProductList(1) = "���"
    ProductList(2) = "�݂���"
    ProductList(3) = "�C�`�S"
    ProductList(4) = "�o�i�i"
    ProductList(5) = "�p�C�i�b�v��"
    
'���i���X�g
    Dim PriceList(5) As Long
    
    PriceList(1) = 150
    PriceList(2) = 80
    PriceList(3) = 600
    PriceList(4) = 180
    PriceList(5) = 1000
    

'�\�J�e�S���쐬
    Cells(1, 5) = "�Ǘ��ԍ�"
    Cells(1, 6) = "���i"
    Cells(1, 7) = "�P��"
    Cells(1, 8) = "�����"
    Cells(1, 9) = "����"
    
'�\���쐬
    For i = 2 To 6

'For���[�v�̌��ݐ��l��ID�ɂ���B
        Cells(i, 5) = i - 1
        
'��̕\ [ProductList] ����\��
        Cells(i, 6) = ProductList(i - 1)
        
'��̕\ [PriceList] ����\��
        Cells(i, 7) = PriceList(i - 1)
        
'������擾�֐������
        Cells(i, 8) = "=�W�v(""" & ProductList(i - 1) & """,�e�[�u��1[��])"

'a����v�Z�֐������
        Cells(i, 9) = "=" & Cells(i, 8) * PriceList(i - 1)
        
'i �� 100�ɂȂ�܂Ń��[�v
    Next i
    
'�񕝂̎�������
    Columns("E").EntireColumn.AutoFit
    Columns("F").EntireColumn.AutoFit
    Columns("G").EntireColumn.AutoFit
    Columns("H").EntireColumn.AutoFit
        
'���X�g�̃e�[�u����
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$E$1:$I$6"), , xlYes).Name = "�e�[�u��2"
    ActiveSheet.ListObjects("�e�[�u��2").TableStyle = "TableStyleLight9"
    
End Sub

Function �W�v(ProductName As String, RangeAria As Range) As Long
    
    Dim NumCount As Long
    NumCount = 2
    
    For Each r In RangeAria
        
        If ProductName = Cells(NumCount, 2) Then
        
            sumPrice = sumPrice + r
        
        End If
        
        NumCount = NumCount + 1
        
    Next r
    
    �W�v = sumPrice
    
End Function

