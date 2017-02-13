Sub cadastrar()
    Range("a2").Value = "Microondas"
    Range("b2").Value = 2
    Range("c2").Value = 450
    Range("d2") = Range("b2") * Range("c2")
    
End Sub

Sub cadastrar2()
    Range("a2").Value = InputBox("Digite o produto")
    Range("b2").Value = InputBox("Digite a quantidade")
    Range("c2").Value = InputBox("Digite o valor unitário")
    Range("d2") = Range("b2") * Range("c2")
    
End Sub

Sub cadastrar3()
    Range("a1048576").Select
    ActiveCell.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell = InputBox("Digite o produto")
    ActiveCell.Offset(0, 1) = InputBox("Digite a quantidade")
    ActiveCell.Offset(0, 2) = InputBox("Digite o valor unitário")
    ActiveCell.Offset(0, 3) = ActiveCell.Offset(0, 1) * ActiveCell.Offset(0, 2)
    Range("a1").Select
    Range(ActiveCell, ActiveCell.End(xlDown).End(xlToRight)).Borders.ColorIndex = 1
            
End Sub

Sub formatar()
    Range("a1:d5").Borders.ColorIndex
    Range("a2:d5").Font.Italic = True
    Range("a2:d5").Font.Name = "Arial"
End Sub

