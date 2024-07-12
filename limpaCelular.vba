Sub MudarCorEMoverSeDozeCaracteres()
    Dim ws As Worksheet
    Dim cell As Range

    ' Alterar para o nome da sua planilha
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Loop através de todas as células na coluna D que contêm dados
    For Each cell In ws.Range("D1:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)
        ' Verifica se a célula contém exatamente 12 caracteres
        If Len(cell.Value) = 12 Then
            ' Move o conteúdo para a célula à esquerda (coluna C)
            cell.Offset(0, -1).Value = cell.Value
            ' Muda a cor de fundo da célula original para vermelho
            cell.Interior.Color = RGB(255, 0, 0)
            ' Limpa o conteúdo da célula original
            cell.Value = ""
        End If
    Next cell
End Sub
