Sub MudarCorEMoverSeTresCaracteres()
    Dim ws As Worksheet
    Dim cell As Range

    ' Alterar para o nome da sua planilha
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Loop através de todas as células na coluna C que contêm dados
    For Each cell In ws.Range("C1:C" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
        ' Verifica se a célula contém exatamente 3 caracteres
        If Len(cell.Value) = 13 Then
            ' Move o conteúdo para a célula à direita (coluna D)
            cell.Offset(0, 1).Value = cell.Value
            ' Muda a cor de fundo da célula original para vermelho
            cell.Interior.Color = RGB(255, 0, 0)
            ' Limpa o conteúdo da célula original
            cell.Value = ""
        End If
    Next cell
End Sub
