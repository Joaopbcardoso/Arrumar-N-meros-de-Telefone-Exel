Sub PadronizarTelefones()
    Dim ws As Worksheet
    Dim cell As Range
    Dim numero As String

    Set ws = ThisWorkbook.Sheets("rd-http-v8brasil-com-br-convers") ' Alterar para o nome da sua planilha

    For Each cell In ws.Range("L2:L" & ws.Cells(ws.Rows.Count, "L").End(xlUp).Row) 'Trocar as coordenadas da celula que precisa ser padronizada
        numero = cell.Value
        
        ' Remover espaços, hífens, parênteses e pontos
        numero = Replace(numero, " ", "")
        numero = Replace(numero, "-", "")
        numero = Replace(numero, "(", "")
        numero = Replace(numero, ")", "")
        numero = Replace(numero, ".", "")
        numero = Replace(numero, "+", "")
        
        ' Adicionar 55 na frente se não tiver
        If Left(numero, 2) <> "55" Then
            numero = "55" & numero
        End If
        
        ' Remover "55" se a célula estava vazia anteriormente
        If Len(cell.Value) = 0 And Len(numero) = 2 Then
            numero = ""
        End If
        
        ' Atualizar a célula com o número padronizado
        cell.Value = numero
    Next cell
End Sub