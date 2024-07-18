Sub ManterApenasPrimeiroNumero()
    Dim ws As Worksheet
    Dim cell As Range
    Dim cellContent As String
    Dim delimiters As Variant
    Dim firstNumber As String
    
    ' Alterar para o nome da sua planilha
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Definir delimitadores (espaço e vírgula)
    delimiters = Array(" ", ",")

    ' Loop através de todas as células na coluna C que contêm dados
    For Each cell In ws.Range("C1:C" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
        cellContent = cell.Value
        ' Separa o conteúdo da célula usando os delimitadores
        For Each delimiter In delimiters
            If InStr(cellContent, delimiter) > 0 Then
                cellContent = Split(cellContent, delimiter)(0)
                Exit For
            End If
        Next delimiter
        ' Mantém apenas o primeiro número encontrado
        cell.Value = cellContent
    Next cell
End Sub
