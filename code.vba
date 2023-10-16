Sub reset()
    Dim wsMenu As Worksheet
    Dim wsCalc As Worksheet
    Dim wsContas As Worksheet
    
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsCalc = ThisWorkbook.Sheets("Calculos")
    Set wsContas = ThisWorkbook.Sheets("Contas")

    wsMenu.Range("C2").value = 0
    wsContas.Range("C2").value = 0
    wsMenu.Range("F9").value = 0
    wsMenu.Range("F10").value = 0
    wsMenu.Range("F11").value = 0
    wsMenu.Range("F12").value = 0
    wsMenu.Range("F13").value = 0

End Sub

Sub btnAdd_value_balance()
    Dim wsMenu As Worksheet
    Dim wsCalc As Worksheet
    Dim wsContas As Worksheet
    Dim balance As Long
    Dim new_value As Long

    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsCalc = ThisWorkbook.Sheets("Calculos")
    Set wsContas = ThisWorkbook.Sheets("Contas")
    
    If WorksheetFunction.CountA(wsMenu.Range("B7")) < 1 Then
        MsgBox "Preencha todas as células antes de adicionar.", vbExclamation
        Exit Sub
    End If
    
    ' Adicionar valor
    wsMenu.Range("C2").value = wsMenu.Range("C2").value + wsMenu.Range("B7").value
    wsContas.Range("C2").value = wsContas.Range("C2").value + wsMenu.Range("B7").value
    wsMenu.Range("F9").value = wsMenu.Range("F9").value + (wsMenu.Range("B7").value * wsCalc.Range("C4").value)
    wsMenu.Range("F10").value = wsMenu.Range("F10").value + (wsMenu.Range("B7").value * wsCalc.Range("C5").value)
    wsMenu.Range("F11").value = wsMenu.Range("F11").value + (wsMenu.Range("B7").value * wsCalc.Range("C6").value)
    wsMenu.Range("F12").value = wsMenu.Range("F12").value + (wsMenu.Range("B7").value * wsCalc.Range("C7").value)
    wsMenu.Range("F13").value = wsMenu.Range("F13").value + (wsMenu.Range("B7").value * wsCalc.Range("C8").value)
    
    wsMenu.Range("B7").ClearContents
    
End Sub

Sub btnAdd_spend_table()
    Dim wsMenu As Worksheet
    Dim wsContas As Worksheet
    Dim newRow As Long
    Dim decrementValue As Long
    Dim selectedCategory As Variant

    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsContas = ThisWorkbook.Sheets("Contas")

    ' Validação
    If WorksheetFunction.CountA(wsContas.Range("C6:E6")) < 3 Then
        MsgBox "Preencha todas as células antes de adicionar.", vbExclamation
        Exit Sub
    End If

    If wsContas.Range("B6").value = "" Then
        wsContas.Range("B6").Formula = Date
    End If
    
    ' Verificar se houve troca de mês
    If Month(wsContas.Range("B10")) <> "" And Month(wsContas.Range("B10")) <> Month(wsContas.Range("B6")) Then
        MsgBox "O mês da data inserida é diferente do mês do último dado da tabela.", vbExclamation
        MsgBox "Clique no botão REINICIAR para gerar o gráfico do mês passado e uma nova tabela.", vbExclamation
        Exit Sub
    End If
    
    ' Subtração do saldo total e da conta específica
    decrementValue = wsContas.Range("E6").value

    wsMenu.Range("C2").value = wsMenu.Range("C2").value - decrementValue
    wsContas.Range("C2").value = wsContas.Range("C2").value - decrementValue

    selectedCategory = wsContas.Range("D6").value
    Select Case selectedCategory
        Case "Gastos Fixos"
            wsMenu.Range("F9").value = wsMenu.Range("F9").value - decrementValue
        Case "Longo-Termo"
            wsMenu.Range("F10").value = wsMenu.Range("F10").value - decrementValue
        Case "Diversão"
            wsMenu.Range("F11").value = wsMenu.Range("F11").value - decrementValue
        Case "Educação"
            wsMenu.Range("F12").value = wsMenu.Range("F12").value - decrementValue
        Case "Investimentos"
            wsMenu.Range("F13").value = wsMenu.Range("F13").value - decrementValue
    End Select

    ' Adicionar despesa à tabela
    newRow = wsContas.Cells(wsContas.Rows.Count, "B").End(xlUp).Row + 1
    If WorksheetFunction.CountA(wsContas.Range("B" & 10 & ":E" & 10)) = 0 Then
        wsContas.Cells(10, "B").value = wsContas.Range("B6").value
        wsContas.Cells(10, "C").value = wsContas.Range("C6").value
        wsContas.Cells(10, "D").value = wsContas.Range("D6").value
        wsContas.Cells(10, "E").value = wsContas.Range("E6").value
    Else
        wsContas.Cells(newRow, "B").value = wsContas.Range("B6").value
        wsContas.Cells(newRow, "C").value = wsContas.Range("C6").value
        wsContas.Cells(newRow, "D").value = wsContas.Range("D6").value
        wsContas.Cells(newRow, "E").value = wsContas.Range("E6").value
    End If
    
    wsContas.Range("B6:E6").ClearContents
    
End Sub

Sub btnReset_table()
    Dim wsMenu As Worksheet
    Dim wsContas As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim chartRange As Range
    Dim chrt As ChartObject

    ' Defina a planilha que você deseja trabalhar
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsContas = ThisWorkbook.Sheets("Contas")

    ' Encontra a última linha com dados na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define o range da tabela
    Set chartRange = ws.Range("A1:B" & lastRow)

    ' Cria um gráfico com base nos dados da tabela
    Set chrt = ws.ChartObjects.add(Left:=100, Width:=375, Top:=75, Height:=225)
    chrt.Chart.SetSourceData Source:=chartRange
    chrt.Chart.HasTitle = True
    chrt.Chart.ChartTitle.Text = "Título do seu gráfico"
    chrt.Chart.ChartType = xlXYScatterLines ' Escolha o tipo de gráfico que deseja

    ' Exclui os dados da tabela original
    chartRange.ClearContents

    ' Deixa apenas a linha do cabeçalho e uma linha em branco
    ws.Rows("3:" & lastRow).Delete

End Sub

Sub btnDelete_item_table()
    Dim wsMenu As Worksheet
    Dim wsContas As Worksheet
    Dim lastRow As Long
    Dim decrementValue As Long
    Dim selectedCategory As Variant
    
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsContas = ThisWorkbook.Sheets("Contas")
    
    ' Pergunte ao usuário se ele deseja realmente excluir a última linha
    resposta = MsgBox("Você realmente deseja excluir a última linha?", vbYesNo + vbQuestion, "Confirmar Exclusão")
    
    If resposta = vbYes Then
        lastRow = wsContas.Cells(wsContas.Rows.Count, "B").End(xlUp).Row
        decrementValue = wsContas.Cells(lastRow, "E").value
        
        wsMenu.Range("C2").value = wsMenu.Range("C2").value + decrementValue
        wsContas.Range("C2").value = wsContas.Range("C2").value + decrementValue
        
        selectedCategory = wsContas.Cells(lastRow, "D").value
        Select Case selectedCategory
            Case "Gastos Fixos"
                wsMenu.Range("F9").value = wsMenu.Range("F9").value + decrementValue
            Case "Longo-Termo"
                wsMenu.Range("F10").value = wsMenu.Range("F10").value + decrementValue
            Case "Diversão"
                wsMenu.Range("F11").value = wsMenu.Range("F11").value + decrementValue
            Case "Educação"
                wsMenu.Range("F12").value = wsMenu.Range("F12").value + decrementValue
            Case "Investimentos"
                wsMenu.Range("F13").value = wsMenu.Range("F13").value + decrementValue
        End Select
        
        wsContas.Rows(lastRow).Delete
    End If

End Sub
