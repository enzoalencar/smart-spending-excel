Sub btnAdd_value_balance()
    Dim wsMenu As Worksheet
    Dim wsDespesas As Worksheet
    Dim wsContas As Worksheet
    Dim balance As Long
    Dim new_value As Long

    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsDespesas = ThisWorkbook.Sheets("Despesas")
    Set wsContas = ThisWorkbook.Sheets("Contas")
    
    If WorksheetFunction.CountA(wsMenu.Range("B7")) < 1 Then
        MsgBox "Preencha todas as células antes de adicionar.", vbExclamation
        Exit Sub
    End If
    
    ' Adicionar valor
    wsMenu.Range("C2").value = wsMenu.Range("C2").value + wsMenu.Range("B7").value
    wsDespesas.Range("C2").value = wsDespesas.Range("C2").value + wsMenu.Range("B7").value
    wsMenu.Range("F9").value = wsMenu.Range("F9").value + (wsMenu.Range("B7").value * wsContas.Range("C12").value)
    wsMenu.Range("F10").value = wsMenu.Range("F10").value + (wsMenu.Range("B7").value * wsContas.Range("C13").value)
    wsMenu.Range("F11").value = wsMenu.Range("F11").value + (wsMenu.Range("B7").value * wsContas.Range("C14").value)
    wsMenu.Range("F12").value = wsMenu.Range("F12").value + (wsMenu.Range("B7").value * wsContas.Range("C15").value)
    wsMenu.Range("F13").value = wsMenu.Range("F13").value + (wsMenu.Range("B7").value * wsContas.Range("C16").value)
    wsMenu.Range("F14").value = wsMenu.Range("F14").value + (wsMenu.Range("B7").value * wsContas.Range("C17").value)
    
    wsMenu.Range("B7").ClearContents
    
End Sub

Sub btnAdd_spend_table()
    Dim wsMenu As Worksheet
    Dim wsDespesas As Worksheet
    Dim wsContas As Worksheet
    Dim newRow As Long
    Dim decrementValue As Long
    Dim selectedCategory As Variant

    Set wsMenu = ThisWorkbook.Sheets("Menu")
    Set wsDespesas = ThisWorkbook.Sheets("Despesas")
    Set wsContas = ThisWorkbook.Sheets("Contas")

    ' Validação
    If WorksheetFunction.CountA(wsDespesas.Range("C6:E6")) < 3 Then
        MsgBox "Preencha todas as células antes de adicionar.", vbExclamation
        Exit Sub
    End If

    If wsDespesas.Range("B6").value = "" Then
        wsDespesas.Range("B6").Formula = Date
    End If
    
    ' Verificar se houve troca de mês
    If wsDespesas.Range("B10").value <> "" And Month(wsDespesas.Range("B10")) <> Month(wsDespesas.Range("B6")) Then
        MsgBox "O mês da data inserida é diferente do mês do último dado da tabela.", vbExclamation
        MsgBox "Clique no botão REINICIAR para gerar o gráfico do mês passado e uma nova tabela.", vbExclamation
        Exit Sub
    End If
    
    ' Subtração do saldo total e da conta específica
    decrementValue = wsDespesas.Range("E6").value

    wsMenu.Range("C2").value = wsMenu.Range("C2").value - decrementValue
    wsDespesas.Range("C2").value = wsDespesas.Range("C2").value - decrementValue

    selectedCategory = wsDespesas.Range("D6").value
    Select Case selectedCategory
        Case "Gastos Fixos"
            wsMenu.Range("F9").value = wsMenu.Range("F9").value - decrementValue
            wsContas.Range("F12").value = wsContas.Range("F12").value + decrementValue
        Case "Longo-Termo"
            wsMenu.Range("F10").value = wsMenu.Range("F10").value - decrementValue
            wsContas.Range("F13").value = wsContas.Range("F13").value + decrementValue
        Case "Diversão"
            wsMenu.Range("F11").value = wsMenu.Range("F11").value - decrementValue
            wsContas.Range("F14").value = wsContas.Range("F14").value + decrementValue
        Case "Educação"
            wsMenu.Range("F12").value = wsMenu.Range("F12").value - decrementValue
            wsContas.Range("F15").value = wsContas.Range("F15").value + decrementValue
        Case "Investimentos"
            wsMenu.Range("F13").value = wsMenu.Range("F13").value - decrementValue
            wsContas.Range("F16").value = wsContas.Range("F16").value + decrementValue
        Case "Doação"
            wsMenu.Range("F13").value = wsMenu.Range("F13").value - decrementValue
            wsContas.Range("F17").value = wsContas.Range("F17").value + decrementValue
    End Select

    ' Adicionar despesa à tabela
    newRow = wsDespesas.Cells(wsDespesas.Rows.Count, "B").End(xlUp).Row + 1
    If WorksheetFunction.CountA(wsDespesas.Range("B" & 10 & ":E" & 10)) = 0 Then
        wsDespesas.Cells(10, "B").value = wsDespesas.Range("B6").value
        wsDespesas.Cells(10, "C").value = wsDespesas.Range("C6").value
        wsDespesas.Cells(10, "D").value = wsDespesas.Range("D6").value
        wsDespesas.Cells(10, "E").value = wsDespesas.Range("E6").value
    Else
        wsDespesas.Cells(newRow, "B").value = wsDespesas.Range("B6").value
        wsDespesas.Cells(newRow, "C").value = wsDespesas.Range("C6").value
        wsDespesas.Cells(newRow, "D").value = wsDespesas.Range("D6").value
        wsDespesas.Cells(newRow, "E").value = wsDespesas.Range("E6").value
    End If
    
    wsDespesas.Range("B6:E6").ClearContents
    
End Sub

Sub btnReset_table()
    Dim wsMenu As Worksheet
    Dim wsDespesas As Worksheet
    Dim wsContas As Worksheet
    Dim spend_tbl As ListObject
    Dim newRange As Range
    Dim response As VbMsgBoxResult

    ' Pergunte ao usuário se ele deseja realmente excluir a última linha
    response = MsgBox("Você realmente deseja reiniciar a tabela?", vbYesNo + vbQuestion, "Confirmar Exclusão")
    
    If response = vbYes Then
        Set wsMenu = ThisWorkbook.Sheets("Menu")
        Set wsDespesas = ThisWorkbook.Sheets("Despesas")
        Set wsContas = ThisWorkbook.Sheets("Contas")
        Set spend_tbl = wsDespesas.ListObjects("main_tbl")
        
        ' Verificar se a tabela tem mais de uma linha (excluindo o cabeçalho)
        If spend_tbl.ListRows.Count > 0 Then
            spend_tbl.DataBodyRange.ClearContents
            Set newRange = spend_tbl.HeaderRowRange.Resize(2)
            spend_tbl.Resize newRange
            
            wsContas.Range("F12").value = ""
            wsContas.Range("F13").value = ""
            wsContas.Range("F14").value = ""
            wsContas.Range("F15").value = ""
            wsContas.Range("F16").value = ""
            wsContas.Range("F17").value = ""
        Else
            MsgBox "Verifique se a tabela realmente possui dados para serem limpos.", vbCritical, "Erro"
        End If
    End If
End Sub

Sub btnDelete_item_table()
    Dim wsMenu As Worksheet
    Dim wsDespesas As Worksheet
    Dim wsContas As Worksheet
    Dim lastRow As Long
    Dim decrementValue As Long
    Dim selectedCategory As Variant
    Dim response As VbMsgBoxResult
    
    ' Pergunte ao usuário se ele deseja realmente excluir a última linha
    response = MsgBox("Você realmente deseja excluir a última linha?", vbYesNo + vbQuestion, "Confirmar Exclusão")
    
    If response = vbYes Then
        Set wsMenu = ThisWorkbook.Sheets("Menu")
        Set wsDespesas = ThisWorkbook.Sheets("Despesas")
        Set wsContas = ThisWorkbook.Sheets("wsContas")
        Set spend_tbl = wsDespesas.ListObjects("main_tbl")
        
        If spend_tbl.ListRows.Count > 1 Then
            lastRow = wsDespesas.Cells(wsDespesas.Rows.Count, "B").End(xlUp).Row
            decrementValue = wsDespesas.Cells(lastRow, "E").value
            
            wsMenu.Range("C2").value = wsMenu.Range("C2").value + decrementValue
            wsDespesas.Range("C2").value = wsDespesas.Range("C2").value + decrementValue
            
            selectedCategory = wsDespesas.Cells(lastRow, "D").value
            Select Case selectedCategory
                Case "Gastos Fixos"
                    wsMenu.Range("F9").value = wsMenu.Range("F9").value + decrementValue
                    wsContas.Range("F12").value = wsContas.Range("F12").value - decrementValue
                Case "Longo-Termo"
                    wsMenu.Range("F10").value = wsMenu.Range("F10").value + decrementValue
                    wsContas.Range("F13").value = wsContas.Range("F13").value - decrementValue
                Case "Diversão"
                    wsMenu.Range("F11").value = wsMenu.Range("F11").value + decrementValue
                    wsContas.Range("F14").value = wsContas.Range("F14").value - decrementValue
                Case "Educação"
                    wsMenu.Range("F12").value = wsMenu.Range("F12").value + decrementValue
                    wsContas.Range("F15").value = wsContas.Range("F15").value - decrementValue
                Case "Investimentos"
                    wsMenu.Range("F13").value = wsMenu.Range("F13").value + decrementValue
                    wsContas.Range("F16").value = wsContas.Range("F16").value - decrementValue
                Case "Doação"
                    wsMenu.Range("F13").value = wsMenu.Range("F13").value + decrementValue
                    wsContas.Range("F17").value = wsContas.Range("F17").value - decrementValue
            End Select
            
            wsDespesas.Rows(lastRow).Delete
        Else
            MsgBox "Não foi possível deletar a linha, já está no mínimo possível.", vbCritical
            MsgBox "Se ainda assim quiser limpar as informações, tente reiniciar a tabela.", vbInformation
        End If
    End If
End Sub
