Module ModPagosFinagil
    Dim taPagfinagil As New PagosFinagilDSTableAdapters.PagosFinagilTableAdapter
    Dim taPagfira As New PagosFinagilDSTableAdapters.CONT_CPF_PagosFiraTableAdapter
    Dim DS As New PagosFinagilDS
    Dim rFinagil As PagosFinagilDS.PagosFinagilRow
    Dim rFira As PagosFinagilDS.CONT_CPF_PagosFiraRow

    Public Sub ExportaPagosFinagilFira()
        Dim idCredito As Decimal = 0
        taPagfinagil.Fill(DS.PagosFinagil)
        Console.WriteLine("Procesa Pagos Finagil")
        For Each rFinagil In DS.PagosFinagil.Rows
            Console.WriteLine(rFinagil.Anexo)
            If idCredito = 0 Then
                rFira = DS.CONT_CPF_PagosFira.NewCONT_CPF_PagosFiraRow
                IniciaFila(rFira)
            ElseIf idCredito <> rFinagil.idCredito Then
                If rFira.Capital <> 0 And rFira.Interes <> 0 Then 'solo si tiene algo que pagar
                    DS.CONT_CPF_PagosFira.AddCONT_CPF_PagosFiraRow(rFira)
                    DS.CONT_CPF_PagosFira.GetChanges()
                    taPagfira.Update(DS.CONT_CPF_PagosFira)
                End If

                rFira = DS.CONT_CPF_PagosFira.NewCONT_CPF_PagosFiraRow
                IniciaFila(rFira)
            End If
            If rFinagil.Tipar = "H" Or rFinagil.Tipar = "C" Then
                If InStr(rFinagil.Concepto, "INTERES") Then
                    rFira.Interes += rFinagil.Importe
                Else
                    rFira.Capital += rFinagil.Importe 'Se incluye Fega'
                End If
            Else
                If InStr(rFinagil.Concepto, "CAPITAL EQUIPO") Or InStr(rFinagil.Concepto, "SALDO INSOLUTO DEL EQUIPO") Then
                    rFira.Capital += rFinagil.Importe
                ElseIf InStr(rFinagil.Concepto, "INTERESES") Then
                    rFira.Interes += rFinagil.Importe
                End If
            End If
            idCredito = rFinagil.idCredito
            taPagfinagil.ProcesaHistoria(True, rFinagil.Anexo, rFinagil.Fecha, rFinagil.Concepto.Trim)
        Next
        If DS.PagosFinagil.Rows.Count > 0 Then
            DS.CONT_CPF_PagosFira.GetChanges()
            taPagfira.Update(DS.CONT_CPF_PagosFira)
        End If
    End Sub

    Sub IniciaFila(ByRef rFira As PagosFinagilDS.CONT_CPF_PagosFiraRow)
        rFira.Capital = 0
        rFira.Interes = 0
        rFira.FechaPagoFira = "01/01/1900"
        rFira.Procesado = False
        rFira.id_Contrato = rFinagil.id_contrato
        rFira.id_credito = rFinagil.idCredito
        rFira.FechaHistoria = rFinagil.Fecha
    End Sub
End Module
