Module ModPagosFira
    Dim DS As New PagosFiraDS

    Function Procesa_Pagos_Fira(ID As Integer)
        Dim X As Integer = 0
        If ID = 0 Then
            taPagosFira.Fill(DS.CONT_CPF_PagosFira)
        Else
            taPagosFira.Fill_ID(DS.CONT_CPF_PagosFira, ID)
        End If
        For Each r As PagosFiraDS.CONT_CPF_PagosFiraRow In DS.CONT_CPF_PagosFira.Rows
            X = 1
            GeneraPago(r.id_Contrato, r.FechaPagoFira, r.Capital, r.Interes, r.FechaHistoria)
            If ID = 0 Then
                ProcesaEstadoCuenta(r.id_Contrato, True, Date.Now.Date)
            End If
        Next
        Return X
    End Function

    Sub GeneraPago(ID As Integer, FechaFira As Date, Capital As Decimal, Interes As Decimal, FechaHistoria As String)
        Dim Tipar As String = TaAnexos.tipar(ID)
        Dim rCalen As PagosFiraDS.CalendariosRow
        Dim rVenc As PagosFiraDS.VencimientosRow
        Dim SaldoCap As Decimal = TaEdoCta.SaldoCapital(ID, "BP")
        Dim Saldo As Decimal = TaEdoCta.SaldoContrato(ID)
        taVencimientos.FillFecha(DS.Vencimientos, ID, FechaFira)
        If DS.Vencimientos.Rows.Count > 0 Then
            taVencimientos.BorraVencimiento(ID, FechaFira)
        End If

        If Tipar = "H" Or Tipar = "C" Then
            ProcesaEstadoCuenta(ID, True, FechaFira)
            CorrigeCapitalVencimiento(ID)
        End If

        taVencimientos.FillPosteriores(DS.Vencimientos, ID, FechaFira)
        For xx As Integer = 0 To 0
            rVenc = DS.Vencimientos.Rows(xx)
            If taCaledarios.ExisteFecha(ID, FechaFira) > 0 Then
                taCaledarios.UpdateFecha(True, False, True, ID, FechaFira)
            Else
                If Capital = 0 And Interes > 0 Then
                    taCaledarios.Insert(ID, FechaFira, False, False, False, True, True)
                Else
                    taCaledarios.Insert(ID, FechaFira, True, False, True, True, True)
                End If
            End If
            If Capital + 0.1 >= rVenc.capital Then
                taVencimientos.Insert(FechaFira, rVenc.capital, FechaFira, "Vigente", rVenc.intereses, ID, 0)
                rVenc.Delete()
                DS.Vencimientos.GetChanges()
                taVencimientos.Update(DS.Vencimientos)
                taCaledarios.FillPosteriores(DS.Calendarios, ID, FechaFira)
                If DS.Calendarios.Rows.Count = 1 Then 'solo queda un vencimiento
                    rCalen = DS.Calendarios.Rows(0)
                    taCaledarios.BorraCalendario(rCalen.ID_Calendario)
                Else
                    rCalen = DS.Calendarios.Rows(0)
                    taCaledarios.BorraCalendario(rCalen.ID_Calendario)
                End If
                'Shell("\\server-raid\Jobs\MOD_PasivoFiraCalculos.exe " & ID, AppWinStyle.NormalFocus, True)
                taPagosFira.ProcesaPago(True, ID, FechaHistoria)
                ProcesaEstadoCuenta(ID, True, Today.Date)
                If TaEdoCta.SaldoCapital(ID, "BP") <= 0 Then
                    TaAnexos.TerminaContrato(ID)
                End If
                TaEdoCta.BorraCeros(ID)
            Else
                If Capital = 0 And Interes > 0 Then
                    taVencimientos.Insert(FechaFira, Capital, FechaFira, "Vigente", Interes, ID, Interes)
                Else
                    Capital += Interes
                    Interes = TaEdoCta.SaldoInteres(ID, "BP", FechaFira)
                    Capital -= Interes
                    taVencimientos.Insert(FechaFira, Capital, FechaFira, "Vigente", Interes, ID, -1)
                End If
                rVenc.capital -= (Capital)
                Capital = 0
                DS.Vencimientos.GetChanges()
                taVencimientos.Update(DS.Vencimientos)
                taPagosFira.ProcesaPago(True, ID, FechaHistoria)
            End If
        Next
    End Sub

End Module
