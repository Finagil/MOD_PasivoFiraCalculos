Module ModuleMain
    Dim TaVeciminetos As New PasivoFiraDSTableAdapters.CONT_CPF_vencimientosTableAdapter
    Dim TaEdoCta As New PasivoFiraDSTableAdapters.CONT_CPF_edocuentaTableAdapter
    Dim TaAnexos As New PasivoFiraDSTableAdapters.SaldosAnexosTableAdapter
    Dim TaSaldoConti As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Dim taCalendar As New PasivoFiraDSTableAdapters.CONT_CPF_CalendariosRevisionTasaTableAdapter
    Dim ds As New PasivoFiraDS

    Sub Main()
        Dim Hoy As Date = "06/juL/2017"
        If CargaTIIE(Hoy) And Hoy.DayOfWeek <> DayOfWeek.Sunday And Hoy.DayOfWeek <> DayOfWeek.Saturday Then
            taCalendar.Fill(ds.CONT_CPF_CalendariosRevisionTasa, Hoy)
            For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres)
                Console.WriteLine(Rc.Id_Contrato)
                taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
            Next
        Else
            Console.WriteLine("Error tasa Tiie : {0}", Hoy)
        End If
    End Sub

    Sub GeneraCorteInteres(Fecha As Date, ID_Contrato As Integer, EsCorteInte As Boolean)
        TaAnexos.Fill(ds.SaldosAnexos, ID_Contrato)
        For Each r As PasivoFiraDS.SaldosAnexosRow In ds.SaldosAnexos.Rows
            If CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMPLE_FIN(Fecha, ID_Contrato, EsCorteInte, r)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMPLE(Fecha, ID_Contrato, EsCorteInte, r)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And InStr(r.des_tipo_tasa.Trim, "Fija con ") Then
                Procesa_FIJA_CON(Fecha, ID_Contrato, EsCorteInte, r)
            End If
        Next
    End Sub

    Sub Procesa_SIMPLE_FIN(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim IntFB As Decimal = 0
        Dim TipoTasa As String
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow

        TipoTasa = "BP"
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        CargaTIIE(r.FechaCorte)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        TIIE_old = TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)

        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL vIGENTE
            If diasX <> diasY And Minis_BASE > 0 Then
                IntORD = Math.Round((r.Capital + r.InteresOrdinario - Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
                IntORD += Math.Round((Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasY), 2)
            Else
                IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            End If
        Else
            IntORD = 0
        End If
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC + InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If EsCorteInte = True Then
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
            Next

            CargaTIIE(Fecha)
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        Else
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)
        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

    Sub Procesa_SIMPLE(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim IntFB As Decimal = 0
        Dim TipoTasa As String
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow

        TipoTasa = "BP"
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        CargaTIIE(r.FechaCorte)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        TIIE_old = TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)

        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL vIGENTE
            If diasX <> diasY And Minis_BASE > 0 Then
                IntORD = Math.Round((r.Capital + r.InteresOrdinario - Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
                IntORD += Math.Round((Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasY), 2)
            Else
                IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            End If
        Else
            IntORD = 0
        End If
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC + InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If EsCorteInte = True Then
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
            Next

            CargaTIIE(Fecha)
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        Else
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)
        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

    Sub Procesa_FIJA_CON(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim IntFB As Decimal = 0
        Dim TipoTasa As String
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow

        TipoTasa = "BP"
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        CargaTIIE(r.FechaCorte)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        TIIE_old = 0 ' TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)

        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL vIGENTE
            If diasX <> diasY And Minis_BASE > 0 Then
                IntORD = Math.Round((r.Capital + r.InteresOrdinario - Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
                IntORD += Math.Round((Minis_BASE) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasY), 2)
            Else
                IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            End If
        Else
            IntORD = 0
        End If
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC + InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If EsCorteInte = True Then
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
            Next

            CargaTIIE(Fecha)
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        Else
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)
        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

End Module
