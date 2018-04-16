Module ModuleMain
    Dim TaVeciminetos As New PasivoFiraDSTableAdapters.CONT_CPF_vencimientosTableAdapter
    Dim TaEdoCta As New PasivoFiraDSTableAdapters.CONT_CPF_edocuentaTableAdapter
    Dim TaAnexos As New PasivoFiraDSTableAdapters.SaldosAnexosTableAdapter
    Dim TaSaldoConti As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Dim taCalendar As New PasivoFiraDSTableAdapters.CONT_CPF_CalendariosRevisionTasaTableAdapter
    Dim taCXSG As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    Dim ds As New PasivoFiraDS
    Dim subsidio As Boolean

    Sub Main()
        Dim Hoy As Date = "02/ene/2018"
        If CargaTIIE(Hoy) And Hoy.DayOfWeek <> DayOfWeek.Sunday And Hoy.DayOfWeek <> DayOfWeek.Saturday Then
            taCalendar.Fill(ds.CONT_CPF_CalendariosRevisionTasa, Hoy)
            For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres, Rc.VencimientoCapital)
                Console.WriteLine(Rc.Id_Contrato)
                taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
            Next
        Else
            Console.WriteLine("Error tasa Tiie : {0}", Hoy)
        End If
    End Sub

    Sub GeneraCorteInteres(Fecha As Date, ID_Contrato As Integer, EsCorteInte As Boolean, EsVencimientoCAP As Boolean)
        TaAnexos.Fill(ds.SaldosAnexos, ID_Contrato)
        For Each r As PasivoFiraDS.SaldosAnexosRow In ds.SaldosAnexos.Rows
            If CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMPLE_FIN(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMPLE(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                Procesa_FIJA_CON(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP)
            End If
        Next
    End Sub

    Sub Procesa_SIMPLE_FIN(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim TipoTasa As String
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow

        TipoTasa = "BP"

        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)
        CargaTIIE(r.FechaCorte)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        TIIE_old = TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
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
        If TasaActivaFN > 0 Then
            IntFN = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)
        End If

        If EsCorteInte = True Then
            Dim Fecha1 As Date
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            Fecha1 = TaEdoCta.SacaFecha1(r.id_contrato)
            Dim InteFinan As Decimal = TaEdoCta.SacaInteresAux1(r.id_contrato, Fecha1, Fecha)

            If EsVencimetoCap Then
                Dim Fecha_ante As DateTime = Fecha.AddDays(-3) 'DAGL  25/01/2018 Se restan 2 dias para traer el cap vigente, y en el query se agrega un between 
                CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
                ' CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato, Fecha)
                SaldoFIN -= CapitalVIG + IntFINAN + InteFinan
                IntORD_Aux = InteFinan * -1
                IntFB_Aux = 0
                IntFN_Aux = 0
            End If
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        InteFinan, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            InteFinan, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        InteFinan, 0, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC)
            End If

            TaSaldoConti.FillByUltimo(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (Rsaldo.cobertura_efectiva / 100)
                If saldonuevo > Rsaldo.monto_efectivo Then
                    TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                   SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                    TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
                End If
            Next

            CargaTIIE(Fecha)
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntFB + InteresAux1FB, r.InteresOrdinario, CapitalVIG, 0, r.id_contrato)
                If TaVeciminetos.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVeciminetos.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)


        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

    Sub Procesa_SIMPLE(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim TipoTasa As String
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow

        TipoTasa = "BP"
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)
        CargaTIIE(r.FechaCorte)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        TIIE_old = TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
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
        If TasaActivaFN > 0 Then
            IntFN = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Dim Fecha_ante As DateTime = Fecha.AddDays(-3) 'DAGL  25/01/2018 Se restan 2 dias para traer el cap vigente, y en el query se agrega un between 
                CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
                '  CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato, Fecha)
                SaldoFIN -= CapitalVIG + IntFINAN
                IntORD_Aux = IntFINAN * -1
                IntFB_Aux = 0
                IntFN_Aux = 0
            End If
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, 0, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC)
            End If

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (Rsaldo.cobertura_efectiva / 100)
                If saldonuevo > Rsaldo.monto_efectivo Then
                    TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                    TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
                End If
            Next

            CargaTIIE(Fecha)
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntFB + InteresAux1FB, 0, CapitalVIG, 0, r.id_contrato)
                If TaVeciminetos.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVeciminetos.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0)

        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

    Sub Procesa_FIJA_CON(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean)
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow
        Dim tasafija As Decimal = 0
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha)
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        ' InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaFinal, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        ' TIIE_old = TIIE28
        TIIE_old = TaAnexos.TIIEACTIVA(r.id_contrato)
        ' TIIE_old = Math.Round(TIIE_old, 2)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha)
        ' diasX = DateDiff(DateInterval.Day, r.FechaCorte, Fecha)
        ' diasY = DateDiff(DateInterval.Day, r.FechaCorte, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido
        SaldoINI = r.Capital + r.Vencido + r.InteresVencido

        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If diasX <> diasY And Minis_BASE > 0 Then
                IntORD = Math.Round((r.Capital + r.InteresOrdinario - Minis_BASE) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
                IntORD += Math.Round((Minis_BASE) * ((TasaActivaBP) / 100 / 360) * (diasY), 2)
            Else
                IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
            End If
        Else
            IntORD = 0
        End If
        If TaVeciminetos.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If
        '   IntFINAN = IntORD + IntVENC + InteresAux1 + InteresAux2
        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN

        tasafija = TaAnexos.TASAFIJA(r.id_contrato)
        IntFB = Math.Round((SaldoINI) * ((tasafija) / 100 / 360) * (diasX), 2)
        ' IntFN = Math.Round((SaldoINI) * ((TasaActivaFN + TIIE_old) / 100 / 360) * (diasX), 2)
        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN + TIIE_old) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Dim Fecha_ante As DateTime = Fecha.AddDays(-3) 'DAGL  25/01/2018 Se restan 2 dias para traer el cap vigente, y en el query se agrega un between 
                CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
                SaldoFIN -= CapitalVIG + IntFINAN
                IntORD_Aux = 0
                IntFB_Aux = 0
                IntFN_Aux = 0
            End If
            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (tasafija), diasX, IntFB_Aux, 0)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, 0, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN_Aux, IntVENC)
            End If

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (Rsaldo.cobertura_efectiva / 100)
                If saldonuevo > Rsaldo.monto_efectivo Then
                    TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                    TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
                End If
            Next

            CargaTIIE(Fecha) 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntFB + InteresAux1FB, 0, CapitalVIG, 0, r.id_contrato)
                If TaVeciminetos.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVeciminetos.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (tasafija), diasX, IntFB, 0)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN, 0)
            End If
            '  CargaTIIE(Fecha)
            '   TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        End If
        TaVeciminetos.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
    End Sub

    Sub CalculaServicioCobro(FecIni As Date, MontoBase As Decimal, PCXSG As Decimal, id_contrato As Integer, Subsidio As Boolean)
        Dim FecFin As Date = TaVeciminetos.SigFechaVenc(id_contrato)
        Dim Dias As Integer
        Dim Cobro As Decimal
        Dim ID As Integer = taCXSG.SacaIDContratoGarantia(id_contrato)
        Dim SubsidioAUX As Decimal

        If Subsidio Then
            SubsidioAUX = 2
        Else
            SubsidioAUX = 1
        End If

        Dias = DateDiff(DateInterval.Day, FecIni, FecFin) 'DAGL 26/01/2018 REVISAR FECHA INICIAL CUALDO LA FECHA DE VENCIMIENTO SEA MENOR A LA FECHA DE REVISION DE TASA
        If FecIni > FecFin Then
            Dim fechaux As Date = FecIni
            FecIni = FecFin
            FecFin = fechaux

        End If
        Dias = DateDiff(DateInterval.Day, FecIni, FecFin)
        Cobro = ((((MontoBase / SubsidioAUX) * (PCXSG / 100)) / 360)) * (Dias)

        taCXSG.Insert(FecIni, FecFin, Dias, FecIni, MontoBase, Cobro, Cobro * TasaIVA, Cobro * (1 + TasaIVA), PCXSG, ID, Subsidio) 'dagl 24/01/2018 se agrega el campo subsidio a la tabla cxg 
    End Sub

End Module
