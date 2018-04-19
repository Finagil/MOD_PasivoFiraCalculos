Module ModuleMain
    Dim TaVeciminetos As New PasivoFiraDSTableAdapters.CONT_CPF_vencimientosTableAdapter
    Dim TaEdoCta As New PasivoFiraDSTableAdapters.CONT_CPF_edocuentaTableAdapter
    Dim TaAnexos As New PasivoFiraDSTableAdapters.SaldosAnexosTableAdapter
    Dim TaMinis As New PasivoFiraDSTableAdapters.CONT_CPF_ministracionesTableAdapter
    Dim Ministraciones As New DescuentosDSTableAdapters.MinistracionesTableAdapter
    Dim TaSaldoConti As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Dim MFIRA As New PasivoFiraDSTableAdapters.mFIRATableAdapter
    Dim taCalendar As New PasivoFiraDSTableAdapters.CONT_CPF_CalendariosRevisionTasaTableAdapter
    Dim taCXSG As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    Dim taContraGarant As New PasivoFiraDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
    Dim tapagos As New PasivoFiraDSTableAdapters.PagosTableAdapter
    Dim ds As New PasivoFiraDS
    Dim ds1 As New DescuentosDS
    Dim subsidio As Boolean
    Dim Consec As Integer
    Dim taGarantias As New PasivoFiraDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
    Dim taCargosXservico As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    Dim SaldoCont As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Dim NoGarantias As Integer
    Dim SaldoINI, SaldoFIN, InteORD, InteORDFN, InteORDFB As Decimal
    Dim FechaUltimoMov As Date
    Public Nominal As Decimal = 0
    Public Efectiva As Decimal = 0
    Public ID_garantina As Integer = 0
    Public PCXSG_Aux As Decimal = 0
    Public FechaVecn As Date
    Dim FechaPago As Date
    Dim FechaFinal As Date
    Dim MontoBase As Decimal = 0
    Dim dias As Integer = 0
    Dim Cobro As Decimal = 0
    Dim id_contratoGarantia As Integer = 0
    Public FN, FB, BP As Decimal




    Sub Main()
        Dim Fila As Integer
        Dim Args() As String = Environment.GetCommandLineArgs()
        Dim Hoy As Date
        Dim ID As Integer = Args(1)
        ID = 96
        Dim Hasta As Date = "02/abr/2018" 'Today.AddDays((Today.Day - 1) * -1).Date
        Dim Tipar As String = TaAnexos.tipar(ID)
        If Tipar = "H" Or Tipar = "C" Then
            Dim x As Integer
            Dim Anexo, Ciclo As String
            Dim mfira As New PasivoFiraDSTableAdapters.mFIRATableAdapter
            Dim Fechas() As Date
            Dim Importes() As Decimal
            Dim HastaX As Date
            Anexo = TaAnexos.anexo(ID)
            Ciclo = TaAnexos.ciclo(ID)
            TaEdoCta.BorraTodo(ID)
            'Hoy = TaEdoCta.SacaFecha1(ID)
            TaVeciminetos.UpdateStatusVencimiento("Vigente", "Vencido", ID, ID)
            TaSaldoConti.BorraSaldoContigente(ID)
            taCXSG.BorraCSG(ID)
            tapagos.BorraPagos(ID)
            TaMinis.BorraMinistraciones(ID)
            mfira.FillByOtorgado(ds.mFIRA, Anexo, Ciclo)
            For Each rx As PasivoFiraDS.mFIRARow In ds.mFIRA.Rows
                x += 1
                ReDim Preserve Fechas(x)
                ReDim Preserve Importes(x)
                Fechas(x) = CtoD(rx.FechaProgramada)
                Importes(x) = rx.Importe
            Next
            x = 0
            CargaTIIE(Fechas(1), "", "")
            TaAnexos.UpdateFechaCorteTIIE(Fechas(1), TIIE28, ID)

            For Each rx As PasivoFiraDS.mFIRARow In ds.mFIRA.Rows
                x += 1
                'InsertaMinistracion(ID, rx)
                Hoy = Fechas(x)
                CalculaServicioCobro(Hoy, ID)

                If x = Fechas.Length - 1 Then
                    HastaX = Hasta
                Else
                    HastaX = Fechas(x + 1)
                End If
                CargaTIIE(Hoy, "", "")
                TaAnexos.UpdateFechaCorteTIIE(Hoy, TIIE28, ID)

                While Hoy <= HastaX
                    If CargaTIIE(Hoy, "", "") And Hoy.DayOfWeek <> DayOfWeek.Sunday And Hoy.DayOfWeek <> DayOfWeek.Saturday Then
                        taCalendar.FillByIdContrato(ds.CONT_CPF_CalendariosRevisionTasa, Hoy, ID)
                        For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                            GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres, Rc.VencimientoCapital)
                            Console.WriteLine(Rc.Id_Contrato)
                            taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
                        Next
                    Else
                        Console.WriteLine("Error tasa Tiie : {0}", Hoy)
                    End If
                    Hoy = Hoy.AddDays(1)
                End While
            Next





        Else
            TaEdoCta.BorraDatos(ID)
            Hoy = TaEdoCta.SacaFecha1(ID)
            CargaTIIE(Hoy, "", "")
            TaAnexos.UpdateFechaCorteTIIE(Hoy, TIIE28, ID)
            TaVeciminetos.UpdateStatusVencimiento("Vigente", "Vencido", ID, ID)
            TaSaldoConti.BorraSaldoContigente(ID)
            taCXSG.BorraCSG(ID)
            tapagos.BorraPagos(ID)
            While Hoy <= Hasta
                If CargaTIIE(Hoy, "", "") And Hoy.DayOfWeek <> DayOfWeek.Sunday And Hoy.DayOfWeek <> DayOfWeek.Saturday Then
                    taCalendar.FillByIdContrato(ds.CONT_CPF_CalendariosRevisionTasa, Hoy, ID)
                    For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                        GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres, Rc.VencimientoCapital)
                        Console.WriteLine(Rc.Id_Contrato)
                        taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
                    Next
                Else
                    Console.WriteLine("Error tasa Tiie : {0}", Hoy)
                End If
                Hoy = Hoy.AddDays(1)
            End While
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

    Sub CalculaServicioCobro(hoy As Date, idcont As Integer)
        Dim FECHAPROG, FECHAPROG_AUX As String
        Dim tasafira As Decimal = 0
        Dim ciclo As String
        Dim dias As Integer
        Dim ANEXOAUX As String
        Dim PCXSG As Decimal = 0
        Dim cont As Integer
        Dim fonaga As String
        Dim gl_mosusa As Decimal = 0


        'funcion para agregar ministraciones firam
        'traer el valor de la tasafira
        tasafira = TaAnexos.TASAFIJA(idcont)
        ANEXOAUX = TaAnexos.anexo(idcont)
        fonaga = TaAnexos.fonaga(idcont)
        'revisar si hay ministraciones del contrato 
        'ciclo de contrato
        ciclo = TaAnexos.ciclo(idcont)
        'Me.MinistracionesTableAdapter.FillByIDContrato(Me.DescuentosDS.Ministraciones, ID_Contrato, ID_Contrato)
        ' For Each row As DataRow In MFIRA(ds.CONT_CPF_ministraciones, (Rc.Id_Contrato, (Rc.Id_Contrato)
        FECHAPROG = hoy.ToString("yyyyMMdd")


        Dim Subsidio As Decimal
        Dim subsidiox As Boolean
        Dim FechaPago As Date = hoy
        Dim FechaFinal As Date
        Dim cobro As Decimal = 0
        Dim PCXSG_Aux As Decimal = 0

        Dim MontoBase As Decimal
        Dim MIN As Integer
        Dim fec As Date
        Dim tipar As String = TaAnexos.tipar(idcont)
        Dim anexo1 As String = TaAnexos.anexo(idcont)
        Dim tipta As String = TaAnexos.tipta(idcont)
        cont = TaMinis.ministraciones_contrato(idcont)

        '  If cont <= 0 Then 'para la primera ministracion
        FECHAPROG = MFIRA.FECMIN1(idcont, hoy.ToString("yyyyMMdd"))
        FECHAPROG_AUX = hoy.ToString("yyyyMMdd")
        '
        ' Else
        'ir a ministraciones y traer el numero de ministracion a excluir
        'traer la fecha programada
        '
        ' End If





        For Each row As DataRow In MFIRA.GetMINPORCONTRATO(idcont, hoy.ToString("yyyyMMdd"))

            MontoBase = row("Importe")
            Dim fec1 As String = row("FechaProgramada")
            fec = fec1.Substring(6, 2) + "/" + fec1.Substring(4, 2) + "/" + fec1.Substring(0, 4)
            Dim fec2 As Date = fec.AddDays(-1)
            MIN = row("Ministracion")
            'MFIRA.Updateprocesado("True", anexo1, ciclo, MIN)
            If tipar = "H" Or tipar = "C" Or tipar = "A" Then
                FechaFinal = TaMinis.primervenavi(anexo1, ciclo)
                '    MontoBase = MinistracionesBindingSource.Current("Importe")
            Else
                FechaFinal = TaMinis.primerventra(anexo1)
                '  MontoBase = MinistracionesBindingSource.Current("MontoFinanciado")
            End If
            PCXSG = TaAnexos.pcxs(idcont)
            dias = DateDiff(DateInterval.Day, fec, FechaFinal)
            subsidiox = TaAnexos.Subsidiocontrato(idcont)
            If subsidiox = True Then
                Subsidio = 2
            Else
                Subsidio = 1
            End If


            cobro = ((((MontoBase / Subsidio) * (PCXSG / 100)) / 360)) * (dias)

            PCXSG_Aux = PCXSG / Subsidio



            'DESCUENTOS



            Dim Consec As Integer = MIN
            Dim taGarantias As New DescuentosDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
            Dim taEdoCta As New DescuentosDSTableAdapters.CONT_CPF_edocuentaTableAdapter
            Dim taCargosXservico As New DescuentosDSTableAdapters.CONT_CPF_csgTableAdapter
            Dim SaldoCont As New DescuentosDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
            Dim NoGarantias As Integer = taGarantias.EXISTENGARANTIAS(idcont)
            Dim SaldoINI, SaldoFIN, InteORD, InteORDFN, InteORDFB As Decimal
            Dim FechaUltimoMov As Date

            SaldoINI = taEdoCta.SaldoContrato(idcont)
            SaldoFIN = SaldoINI + MontoBase
            Dim importe As Decimal = 0
            Dim monto As Decimal = 0
            Dim iva As Decimal = 0
            Dim ID_garantina As Integer
            importe = cobro
            iva = importe * 0.16
            monto = importe + iva
            TaMinis.InsertQueryMinistraciones(MontoBase, fec, Consec, PCXSG_Aux, iva, importe, idcont, "Otorgado", fec)
            Ministraciones.Descontar(anexo1, ciclo, FechaPago.ToString("yyyyMMdd"))

            ' Taminis.InsertQueryMinistracion(monto, FECHAPROG, Consec, PCXSG_Aux, iva, importe, idcont, "Solicitado", FECHAPROG)
            For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
                fonaga = TaAnexos.fonaga(idcont)
                gl_mosusa = TaAnexos.gl_mosusa(idcont)
                If fonaga = "SI" Then
                    ' PCXSG = PCXSG_FONAGA
                    ID_garantina = 2 ' id tabla de garantias
                    If gl_mosusa = 0 Then
                        Nominal = 45
                        Efectiva = 45
                    Else
                        Nominal = 50
                        Efectiva = 45
                    End If
                Else
                    'F2.PCXSG = PCXSG_FEGA
                    ID_garantina = 1 ' id tabla de garantias
                    Nominal = 50
                    Efectiva = 50
                End If


                FB = row1("FB")
                BP = row1("BP")
                FN = row1("FN")
            Next
            If NoGarantias = 0 Then
                taGarantias.Insert(idcont, ID_garantina, Nominal, MontoBase * (Nominal / 100), Efectiva, True)
                'Me.CONT_CPF_contratosTableAdapter.Updatesubsidio(CheckBox1.Checked, ID_Contrato)


                '            Else

                ' For Each row1 As DataRow In Taminis.GetDataByIDCONTRATO(idcont, idcont)






                ' Nominal = row1("Cobertura_Nominal")
                'Efectiva = row1("Cobertura_Efectiva")
                'taGarantias.UpdateSaldoConti(SaldoFIN * (Efectiva / 100), idcont)
                'FB = row1("FB")
                'BP = row1("BP")
                'FN = row1("FN")
                'Next
            End If




            If SaldoINI > 0 Then
                Dim IDC As Integer = idcont
                For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
                    CargaTIIE(row1("FechaCorte"), row1("Tipta"), row1("ClaveEsquema"))
                    FechaUltimoMov = TaMinis.Fechaultimomov(idcont)
                    fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
                    Dim DiasX As Integer = DateDiff(DateInterval.Day, FechaUltimoMov, fec)
                    InteORD = SaldoINI * ((BP + TIIE_Aplica) / 100 / 360) * DiasX
                    If row1("Tipta") = "7" Then 'DAGL 25/01/2018 En tasa fija se resta el valor FB
                        InteORDFB = SaldoINI * ((FB + tasafira) / 100 / 360) * DiasX
                    Else
                        InteORDFB = SaldoINI * ((FB + TIIE_Aplica) / 100 / 360) * DiasX
                    End If

                    InteORDFN = SaldoINI * ((FN + TIIE_Aplica) / 100 / 360) * DiasX

                Next
            Else
                fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
                ' FechaUltimoMov = fec.ToShortDateString

                For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
                    CargaTIIE(fec, row1("Tipta"), row1("ClaveEsquema"))
                    '  If MinistracionesBindingSource.Current("Tipta") = "7" Then TIIE_Aplica = TxttasaFira.Text - FB 'DAGL 25/01/2018 En tasa fija se resta el valor FB
                    FechaUltimoMov = fec.ToShortDateString
                Next


            End If
            id_contratoGarantia = TaMinis.sacaid(idcont)
            Dim subsidioaux As Boolean
            subsidioaux = TaAnexos.Subsidiocontrato(idcont) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos
            fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
            For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
                taCargosXservico.Insert(fec, FechaFinal, dias, Date.Now, MontoBase, cobro, cobro * TasaIVA, cobro * (1 + TasaIVA), PCXSG, id_contratoGarantia, subsidioaux)
                SaldoCont.Insert(fec, Nothing, Nothing, Nothing, Nothing, MontoBase, SaldoFIN, Nominal, Efectiva, SaldoFIN * (Nominal / 100), SaldoFIN * (Efectiva / 100), id_contratoGarantia)


                If row1("FN") > 0 Then
                    'If MinistracionesBindingSource.Current("Tipta") = "7" Then
                    '  taEdoCta.Insert("FN", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, id, TIIE_Aplica, 0, InteORDFN, 0)
                    'Else
                    taEdoCta.Insert("FN", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FN + TIIE_Aplica, 0, InteORDFN, 0)
                    'End If

                End If

                If row1("Tipta") = "7" Then
                    taEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP, 0, InteORD, 0)
                Else
                    taEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP + TIIE_Aplica, 0, InteORD, 0)
                End If

                If row1("Tipta") = "7" Then
                    ' tasafira = TxttasaFira.Text 'DAGL 25/01/2018 En tasa fija se resta el valor FB

                    taEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, tasafira, 0, InteORDFB, 0) 'tasafijafira
                Else
                    taEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FB + TIIE_Aplica, 0, InteORDFB, 0)

                End If
                TaAnexos.UpdateFechaCorteTIIE(fec, TIIE28, idcont)
            Next

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
        TasaActivaFn = TaAnexos.TasaActivaFN(r.id_contrato)
        CargaTIIE(r.FechaCorte, "", "")
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
                CapitalVIG = TaVeciminetos.CapitalVencimiento(r.id_contrato, Fecha)
                'CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
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

            If ds.CONT_CPF_saldos_contingente.Rows.Count <= 0 Then 'inserta el primer saldo contingente
                taContraGarant.FillByidContrato(ds.CONT_CPF_contratos_garantias, ID_Contrato)
                Dim rz As PasivoFiraDS.CONT_CPF_contratos_garantiasRow
                rz = ds.CONT_CPF_contratos_garantias.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (rz.cobertura_efectiva / 100)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, rz.cobertura_nominal, rz.cobertura_efectiva,
                                   SaldoFIN * (rz.cobertura_nominal / 100), SaldoFIN * (rz.cobertura_efectiva / 100), rz.id_contrato_garantia)
            End If

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (Rsaldo.cobertura_efectiva / 100)
                If saldonuevo > Rsaldo.monto_efectivo Then
                    TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                   SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                    TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
                End If

            Next

            CargaTIIE(Fecha, "", "")
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
        CargaTIIE(r.FechaCorte, "", "")
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
                CapitalVIG = TaVeciminetos.CapitalVencimiento(r.id_contrato, Fecha)
                'CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
                '  CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato, Fecha)
                SaldoFIN -= CapitalVIG + IntFINAN
                IntORD_Aux = IntFINAN * -1
                IntFB_Aux = 0
                IntFN_Aux = 0
            End If
            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_aux, 0)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, 0, ID_Contrato, (TasaActivaFN), diasX, IntFN_aux, IntVENC)
            End If

            TaSaldoConti.Fill(ds.CONT_CPF_saldos_contingente, r.id_contrato_garantia)

            If ds.CONT_CPF_saldos_contingente.Rows.Count <= 0 Then 'inserta el primer saldo contingente
                taContraGarant.FillByidContrato(ds.CONT_CPF_contratos_garantias, ID_Contrato)
                Dim rz As PasivoFiraDS.CONT_CPF_contratos_garantiasRow
                rz = ds.CONT_CPF_contratos_garantias.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (rz.cobertura_efectiva / 100)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, rz.cobertura_nominal, rz.cobertura_efectiva,
                                   SaldoFIN * (rz.cobertura_nominal / 100), SaldoFIN * (rz.cobertura_efectiva / 100), rz.id_contrato_garantia)
            End If

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
            Next

            CargaTIIE(Fecha, "", "")
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

        CargaTIIE(Fecha, "", "")
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
                CapitalVIG = TaVeciminetos.CapitalVencimiento(r.id_contrato, Fecha)
                'CapitalVIG = TaVeciminetos.CapitalVigente(ID_Contrato)
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

            If ds.CONT_CPF_saldos_contingente.Rows.Count <= 0 Then 'inserta el primer saldo contingente
                taContraGarant.FillByidContrato(ds.CONT_CPF_contratos_garantias, ID_Contrato)
                Dim rz As PasivoFiraDS.CONT_CPF_contratos_garantiasRow
                rz = ds.CONT_CPF_contratos_garantias.Rows(0)
                Dim saldonuevo As Decimal = SaldoFIN * (rz.cobertura_efectiva / 100)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, rz.cobertura_nominal, rz.cobertura_efectiva,
                                   SaldoFIN * (rz.cobertura_nominal / 100), SaldoFIN * (rz.cobertura_efectiva / 100), rz.id_contrato_garantia)
            End If

            For Each Rsaldo In ds.CONT_CPF_saldos_contingente.Rows
                Rsaldo = ds.CONT_CPF_saldos_contingente.Rows(0)
                TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoFIN, Rsaldo.cobertura_nominal, Rsaldo.cobertura_efectiva,
                                    SaldoFIN * (Rsaldo.cobertura_nominal / 100), SaldoFIN * (Rsaldo.cobertura_efectiva / 100), Rsaldo.id_contrato_garantia)
                TaSaldoConti.UpdateSaldoConti(SaldoFIN * (Rsaldo.cobertura_efectiva / 100), ID_Contrato, 0)
            Next

            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
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
