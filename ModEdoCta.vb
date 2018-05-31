Module ModEdoCta

    Dim ds As New PasivoFiraDS
    Dim ds1 As New DescuentosDS
    Dim subsidio As Boolean
    Dim Consec As Integer
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
    Dim diasProm As Integer = 0
    Dim Cobro As Decimal = 0
    Dim id_contratoGarantia As Integer = 0
    Dim Tipar As String
    Public FN, FB, BP As Decimal
    Dim rCalen As PagosFiraDS.CalendariosRow

    Sub ProcesaEstadoCuenta(ID As Integer, Continuo As Boolean)
        Dim Hoy As Date
        Dim Hasta As Date
        Dim Aux As Decimal
        Try
            Hasta = Today.Date
            Tipar = TaAnexos.tipar(ID)
            If TaAnexos.ExisteContrato(ID) <= 0 Then
                Console.WriteLine("NO EXISTE CONTRATO " & ID)
                Exit Sub

            End If
            If TaAnexos.SacaEstatus(ID) <> "ACTIVO" Then
                Console.WriteLine("Contrato TERMINADO " & ID)
                Exit Sub
            End If
            If ID = 411 Then
                Tipar = "H"
            End If
            If Tipar = "H" Or Tipar = "C" Then
                CorrigeCapitalVencimiento(ID)
                diasProm = 99
                Dim Anexo, Ciclo As String
                Dim mfira As New PasivoFiraDSTableAdapters.mFIRATableAdapter
                Anexo = TaAnexos.anexo(ID)
                Ciclo = TaAnexos.ciclo(ID)
                If ID = 411 Then
                    Hoy = "10/04/2018"
                Else
                    Hoy = TaEdoCta.SacaFecha1(ID) '1 saca la fecha
                End If
                TaEdoCta.BorraTodo(ID) '2 borra todo
                CargaTIIE(Hoy, "", "") '3.1 saca tiie
                TaAnexos.UpdateFechaCorteTIIE(Hoy, TIIE28, ID) '3.2 saca tiie
                Aux = mfira.ministracionporfecha(Anexo, Hoy.ToString("yyyyMMdd"), Ciclo)
                TaEdoCta.Insert("BP", Hoy, Hoy, 0, Aux, 0, 0, 0, 0, 0, 0, 0, 0, Aux, ID, TIIE28, 0, 0, 0, 0) ' 4 inserta primera linea
                TaVevcimientosCPF.UpdateStatusVencimiento("Vigente", "Vencido", ID, ID)
                TaSaldoConti.BorraSaldoContigente(ID)
                taCXSG.BorraCSG(ID)
                tapagos.BorraPagos(ID)
                TaMinis.BorraMinistraciones(ID)
                mfira.FillByOtorgado(ds.mFIRA, Anexo, Ciclo)
                For Each rx As PasivoFiraDS.mFIRARow In ds.mFIRA.Rows
                    If TaMinis.ExisteMinistracion(ID, CtoD(rx.FechaProgramada)) <= 0 Then
                        CalculaServicioCobroX(CtoD(rx.FechaProgramada), ID)
                    End If
                    If taCalendar.ExisteFecha(ID, CtoD(rx.FechaProgramada)) <= 0 Then
                        taCalendar.Insert(ID, CtoD(rx.FechaProgramada), 0, 0, 0, 0, 0)
                    End If
                Next
                Hoy = Hoy.AddDays(1)

                While Hoy <= Hasta
                    taCalendar.FillByIdContrato(ds.CONT_CPF_CalendariosRevisionTasa, Hoy, ID)
                    For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                        GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres, Rc.VencimientoCapital, Rc.AcumulaInteres, Rc.RevisionTasa)
                        Console.WriteLine(Rc.Id_Contrato)
                        taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
                    Next
                    Hoy = Hoy.AddDays(1)
                End While
                CorrigeCapitalVencimiento(ID)
            Else
                TaEdoCta.BorraDatos(ID)
                Hoy = TaEdoCta.SacaFecha1(ID)
                CargaTIIE(Hoy, "", "")
                TaAnexos.UpdateFechaCorteTIIE(Hoy, TIIE28, ID)
                TaVevcimientosCPF.UpdateStatusVencimiento("Vigente", "Vencido", ID, ID)
                TaSaldoConti.BorraSaldoContigente(ID)
                taCXSG.BorraCSG(ID)
                tapagos.BorraPagos(ID)
                diasProm = DiasEntreVecn(ID)
                While Hoy <= Hasta
                    If CargaTIIE(Hoy, "", "") Then
                        Console.WriteLine("Procesando " & Hoy.ToShortDateString)
                        taCalendar.FillByIdContrato(ds.CONT_CPF_CalendariosRevisionTasa, Hoy, ID)
                        For Each Rc As PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaRow In ds.CONT_CPF_CalendariosRevisionTasa.Rows
                            GeneraCorteInteres(Hoy, Rc.Id_Contrato, Rc.VencimientoInteres, Rc.VencimientoCapital, Rc.AcumulaInteres, Rc.RevisionTasa)
                            Console.WriteLine(Rc.Id_Contrato)
                            taCalendar.ProcesaCalendario(True, Rc.ID_Calendario, Rc.ID_Calendario)
                        Next
                    Else
                        Console.WriteLine("Error tasa Tiie : {0}", Hoy)
                    End If
                    Hoy = Hoy.AddDays(1)
                End While
            End If
            TaEdoCta.BorraCeros(ID)
        Catch ex As Exception
            Console.WriteLine("Error: ID-" & ID & " " & ex.Message & " " & Date.Now)
            taCorreos.Insert("PasivoFira@finagil.com.mx", "ecacerest@finagil.com.mx", "Error: " & ID, ex.Message, False, Date.Now, "")
        End Try
        If Continuo <> True Then
            End
        End If
    End Sub

    Sub GeneraCorteInteres(Fecha As Date, ID_Contrato As Integer, EsCorteInte As Boolean, EsVencimientoCAP As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean)
        Dim r As PasivoFiraDS.SaldosAnexosRow
        TaAnexos.Fill(ds.SaldosAnexos, ID_Contrato)
        For Each r In ds.SaldosAnexos.Rows
            Console.WriteLine(ID_Contrato & " " & Fecha.ToShortDateString & r.des_tipo_tasa & " Esquema: " & r.claveCobro.Trim)

            If CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMPLE_VAR(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, False)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                Procesa_SIMPLE_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                If diasProm < 32 Then
                    Procesa_SIMPLE_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
                Else
                    Procesa_SIMPLE_CON_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
                End If
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And InStr(r.des_tipo_tasa.Trim, "Variable") Then  '"Fija con "
                If diasProm < 32 Then
                    Procesa_SIMPLE_VAR(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
                Else
                    Procesa_SIMPLE_CON_VAR(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
                End If
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.TRADICIONAL And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                Procesa_TRADICIONAL_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, True)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMFAA And InStr(r.des_tipo_tasa.Trim, "Variable") Then
                Procesa_SIMFAA_VAR(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, False)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.SIMFAA And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                Procesa_SIMFAA_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, False)
            ElseIf CInt(r.claveCobro.Trim) = EsquemaCobro.COBRO_MENSUAL And InStr(r.des_tipo_tasa.Trim, "Tasa Fija con Pago") Then  '"Fija con "
                Procesa_COBRO_MENSUAL_FIJA(Fecha, ID_Contrato, EsCorteInte, r, EsVencimientoCAP, AcumulaInteres, RevisionTasa, False)
            Else
                Console.WriteLine("SIN ESQUEMA DE COBRO....")
                Console.WriteLine("SIN ESQUEMA DE COBRO....")
                Console.WriteLine("SIN ESQUEMA DE COBRO....")
                Console.WriteLine("SIN ESQUEMA DE COBRO....")
            End If
        Next
        'TERMINA CONTRATOS+++++++++++++++++++++++++++++++++++++++++++
        TaAnexos.Fill(ds.SaldosAnexos, ID_Contrato)
        If ds.SaldosAnexos.Rows.Count > 0 Then
            r = ds.SaldosAnexos.Rows(0)
            If r.SaldoInsoluto = 0 And r.Ministrado > 0 Then
                TaAnexos.TerminaContrato(r.id_contrato)
            End If
        End If
        'TERMINA CONTRATOS+++++++++++++++++++++++++++++++++++++++++++
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
            'Dim taGarantias As New DescuentosDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
            'Dim taEdoCta As New DescuentosDSTableAdapters.CONT_CPF_edocuentaTableAdapter
            'Dim taCargosXservico As New DescuentosDSTableAdapters.CONT_CPF_csgTableAdapter
            'Dim SaldoCont As New DescuentosDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
            Dim NoGarantias As Integer = taGarantias.ExistenGarantias(idcont)
            Dim SaldoINI, SaldoFIN, InteORD, InteORDFN, InteORDFB As Decimal
            Dim FechaUltimoMov As Date

            SaldoINI = TaEdoCta.SaldoContrato(idcont)
            SaldoFIN = SaldoINI + MontoBase
            Dim importe As Decimal = 0
            Dim monto As Decimal = 0
            Dim iva As Decimal = 0
            Dim ID_garantina As Integer
            Dim DiasX As Integer
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
            End If




            If SaldoINI > 0 Then
                Dim IDC As Integer = idcont
                For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
                    CargaTIIE(row1("FechaCorte"), row1("Tipta"), row1("ClaveEsquema"))
                    FechaUltimoMov = TaMinis.Fechaultimomov(idcont)
                    fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
                    DiasX = DateDiff(DateInterval.Day, FechaUltimoMov, fec)
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
                    TaEdoCta.Insert("FN", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FN + TIIE_Aplica, DiasX, InteORDFN, 0, 0)
                    'End If

                End If

                If row1("Tipta") = "7" Then
                    TaEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP, DiasX, 0, 0, InteORD)
                Else
                    TaEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP + TIIE_Aplica, DiasX, 0, 0, InteORD)
                End If

                If row1("Tipta") = "7" Then
                    ' tasafira = TxttasaFira.Text 'DAGL 25/01/2018 En tasa fija se resta el valor FB

                    TaEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, tasafira, DiasX, InteORDFB, 0, 0) 'tasafijafira
                Else
                    TaEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FB + TIIE_Aplica, DiasX, InteORDFB, 0, 0)

                End If
                TaAnexos.UpdateFechaCorteTIIE(fec, TIIE28, idcont)
            Next

        Next
    End Sub

    Sub CalculaServicioCobroX(hoy As Date, idcont As Integer)
        Dim FECHAPROG, FECHAPROG_AUX As String
        Dim tasafira As Decimal = 0
        Dim ciclo As String
        Dim dias As Integer
        Dim ANEXOAUX As String
        Dim PCXSG As Decimal = 0
        Dim cont As Integer
        Dim fonaga As String
        Dim gl_mosusa As Decimal = 0

        tasafira = TaAnexos.TASAFIJA(idcont)
        ANEXOAUX = TaAnexos.anexo(idcont)
        fonaga = TaAnexos.fonaga(idcont)
        ciclo = TaAnexos.ciclo(idcont)
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

        FECHAPROG = MFIRA.FECMIN1(idcont, hoy.ToString("yyyyMMdd"))
        FECHAPROG_AUX = hoy.ToString("yyyyMMdd")

        For Each row As DataRow In MFIRA.GetMINPORCONTRATO(idcont, hoy.ToString("yyyyMMdd"))
            MontoBase = row("Importe")
            Dim fec1 As String = row("FechaProgramada")
            fec = fec1.Substring(6, 2) + "/" + fec1.Substring(4, 2) + "/" + fec1.Substring(0, 4)
            Dim fec2 As Date = fec.AddDays(-1)
            MIN = row("Ministracion")
            FechaFinal = TaMinis.primervenavi(anexo1, ciclo)
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

            Dim Consec As Integer = MIN
            Dim NoGarantias As Integer = taGarantias.ExistenGarantias(idcont)
            Dim SaldoINI, SaldoFIN, InteORD, InteORDFN, InteORDFB As Decimal
            Dim FechaUltimoMov As Date

            SaldoINI = TaEdoCta.SaldoContrato(idcont)
            SaldoFIN = SaldoINI + MontoBase
            Dim importe As Decimal = 0
            Dim monto As Decimal = 0
            Dim iva As Decimal = 0
            Dim ID_garantina As Integer
            Dim DiasX As Integer
            importe = cobro
            iva = importe * 0.16
            monto = importe + iva
            TaMinis.InsertQueryMinistraciones(MontoBase, fec, Consec, PCXSG_Aux, iva, importe, idcont, "Otorgado", fec)
            Ministraciones.Descontar(anexo1, ciclo, FechaPago.ToString("yyyyMMdd"))

            'For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
            '    fonaga = TaAnexos.fonaga(idcont)
            '    gl_mosusa = TaAnexos.gl_mosusa(idcont)
            '    If fonaga = "SI" Then
            '        ID_garantina = 2 ' id tabla de garantias
            '        If gl_mosusa = 0 Then
            '            Nominal = 45
            '            Efectiva = 45
            '        Else
            '            Nominal = 50
            '            Efectiva = 45
            '        End If
            '    Else
            '        ID_garantina = 1 ' id tabla de garantias
            '        Nominal = 50
            '        Efectiva = 50
            '    End If

            '    FB = row1("FB")
            '    BP = row1("BP")
            '    FN = row1("FN")
            'Next
            'If NoGarantias = 0 Then
            '    taGarantias.Insert(idcont, ID_garantina, Nominal, MontoBase * (Nominal / 100), Efectiva, True)
            'End If

            'If SaldoINI > 0 Then
            '    Dim IDC As Integer = idcont
            '    For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
            '        CargaTIIE(row1("FechaCorte"), row1("Tipta"), row1("ClaveEsquema"))
            '        FechaUltimoMov = TaMinis.Fechaultimomov(idcont)
            '        fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
            '        DiasX = DateDiff(DateInterval.Day, FechaUltimoMov, fec)
            '        InteORD = SaldoINI * ((BP + TIIE_Aplica) / 100 / 360) * DiasX
            '        If row1("Tipta") = "7" Then 'DAGL 25/01/2018 En tasa fija se resta el valor FB
            '            InteORDFB = SaldoINI * ((FB + tasafira) / 100 / 360) * DiasX
            '        Else
            '            InteORDFB = SaldoINI * ((FB + TIIE_Aplica) / 100 / 360) * DiasX
            '        End If
            '        InteORDFN = SaldoINI * ((FN + TIIE_Aplica) / 100 / 360) * DiasX
            '    Next
            'Else
            '    fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
            '    For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
            '        CargaTIIE(fec, row1("Tipta"), row1("ClaveEsquema"))
            '        FechaUltimoMov = fec.ToShortDateString
            '    Next
            'End If
            'id_contratoGarantia = TaMinis.sacaid(idcont)
            'Dim subsidioaux As Boolean
            'subsidioaux = TaAnexos.Subsidiocontrato(idcont) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos
            'fec = FECHAPROG_AUX.Substring(6, 2) + "/" + FECHAPROG_AUX.Substring(4, 2) + "/" + FECHAPROG_AUX.Substring(0, 4)
            'For Each row1 As DataRow In TaMinis.GetDataByIDCONTRATO(idcont, idcont)
            '    taCargosXservico.Insert(fec, FechaFinal, dias, Date.Now, MontoBase, cobro, cobro * TasaIVA, cobro * (1 + TasaIVA), PCXSG, id_contratoGarantia, subsidioaux)
            '    SaldoCont.Insert(fec, Nothing, Nothing, Nothing, Nothing, MontoBase, SaldoFIN, Nominal, Efectiva, SaldoFIN * (Nominal / 100), SaldoFIN * (Efectiva / 100), id_contratoGarantia)


            '    If row1("FN") > 0 Then
            '        'If MinistracionesBindingSource.Current("Tipta") = "7" Then
            '        '  taEdoCta.Insert("FN", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, id, TIIE_Aplica, 0, InteORDFN, 0)
            '        'Else
            '        TaEdoCta.Insert("FN", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FN + TIIE_Aplica, DiasX, InteORDFN, 0, 0)
            '        'End If

            '    End If

            '    If row1("Tipta") = "7" Then
            '        TaEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP, DiasX, 0, 0, InteORD)
            '    Else
            '        TaEdoCta.Insert("BP", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, BP + TIIE_Aplica, DiasX, 0, 0, InteORD)
            '    End If

            '    If row1("Tipta") = "7" Then
            '        ' tasafira = TxttasaFira.Text 'DAGL 25/01/2018 En tasa fija se resta el valor FB

            '        TaEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, tasafira, DiasX, InteORDFB, 0, 0) 'tasafijafira
            '    Else
            '        TaEdoCta.Insert("FB", FechaUltimoMov, fec.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, idcont, FB + TIIE_Aplica, DiasX, InteORDFB, 0, 0)

            '    End If
            '    TaAnexos.UpdateFechaCorteTIIE(fec, TIIE28, idcont)
            'Next

        Next
    End Sub

    Function DiasEntreVecn(ID) As Integer
        TaVevcimientosCPF.FillByID(ds.CONT_CPF_vencimientos, ID)
        Dim Acum, Cont As Decimal
        Dim Fec As Date
        For Each r As PasivoFiraDS.CONT_CPF_vencimientosRow In ds.CONT_CPF_vencimientos.Rows
            If Fec <= CDate("01/01/01") Then
                Fec = r.fecha
            Else
                Acum = DateDiff(DateInterval.Day, Fec, r.fecha)
                Cont += 1
            End If
        Next
        If Cont = 0 Then
            DiasEntreVecn = 33 ' solo un vencimiento
        Else
            DiasEntreVecn = (Acum / Cont)
        End If
        Return DiasEntreVecn
    End Function

    Sub CalculaServicioCobro(FecIni As Date, MontoBase As Decimal, PCXSG As Decimal, id_contrato As Integer, Subsidio As Boolean)
        Dim FecFin As Date
        Try
            FecFin = TaVevcimientosCPF.SigFechaVenc(id_contrato)
        Catch ex As Exception
            TaAnexos.TerminaContrato(id_contrato)
            Exit Sub
        End Try
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

    Sub CalculaSaldoContigente(ID_Contrato As Integer, Fecha As Date)
        Dim EsquemaCobro As Integer = TaAnexos.EsquemaCobro(ID_Contrato)
        Dim Tope As Date
        If EsquemaCobro = 1 Then 'simple con financiamiento
            Tope = Fecha.AddDays(-999)
        Else
            Tope = Fecha.AddDays(-121)
        End If
        Dim SaldoConti, Interes, CapVig, SaldoCap As Decimal

        TaEdoCta.FillDesc(ds.CONT_CPF_edocuenta, ID_Contrato, "BP")
        SaldoCap = TaEdoCta.SaldoCapital(ID_Contrato, "BP")
        For Each r As PasivoFiraDS.CONT_CPF_edocuentaRow In ds.CONT_CPF_edocuenta.Rows
            If r.fecha_fin >= Tope Then
                Interes += r.int_ord
                CapVig += r.cap_vigente
            Else
                Exit For
            End If
        Next

        SaldoConti = CapVig + Interes + SaldoCap
        taContraGarant.FillByidContrato(ds.CONT_CPF_contratos_garantias, ID_Contrato)
        Dim rz As PasivoFiraDS.CONT_CPF_contratos_garantiasRow
        rz = ds.CONT_CPF_contratos_garantias.Rows(0)

        TaSaldoConti.Insert(Fecha, Nothing, Nothing, 0, 0, Nothing, SaldoConti, rz.cobertura_nominal, rz.cobertura_efectiva,
                               SaldoConti * (rz.cobertura_nominal / 100), SaldoConti * (rz.cobertura_efectiva / 100), rz.id_contrato_garantia)
        taContraGarant.UpdateSaldoContingente(SaldoConti * (rz.cobertura_nominal / 100), rz.id_contrato_garantia, rz.id_contrato)

    End Sub

    '++++++++++++++++++++++++++++++++++++++++++++++++++++
    Sub Procesa_SIMPLE_VAR(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim diasX As Integer
        Dim AjusteDias As Integer = 0
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision, Aux1, Aux2 As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim TipoTasa As String

        AjusteDias = TaEdoCta.EsDiaFestivo(Fecha, "MXN")
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Saturday And (EsVencimetoCap = True Or RevisionTasa = True) Then AjusteDias += 2
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Sunday And (EsVencimetoCap = True Or RevisionTasa = True) Then AjusteDias += 1
        AjusteDias += TaEdoCta.EsDiaFestivo(Fecha.AddDays(AjusteDias), "MXN")

        TipoTasa = "BP"
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)
        CargaTIIE(r.FechaCorte, "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        TIIE_old = TIIE28

        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha.AddDays(AjusteDias))
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital  'no se acumula en simple
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha.AddDays(AjusteDias))

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL vIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha.AddDays(AjusteDias))
            End If
            IntORD = Math.Round((r.Capital) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        End If

        SaldoFIN = SaldoINI
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
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
            Else
                If EsCorteInte = True Then
                    IntORD_Aux = InteresAux1 * -1
                Else
                    IntORD = 0
                End If

            End If
            SaldoFIN = SaldoINI - CapitalVIG + Minis_BASE  ' no acumula en simple

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha.AddDays(AjusteDias), "", "")
            If EsCorteInte = True Then
                TaAnexos.UpdateFechaCorteTIIE(Fecha.AddDays(AjusteDias), TIIE28, ID_Contrato)
            Else

            End If


            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB + InteresAux1FB, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha.AddDays(AjusteDias), ID_Contrato)
                    CalculaServicioCobro(Fecha.AddDays(AjusteDias), SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoFIN = SaldoINI + Minis_BASE ' no acumula en simple
            '++++++APLICA PAGO SOLO INTERES+++++++++++
            Aux1 = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
            If Aux1 > 0 Then
                Aux2 = Aux1 / IntORD
                IntORD -= Aux1
                Aux2 = IntFB * Aux2
                IntFB = 0
                SaldoFIN = SaldoINI
                Provision = IntORD
                IntORD = 0
            Else
                Aux1 = 0
                Aux2 = 0
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++

            TaEdoCta.Insert(TipoTasa, r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0, Provision)

        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha.AddDays(AjusteDias), "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha.AddDays(AjusteDias))
    End Sub

    Sub Procesa_SIMPLE_FIJA(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim AjusteDias As Integer = 0
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision, Aux1, Aux2 As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow
        Dim saldoINIfn As Decimal = 0

        AjusteDias = TaEdoCta.EsDiaFestivo(Fecha, "MXN")
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Saturday And (EsVencimetoCap = True Or RevisionTasa = True) Then AjusteDias += 2
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Sunday And (EsVencimetoCap = True Or RevisionTasa = True) Then AjusteDias += 1
        AjusteDias += TaEdoCta.EsDiaFestivo(Fecha.AddDays(AjusteDias), "MXN")

        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha.AddDays(AjusteDias), "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha.AddDays(AjusteDias))
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha.AddDays(AjusteDias))
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital '+ r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital) * ((TasaActivaBP) / 100 / 360) * (diasX), 2) ' no capitaliza
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If
        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI '+ IntFINAN NO CAUMULA
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN) / 100 / 360) * (diasX)
        End If

        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = 0
            Dim IntFB_Aux As Decimal = 0
            Dim IntFN_Aux As Decimal = 0
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                'CapitalVIG = TaVevcimientosCPF.CapitalVigente(ID_Contrato)
                SaldoFIN -= CapitalVIG '+ IntFINAN
                IntORD_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "BP")
                IntFB_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "FB")
                IntFN_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "FN")
            Else
                'SaldoFIN = SaldoINI + IntORD - CapitalVIG
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha.AddDays(AjusteDias), "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha.AddDays(AjusteDias), TIIE28, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB + InteresAux1FB, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha.AddDays(AjusteDias), SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoFIN = SaldoINI '+ IntORD NO CAPITALIZA
            '++++++APLICA PAGO SOLO INTERES+++++++++++
            Aux1 = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
            If Aux1 > 0 Then
                Aux2 = Aux1 / IntORD
                IntORD -= Aux1
                Aux2 = IntFB * Aux2
                IntFB = 0
                SaldoFIN = SaldoINI
                Provision = IntORD
                IntORD = 0
            Else
                Aux1 = 0
                Aux2 = 0
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++
            TaEdoCta.Insert("BP", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFN), diasX, IntFN, 0, Provision)
            End If
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha.AddDays(AjusteDias))
    End Sub

    Sub Procesa_SIMPLE_CON_FIJA(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim AjusteDias As Integer
        AjusteDias = TaEdoCta.EsDiaFestivo(Fecha, "MXN")
        If Fecha.DayOfWeek = DayOfWeek.Saturday And EsVencimetoCap = False Then AjusteDias += 2
        If Fecha.DayOfWeek = DayOfWeek.Sunday And EsVencimetoCap = False Then AjusteDias += 1
        AjusteDias += TaEdoCta.EsDiaFestivo(Fecha.AddDays(AjusteDias), "MXN")
        If AjusteDias > 0 Then
            Dim t As New PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaDataTable
            taCalendar.FillByIdContrato(t, Fecha.AddDays(AjusteDias), ID_Contrato)
            If t.Rows.Count <= 0 Then
                taCalendar.Insert(ID_Contrato, Fecha.AddDays(AjusteDias), False, AcumulaInteres, EsCorteInte, RevisionTasa, True)
            End If
            EsCorteInte = False
            AcumulaInteres = False
            RevisionTasa = False
        End If


        Dim diasX As Integer
        Dim Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision, Aux1, Aux2 As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha, "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                'CapitalVIG = TaVevcimientosCPF.CapitalVigente(ID_Contrato)
                SaldoFIN = r.Capital - CapitalVIG
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
            Else
                SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            If ID_Contrato = 15 And Minis_BASE = 381600.0 Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                SaldoFIN = 3500000
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++
            Aux1 = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
            If Aux1 > 0 Then
                Aux2 = Aux1 / IntORD
                IntORD -= Aux1
                Aux2 = IntFB * Aux2
                IntFB = 0
                SaldoFIN = SaldoINI
                Provision = IntORD
                IntORD = 0
            Else
                Aux1 = 0
                Aux2 = 0
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN, 0, Provision)
            End If
            '  CargaTIIE(Fecha)
            '   TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

    Sub Procesa_SIMPLE_CON_VAR(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim AjusteDias As Integer
        AjusteDias = TaEdoCta.EsDiaFestivo(Fecha, "MXN")
        If Fecha.DayOfWeek = DayOfWeek.Saturday And EsVencimetoCap = False Then AjusteDias += 2
        If Fecha.DayOfWeek = DayOfWeek.Sunday And EsVencimetoCap = False Then AjusteDias += 1
        AjusteDias += TaEdoCta.EsDiaFestivo(Fecha.AddDays(AjusteDias), "MXN")
        If AjusteDias > 0 Then
            Dim t As New PasivoFiraDS.CONT_CPF_CalendariosRevisionTasaDataTable
            taCalendar.FillByIdContrato(t, Fecha.AddDays(AjusteDias), ID_Contrato)
            If t.Rows.Count <= 0 Then
                taCalendar.Insert(ID_Contrato, Fecha.AddDays(AjusteDias), False, AcumulaInteres, EsCorteInte, RevisionTasa, True)
            End If
            EsCorteInte = False
            AcumulaInteres = False
            RevisionTasa = False
        End If

        Dim diasX As Integer
        Dim TIIE_old, Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision, Aux1, Aux2 As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        'Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow
        Dim saldoINIfn As Decimal = 0

        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TasaActivaFB(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)
        CargaTIIE(r.FechaCorte, "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        If InteresAux1 < 0 Then
            InteresAux1 = 0
        End If
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        TIIE_old = TIIE28
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)

        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        'SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido
        SaldoINI = r.Capital + r.InteresOrdinario

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            'If Provision > 0 And InteresAux1 = 0 Then
            '    'diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            'End If
            IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2) + Provision

            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE > 0 And Tipar = "H" Then
                    Provision = IntORD - Provision
                    IntORD = 0
                ElseIf Minis_BASE = 0 Then
                    Provision = IntORD - Provision
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN + TIIE_old) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                'Try
                CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                If CapitalVIG < 0 Then
                    TaVevcimientosCPF.PorcesarInteresPagado(IntFB, r.id_contrato, Fecha)
                End If
                CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                'Catch ex As Exception
                'Try
                '    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                'Catch ex1 As Exception
                '    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                'End Try
                'End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
                SaldoFIN -= (CapitalVIG + IntORD)
                SaldoFIN += IntORD_Aux
                SaldoFIN += Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            If ID_Contrato = 15 And Minis_BASE = 381600.0 Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                SaldoFIN = 3500000
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            If RevisionTasa = True Then
                CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
                TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)
            End If

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++
            Aux1 = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
            If Aux1 > 0 Then
                Aux2 = Aux1 / IntORD
                IntORD -= Aux1
                Aux2 = IntFB * Aux2
                IntFB = 0
                SaldoFIN = SaldoINI
                Provision = IntORD
                IntORD = 0
            Else
                Aux1 = 0
                Aux2 = 0
            End If
            '++++++APLICA PAGO SOLO INTERES+++++++++++

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, Aux1, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, Aux2, 0, 0, 0, 0,
                                0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                    TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN, 0, Provision)
                End If
            End If
            TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

    Sub Procesa_TRADICIONAL_FIJA(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim AjusteDias As Integer = 0
        Dim diasX, diasY As Integer
        Dim FechaAnt As Date
        Dim Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim Rsaldo As PasivoFiraDS.CONT_CPF_saldos_contingenteRow
        Dim saldoINIfn As Decimal = 0

        AjusteDias = TaEdoCta.EsDiaFestivo(Fecha, "MXN")
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Saturday And EsVencimetoCap = True Then AjusteDias += 2
        If Fecha.AddDays(AjusteDias).DayOfWeek = DayOfWeek.Sunday And EsVencimetoCap = True Then AjusteDias += 1
        AjusteDias += TaEdoCta.EsDiaFestivo(Fecha.AddDays(AjusteDias), "MXN")

        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha.AddDays(AjusteDias), "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha.AddDays(AjusteDias))
        Minis_BASE = TaEdoCta.Minis_Base(r.id_contrato)
        FechaAnt = TaEdoCta.Minis_Base_Fec(r.id_contrato)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha.AddDays(AjusteDias))
        diasY = DateDiff(DateInterval.Day, FechaAnt, Fecha.AddDays(AjusteDias))
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital '+ r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            'If ID_Contrato = 20 And Fecha = CDate("28/03/2018") Then diasX -= 1
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital) * ((TasaActivaBP) / 100 / 360) * (diasX), 2) ' no capitaliza
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If
        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI '+ IntFINAN NO CAUMULA
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN) / 100 / 360) * (diasX)
        End If

        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = 0
            Dim IntFB_Aux As Decimal = 0
            Dim IntFN_Aux As Decimal = 0
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                'CapitalVIG = TaVevcimientosCPF.CapitalVigente(ID_Contrato)
                SaldoFIN -= CapitalVIG '+ IntFINAN
                IntORD_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "BP")
                IntFB_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "FB")
                IntFN_Aux -= TaAnexos.InteresAcumulado(r.id_contrato, "FN")
            Else
                'SaldoFIN = SaldoINI + IntORD - CapitalVIG
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha.AddDays(AjusteDias), "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha.AddDays(AjusteDias), TIIE28, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB + InteresAux1FB, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha.AddDays(AjusteDias), SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha.AddDays(AjusteDias), 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoFIN = SaldoINI '+ IntORD NO CAPITALIZA
            TaEdoCta.Insert("BP", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                        0, 0, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFB), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha.AddDays(AjusteDias), SaldoINI, SaldoINI, 0, 0, 0, 0, 0, 0, 0,
                            0, 0, ID_Contrato, (TasaActivaFN), diasX, IntFN, 0, Provision)
            End If
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha.AddDays(AjusteDias))
    End Sub

    Sub Procesa_SIMFAA_FIJA(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        EsCorteInte = True ' SIEMPRE ES CORTE DE INTERES
        Dim diasX As Integer
        Dim Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha, "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
                SaldoFIN -= (CapitalVIG + IntORD)
                SaldoFIN += IntORD_Aux
                SaldoFIN += Minis_BASE
            Else
                CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            If ID_Contrato = 15 And Minis_BASE = 381600.0 Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                SaldoFIN = 3500000
            ElseIf ID_Contrato = 25 And Fecha = CDate("2017-07-31") Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
            ElseIf ID_Contrato = 52 Then
                If Fecha = CDate("30-JUN-16") Then IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                If Fecha = CDate("30-JUN-17") Then IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
            ElseIf ID_Contrato = 58 Then
                If Fecha = CDate("30-JUL-16") Then
                    IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                End If
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN, 0, Provision)
            End If
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

    Sub Procesa_SIMFAA_VAR(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        Dim diasX As Integer
        Dim Minis_BASE, TIIE_old As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(r.FechaCorte, "", "")
        Select Case r.TasaTiie.Trim
            Case "TIIE28"
                TIIE_old = TIIE28
            Case "TIIE91"
                TIIE_old = TIIE91
            Case "TIIE182"
                TIIE_old = TIIE182
            Case "TIIE365"
                TIIE_old = TIIE365
            Case "TIIE_PROM"
                TIIE_old = TIIE_Promedio
        End Select

        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN + TIIE_old) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
                SaldoFIN -= (CapitalVIG + IntORD)
                SaldoFIN += IntORD_Aux
                SaldoFIN += Minis_BASE
            Else
                CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            If ID_Contrato = 15 And Minis_BASE = 381600.0 Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                SaldoFIN = 3500000
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN_Aux, IntVENC, Provision)
            End If
            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN, 0, Provision)
            End If
            '  CargaTIIE(Fecha)
            '   TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE28, ID_Contrato)
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

    Sub Procesa_COBRO_MENSUAL_VAR(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        EsCorteInte = True ' SIEMPRE ES CORTE DE INTERES
        Dim diasX As Integer
        Dim Minis_BASE, TIIE_old As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(r.FechaCorte, "", "")
        Select Case r.TasaTiie.Trim
            Case "TIIE28"
                TIIE_old = TIIE28
            Case "TIIE91"
                TIIE_old = TIIE91
            Case "TIIE182"
                TIIE_old = TIIE182
            Case "TIIE365"
                TIIE_old = TIIE365
            Case "TIIE_PROM"
                TIIE_old = TIIE_Promedio
        End Select

        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital + r.InteresOrdinario + r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP + TIIE_old) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        SaldoFIN = SaldoINI + IntFINAN
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB + TIIE_old) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN + TIIE_old) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
                SaldoFIN -= (CapitalVIG + IntORD)
                SaldoFIN += IntORD_Aux
                SaldoFIN += Minis_BASE
            Else
                CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            'SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")

            If ID_Contrato = 15 And Minis_BASE = 381600.0 Then
                IntORD_Aux = TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1
                SaldoFIN = 3500000
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP + TIIE_old), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB + TIIE_old), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaFN + TIIE_old), diasX, IntFN, 0, Provision)
            End If
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

    Sub Procesa_COBRO_MENSUAL_FIJA(ByRef Fecha As Date, ByRef ID_Contrato As Integer, ByRef EsCorteInte As Boolean, ByRef r As PasivoFiraDS.SaldosAnexosRow, ByRef EsVencimetoCap As Boolean, AcumulaInteres As Boolean, RevisionTasa As Boolean, tasaFija As Boolean)
        EsCorteInte = True ' SIEMPRE ES CORTE DE INTERES
        Dim diasX As Integer
        Dim Minis_BASE As Decimal
        Dim TasaActivaBP, TasaActivaFB, TasaActivaFN, IntORD, IntVENC, IntFINAN, SaldoINI, SaldoFIN, InteresAux1, InteresAux2 As Decimal
        Dim InteresAux1FB, InteresAux2FB, Provision As Decimal
        Dim InteresAux1FN, InteresAux2FN As Decimal
        Dim IntFB As Decimal = 0
        Dim IntFN As Decimal = 0
        Dim saldoINIfn As Decimal = 0
        TasaActivaBP = TaAnexos.TasaActivaBP(r.id_contrato)
        TasaActivaFB = TaAnexos.TASAFIJA(r.id_contrato)
        TasaActivaFN = TaAnexos.TasaActivaFN(r.id_contrato)

        CargaTIIE(Fecha, "", "")
        InteresAux1 = TaEdoCta.SacaInteresAux1(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2 = TaEdoCta.SacaInteresAux2(r.id_contrato, r.FechaCorte, Fecha)
        Provision = TaEdoCta.Provision(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FB = TaEdoCta.SacaInteresAux1FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FB = TaEdoCta.SacaInteresAux2FB(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux1FN = TaEdoCta.SacaInteresAux1FN(r.id_contrato, r.FechaCorte, Fecha)
        InteresAux2FN = TaEdoCta.SacaInteresAux2FN(r.id_contrato, r.FechaCorte, Fecha)
        Minis_BASE = TaMinis.SacaMontoMinisXfecha(ID_Contrato, Fecha)
        diasX = DateDiff(DateInterval.Day, r.FechaFinal, Fecha)
        subsidio = TaAnexos.Subsidiocontrato(r.id_contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos 
        SaldoINI = r.Capital '+ r.InteresOrdinario + r.Vencido + r.InteresVencido

        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vigente") > 0 Then ' CAPITAL VIGENTE
            If Provision > 0 And InteresAux1 = 0 Then
                diasX += TaEdoCta.ProvisionDias(r.id_contrato, r.FechaFinal, Fecha)
            End If
            'IntORD = Math.Round((r.Capital + r.InteresOrdinario) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
            IntORD = Math.Round((r.Capital) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
            If EsVencimetoCap = False And EsCorteInte = False And RevisionTasa = False And AcumulaInteres = False Then
                If Minis_BASE = 0 Then
                    Provision = IntORD
                    IntORD = 0
                Else
                    Provision = 0
                End If
            Else
                Provision = 0
            End If
        Else
            IntORD = 0
        End If
        If TaVevcimientosCPF.TotalCapitalStatus(r.id_contrato, "Vencido") > 0 Then ' CAPITAL VENCIDO
            IntVENC = Math.Round((r.Vencido + r.InteresVencido) * ((TasaActivaBP) / 100 / 360) * (diasX), 2)
        Else
            IntVENC = 0
        End If

        IntFINAN = IntORD + IntVENC - InteresAux1 + InteresAux2
        IntFB = Math.Round((SaldoINI) * ((TasaActivaFB) / 100 / 360) * (diasX), 2)

        If TasaActivaFN > 0 Then
            IntFN = (SaldoINI) * ((TasaActivaFN) / 100 / 360) * (diasX)
        End If

        If EsCorteInte = True Then
            Dim CapitalVIG As Decimal = 0
            Dim IntORD_Aux As Decimal = IntORD
            Dim IntFB_Aux As Decimal = IntFB
            Dim IntFN_Aux As Decimal = IntFN
            If EsVencimetoCap Then
                Try
                    CapitalVIG = TaVevcimientosCPF.InteresPagado(r.id_contrato, Fecha)
                    If CapitalVIG < 0 Then
                        TaVevcimientosCPF.PorcesarInteresPagado(Math.Abs(IntFB), r.id_contrato, Fecha)
                    End If
                    CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                Catch ex As Exception
                    Try
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-1))
                    Catch ex1 As Exception
                        CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha.AddDays(-2))
                    End Try
                End Try
                IntORD_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "BP") * -1)
                IntFB_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FB") * -1)
                IntFN_Aux = (TaAnexos.InteresAcumulado(r.id_contrato, "FN") * -1)
                'SaldoFIN -= (CapitalVIG + IntORD)
                'SaldoFIN += IntORD_Aux
                'SaldoFIN += Minis_BASE
            Else
                CapitalVIG = TaVevcimientosCPF.CapitalVencimiento(r.id_contrato, Fecha)
                'SaldoFIN = SaldoINI + IntORD - CapitalVIG + Minis_BASE
            End If
            SaldoFIN = SaldoINI - CapitalVIG + Minis_BASE

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntORD + InteresAux1, IntVENC + InteresAux2, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD_Aux, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFB + InteresAux1FB, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB_Aux, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, CapitalVIG, 0, IntFN + InteresAux1FN, IntVENC + InteresAux2FN, 0, 0, 0,
                                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN_Aux, IntVENC, Provision)
            End If

            CargaTIIE(Fecha, "", "") 'NO APLICA EN TASA FIJA   dagl pero si lo estan aplicando para fn 
            TaAnexos.UpdateFechaCorteTIIE(Fecha, TIIE_Promedio, ID_Contrato)

            If EsVencimetoCap Then 'Pago automatico por Vencimiento de Capital
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("AUTOMATICO", "PAGADO", Fecha, 0, 0, 0, IntORD_Aux, 0, CapitalVIG, 0, r.id_contrato)
                If TaVevcimientosCPF.VencimientosXdevengar(ID_Contrato) > 0 Then 'pago de cobro de servicio por garantia
                    TaVevcimientosCPF.UpdateEstatus("Vencido", Fecha, ID_Contrato)
                    CalculaServicioCobro(Fecha, SaldoFIN, r.porcentaje_cxsg, ID_Contrato, subsidio)
                End If
            Else
                Dim Pag As New PasivoFiraDSTableAdapters.PagosTableAdapter
                Pag.Insert("PAGO POR REF", "APLICADO", Fecha, 0, 0, 0, IntFB, 0, 0, 0, r.id_contrato) 'DAGL Ingresar pago de interes 23/01/2018
            End If
        Else
            SaldoINI += TaAnexos.InteresAcumulado(r.id_contrato, "BP")
            If EsCorteInte = False And RevisionTasa = False Then
                SaldoFIN = SaldoINI + Minis_BASE
            Else
                SaldoFIN = SaldoINI + IntORD + Minis_BASE
            End If

            TaEdoCta.Insert("BP", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaBP), diasX, IntORD, IntVENC, Provision)

            TaEdoCta.Insert("FB", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                            0, Minis_BASE, ID_Contrato, (TasaActivaFB), diasX, IntFB, 0, Provision)

            If TasaActivaFN > 0 Then
                TaEdoCta.Insert("FN", r.FechaFinal, Fecha, SaldoINI, SaldoFIN, 0, 0, 0, 0, 0, 0, 0,
                        0, Minis_BASE, ID_Contrato, (TasaActivaFN), diasX, IntFN, 0, Provision)
            End If
        End If
        TaVevcimientosCPF.UpdateStatusALL("Vencido", Fecha, "Vigente", ID_Contrato, 0)
        CalculaSaldoContigente(ID_Contrato, Fecha)
    End Sub

End Module
