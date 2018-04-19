Module ModGlobal
    Public Enum EsquemaCobro As Integer
        SIMFA = 20
        SIMPLE = 71
        SIMPLE_FIN = 21
    End Enum

    Public TasaIVA As Decimal = 0.16
    Public TIIE28 As Decimal = 0
    Public TIIE91 As Decimal = 0
    Public TIIE182 As Decimal = 0
    Public TIIE365 As Decimal = 0
    Public TIIE_Promedio As Decimal = 0
    Public TIIE_Aplica As Decimal = 0
    Public Function CargaTIIE(ByVal Fecha As Date, ByVal Tipta As String, ByVal claveCobro As String) As Boolean
        CargaTIIE = True
        Dim ta As New PasivoFiraDSTableAdapters.TIIETableAdapter
        ta.Connection.ConnectionString = "Data Source=server-raid;Persist Security Info=True;Password=User_PRO2015;User ID=User_PRO"
        TIIE28 = ta.SacaTIIE28(Fecha.ToString("yyyyMMdd"))
        TIIE91 = ta.SacaTIIE91(Fecha.ToString("yyyyMMdd"))
        TIIE182 = ta.SacaTIIE182(Fecha.ToString("yyyyMMdd"))
        TIIE365 = ta.SacaTIIE365(Fecha.ToString("yyyyMMdd"))
        TIIE_Promedio = ta.SacaTIIEpromedio(Fecha.AddMonths(-1).ToString("yyyyMMdd"))
        'If TIIE28 = 0 Or TIIE91 = 0 Or TIIE182 = 0 Or TIIE365 = 0 Then
        If TIIE28 = 0 Or TIIE91 = 0 Then
            Console.WriteLine("No hay TIIE Capturada para la Fecha {0}", Fecha.ToShortDateString)
            CargaTIIE = False
        End If
        If claveCobro = "" Then
            claveCobro = 0

        End If
        If CInt(claveCobro.Trim) = EsquemaCobro.SIMPLE_FIN And Tipta.Trim <> "7" Then
            TIIE_Aplica = TIIE28
        End If

        If CInt(claveCobro.Trim) = EsquemaCobro.SIMPLE And Tipta.Trim = "7" Then 'SIMPLE Y FIJA TRAER LA TIIE28 DAGL
            TIIE_Aplica = TIIE28
        End If
        ta.Dispose()

        ta.Dispose()
    End Function

    Function CtoD(Fec As String) As Date
        Dim f As Date = New DateTime(Fec.Substring(0, 4), Fec.Substring(4, 2), Fec.Substring(6, 2))
        Return f
    End Function


    'Private Sub InsertaMinistracion(ID As Integer, ByRef rx As PasivoFiraDS.mFIRARow, ByRef hoy As Date, ByRef MontoBase As Decimal, Anexo As String, Ciclo As String)
    '    Dim CONT_CPF_ministracionesTableAdapter As New PasivoFiraDSTableAdapters.CONT_CPF_ministracionesTableAdapter
    '    Dim MinistracionesTableAdapter As New PasivoFiraDSTableAdapters.MinistracionesTableAdapter
    '    Dim tasafira As Decimal = 0
    '    'Dim Consec As Integer = Ministraciones.SacaConsecutivo(ID)
    '    Dim taGarantias As New PasivoFiraDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
    '    Dim taEdoCta As New PasivoFiraDSTableAdapters.CONT_CPF_edocuentaTableAdapter
    '    Dim taCargosXservico As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    '    Dim SaldoCont As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    '    Dim NoGarantias As Integer = taGarantias.ExistenGarantias(ID)
    '    Dim SaldoINI, SaldoFIN, InteORD, InteORDFN, InteORDFB As Decimal
    '    Dim FechaUltimoMov As Date

    '    CalculaServicioCobro(hoy, ID)
    '    SaldoINI = taEdoCta.SaldoContrato(ID)
    '    SaldoFIN = SaldoINI + MontoBase

    '    CONT_CPF_ministracionesTableAdapter.InsertQueryMinistracion(MontoBase, hoy, rx.Ministracion, PCXSG_Aux, TXT_IVA.Text, txt_importe.Text, ID, "Solicitado", dt_descuento.Text)
    '    MinistracionesTableAdapter.Descontar(Anexo, Ciclo, hoy.ToString("yyyyMMdd"))

    '    If NoGarantias = 0 Then
    '        taGarantias.Insert(ID, ID_garantina, Nominal, MontoBase * (Nominal / 100), Efectiva, True)
    '        CONT_CPF_contratosTableAdapter.Updatesubsidio(CheckBox1.Checked, ID_Contrato)


    '    Else
    '        Nominal = MinistracionesBindingSource.Current("Cobertura_Nominal")
    '        Efectiva = MinistracionesBindingSource.Current("Cobertura_Efectiva")
    '        taGarantias.UpdateSaldoConti(SaldoFIN * (Efectiva / 100), ID)
    '        FB = MinistracionesBindingSource.Current("FB")
    '        BP = MinistracionesBindingSource.Current("BP")
    '        FN = MinistracionesBindingSource.Current("FN")
    '    End If

    '    If SaldoINI > 0 Then
    '        CargaTIIE(MinistracionesBindingSource.Current("FechaCorte"), MinistracionesBindingSource.Current("Tipta"), MinistracionesBindingSource.Current("ClaveEsquema"))

    '        FechaUltimoMov = MinistracionesTableAdapter.FechaUltimoMov(ID)
    '        Dim DiasX As Integer = DateDiff(DateInterval.Day, FechaUltimoMov, dt_descuento.Value)
    '        InteORD = SaldoINI * ((BP + TIIE_Aplica) / 100 / 360) * DiasX
    '        If MinistracionesBindingSource.Current("Tipta") = "7" Then 'DAGL 25/01/2018 En tasa fija se resta el valor FB
    '            InteORDFB = SaldoINI * ((FB + tasafira) / 100 / 360) * DiasX
    '        Else
    '            InteORDFB = SaldoINI * ((FB + TIIE_Aplica) / 100 / 360) * DiasX
    '        End If

    '        InteORDFN = SaldoINI * ((FN + TIIE_Aplica) / 100 / 360) * DiasX
    '    Else
    '        CargaTIIE(dt_descuento.Value, MinistracionesBindingSource.Current("Tipta"), MinistracionesBindingSource.Current("ClaveEsquema"))
    '        '  If MinistracionesBindingSource.Current("Tipta") = "7" Then TIIE_Aplica = TxttasaFira.Text - FB 'DAGL 25/01/2018 En tasa fija se resta el valor FB
    '        FechaUltimoMov = dt_descuento.Value.ToShortDateString
    '    End If

    '    id_contratoGarantia = Me.MinistracionesTableAdapter.SacaID(ID)
    '    Dim subsidioaux As Boolean
    '    subsidioaux = Me.CONT_CPF_contratosTableAdapter.subsidio_contrato(ID_Contrato) 'dagl 24/01/2018 se agrega subsidio de la tabla contratos
    '    taCargosXservico.Insert(dt_descuento.Value, FechaFinal, dias, Date.Now, MontoBase, Cobro, Cobro * TasaIVA, Cobro * (1 + TasaIVA), txt_porcentaje.Text, id_contratoGarantia, subsidioaux)
    '    SaldoCont.Insert(dt_descuento.Value, Nothing, Nothing, Nothing, Nothing, MontoBase, SaldoFIN, Nominal, Efectiva, SaldoFIN * (Nominal / 100), SaldoFIN * (Efectiva / 100), id_contratoGarantia)


    '    If MinistracionesBindingSource.Current("FN") > 0 Then
    '        'If MinistracionesBindingSource.Current("Tipta") = "7" Then
    '        '  taEdoCta.Insert("FN", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, id, TIIE_Aplica, 0, InteORDFN, 0)
    '        'Else
    '        taEdoCta.Insert("FN", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, ID, FN + TIIE_Aplica, 0, InteORDFN, 0)
    '        'End If

    '    End If

    '    If MinistracionesBindingSource.Current("Tipta") = "7" Then
    '        taEdoCta.Insert("BP", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, ID, BP, 0, InteORD, 0)
    '    Else
    '        taEdoCta.Insert("BP", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, ID, BP + TIIE_Aplica, 0, InteORD, 0)
    '    End If

    '    ' Dim TIIEfirafija As Decimal = 0
    '    '  TIIE_APLICA_TABLA = Me.TIIETableAdapter.SacaTIIE28(dt_descuento.Value.ToShortDateString)
    '    If MinistracionesBindingSource.Current("Tipta") = "7" Then
    '        tasafira = TxttasaFira.Text 'DAGL 25/01/2018 En tasa fija se resta el valor FB

    '        taEdoCta.Insert("FB", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, ID, tasafira, 0, InteORDFB, 0) 'tasafijafira
    '    Else
    '        taEdoCta.Insert("FB", FechaUltimoMov, dt_descuento.Value.ToShortDateString, SaldoINI, SaldoFIN, 0, Nothing, 0, 0, 0, 0, 0, 0, MontoBase, ID, FB + TIIE_Aplica, 0, InteORDFB, 0)

    '    End If


    '    If ID_Contrato > 0 Then


    '        Me.MinistracionesTableAdapter.UpdateFechaCorteTIIE(dt_descuento.Value.ToShortDateString, TIIE_Aplica, ID_Contrato)
    '        Me.CONT_CPF_contratosTableAdapter.updatetasafijafira(tasafira, ID_Contrato) 'INGRESAMOS VALOR DE LA TASA FIRA FIJA
    '        CreaCalendarioRevisoinTasa(ID_Contrato)
    '        Me.DialogResult = Windows.Forms.DialogResult.OK
    '    End If

    '    CargaDatosDS()

    'End Sub

End Module
