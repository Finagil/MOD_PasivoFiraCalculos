Module ModGlobal

    Public TaVencimientosCPF As New PasivoFiraDSTableAdapters.CONT_CPF_vencimientosTableAdapter
    Public TaEdoCta As New PasivoFiraDSTableAdapters.CONT_CPF_edocuentaTableAdapter
    Public TaAnexos As New PasivoFiraDSTableAdapters.SaldosAnexosTableAdapter
    Public TaMinis As New PasivoFiraDSTableAdapters.CONT_CPF_ministracionesTableAdapter
    Public Ministraciones As New DescuentosDSTableAdapters.MinistracionesTableAdapter
    Public TaSaldoConti As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Public MFIRA As New PasivoFiraDSTableAdapters.mFIRATableAdapter
    Public taCalendar As New PasivoFiraDSTableAdapters.CONT_CPF_CalendariosRevisionTasaTableAdapter
    Public taCXSG As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    Public taContraGarant As New PasivoFiraDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
    Public tapagos As New PasivoFiraDSTableAdapters.PagosTableAdapter
    Public taGarantias As New PasivoFiraDSTableAdapters.CONT_CPF_contratos_garantiasTableAdapter
    Public taCargosXservico As New PasivoFiraDSTableAdapters.CONT_CPF_csgTableAdapter
    Public SaldoCont As New PasivoFiraDSTableAdapters.CONT_CPF_saldos_contingenteTableAdapter
    Public taVencimientos As New PagosFiraDSTableAdapters.VencimientosTableAdapter
    Public taCaledarios As New PagosFiraDSTableAdapters.CalendariosTableAdapter
    Public taPagosFira As New PagosFiraDSTableAdapters.CONT_CPF_PagosFiraTableAdapter
    Public taCorreos As New PagosFinagilDSTableAdapters.Correos_SistemaFinagilTableAdapter
    Public taProcContra As New PasivoFiraDSTableAdapters.ContratosProcesarFechaTableAdapter
    Public DS As New PagosFinagilDS


    Public Enum EsquemaCobro As Integer
        SIMFAA = 20
        SIMPLE = 71
        SIMPLE_FIN = 21
        TRADICIONAL = 70
        COBRO_MENSUAL = 30
    End Enum

    Public TasaIVA As Decimal = 0.16
    Public TIIE28 As Decimal = 0
    Public TIIE91 As Decimal = 0
    Public TIIE182 As Decimal = 0
    Public TIIE365 As Decimal = 0
    Public TIIE_Promedio As Decimal = 0
    Public TIIE_Aplica As Decimal = 0
    Public Function CargaTIIE(ByVal Fecha As Date, ByVal Tipta As String, ByVal claveCobro As String) As Boolean
        Try
            CargaTIIE = True
            Dim ta As New PasivoFiraDSTableAdapters.TIIETableAdapter
            ta.Connection.ConnectionString = "Data Source=server-raid2;Persist Security Info=True;Password=User_PRO2015;User ID=User_PRO;Initial Catalog=Production;"
            TIIE28 = ta.SacaTIIE28(Fecha.ToString("yyyyMMdd"))
            TIIE91 = ta.SacaTIIE91(Fecha.ToString("yyyyMMdd"))
            TIIE182 = ta.SacaTIIE182(Fecha.ToString("yyyyMMdd"))
            TIIE365 = ta.SacaTIIE365(Fecha.ToString("yyyyMMdd"))
            TIIE_Promedio = ta.SacaTIIEpromedio(Fecha.AddMonths(-1).ToString("yyyyMMdd"))
            'If TIIE28 = 0 Or TIIE91 = 0 Or TIIE182 = 0 Or TIIE365 = 0 Then
            If TIIE28 = 0 Then
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
        Catch ex As Exception
            Console.WriteLine("Error: ID-" & " " & ex.Message & " " & Date.Now)
            CorreosFases("Error: ", ex.Message, "SISTEMAS_FIRA")
        End Try
    End Function

    Function CtoD(Fec As String) As Date
        Dim f As Date = New DateTime(Fec.Substring(0, 4), Fec.Substring(4, 2), Fec.Substring(6, 2))
        Return f
    End Function

    Sub CorrigeCapitalVencimiento(ID As Integer)
        Dim DS2 As New PagosFiraDS
        Dim rVenc As PagosFiraDS.VencimientosRow
        Dim SaldoCap As Decimal = TaEdoCta.SaldoCapital(ID, "BP")
        taVencimientos.FillByUltimo(DS2.Vencimientos, ID)
        For Each rVenc In DS2.Vencimientos.Rows
            If rVenc.capital <> SaldoCap And SaldoCap > 0 Then
                rVenc.capital = SaldoCap
                DS2.Vencimientos.GetChanges()
                taVencimientos.Update(DS2.Vencimientos)
            End If
        Next
        If SaldoCap < 0 Then
            CorreosFases("Error Saldo Negativo: " & ID, "Error Saldo Negativo: " & ID, "SISTEMAS_FIRA")
        End If
    End Sub

    Public Sub CorreosFases(Titulo As String, Mensaje As String, Fase As String)
        Dim taFases As New PagosFinagilDSTableAdapters.GEN_CorreosFasesTableAdapter
        taFases.Fill(DS.GEN_CorreosFases, Fase)
        For Each r As PagosFinagilDS.GEN_CorreosFasesRow In DS.GEN_CorreosFases.Rows
            taCorreos.Insert("PasivoFira@finagil.com.mx", r.Correo, Titulo, Mensaje, False, Date.Now, "")
        Next
        taFases.Dispose()
    End Sub

End Module
