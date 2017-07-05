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

    Public Function CargaTIIE(ByVal Fecha As Date) As Boolean
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
        ta.Dispose()
    End Function

End Module
