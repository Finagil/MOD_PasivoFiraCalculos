Module ModuleMain


    Sub Main()
        Dim Args() As String = Environment.GetCommandLineArgs()
        If Args.Length > 1 Then
            If Args(1) > 0 Then
                ProcesaEstadoCuenta(Args(1))
            Else
                ExportaPagosFinagilFira()
                Procesa_Pagos_Fira()
            End If
        End If
    End Sub

End Module
