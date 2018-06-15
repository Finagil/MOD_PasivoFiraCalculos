Module ModuleMain


    Sub Main()
        Dim Args() As String = Environment.GetCommandLineArgs()
        Try
            If Args.Length > 1 Then
                If IsNumeric(Args(1)) Then
                    If Args(1) > 0 Then
                        If Procesa_Pagos_Fira(Args(1)) = 0 Then
                            Console.WriteLine("Proceso Pagos")
                        Else
                            Console.WriteLine("Sin Pagos")
                        End If
                        ProcesaEstadoCuenta(Args(1), False, Today.Date)
                    Else
                        Console.WriteLine("ID incorrecto")
                    End If
                ElseIf Args(1).ToUpper.Trim = "PAGOS" Then
                    ExportaPagosFinagilFira()
                ElseIf Args(1).ToUpper.Trim = "PROCESA_PAGOS" Then
                    Procesa_Pagos_Fira(0)
                ElseIf Args(1).ToUpper.Trim = "TODO" Then
                    Dim Tabla As New PasivoFiraDS.SaldosAnexosDataTable
                    TaAnexos.Fill_ConSaldo(Tabla)
                    For Each x As PasivoFiraDS.SaldosAnexosRow In Tabla.Rows
                        If Procesa_Pagos_Fira(x.id_contrato) = 0 Then
                            ProcesaEstadoCuenta(x.id_contrato, True, Today.Date)
                        End If
                    Next
                ElseIf Args(1).ToUpper.Trim = "PROCESA_FECHA" Then
                    Dim Fecha As Date = Today.Date
                    If Args.Length = 3 Then
                        Fecha = CDate(Args(2))
                    End If
                    Dim Tabla1 As New PasivoFiraDS.ContratosProcesarFechaDataTable
                    taProcContra.Fill(Tabla1, Fecha.Date)
                    For Each y As PasivoFiraDS.ContratosProcesarFechaRow In Tabla1.Rows
                        ProcesaEstadoCuenta(y.Id_Contrato, True, Fecha.Date)
                    Next
                Else
                    Console.WriteLine("Opcion Inválida")
                End If
            Else
                Console.WriteLine("Opcion Inválida")
            End If
        Catch ex As Exception
            Console.WriteLine("Error: ID-" & Args(1) & " " & ex.Message & " " & Date.Now)
            taCorreos.Insert("PasivoFira@finagil.com.mx", "ecacerest@finagil.com.mx", "Error: " & Args(1), ex.Message, False, Date.Now, "")
        End Try

    End Sub

End Module
