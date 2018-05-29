﻿Module ModuleMain


    Sub Main()
        Dim Args() As String = Environment.GetCommandLineArgs()
        Try
            If Args.Length > 1 Then
                If IsNumeric(Args(1)) Then
                    If Args(1) > 0 Then
                        If Procesa_Pagos_Fira(Args(1)) = 0 Then
                            ProcesaEstadoCuenta(Args(1))
                        End If
                    Else
                        ExportaPagosFinagilFira()
                        Procesa_Pagos_Fira(0)
                    End If
                ElseIf Args(1) = "TODO" Then
                    Dim Tabla As New PasivoFiraDS.SaldosAnexosDataTable
                    TaAnexos.Fill_ConSaldo(Tabla)
                    For Each x As PasivoFiraDS.SaldosAnexosRow In Tabla.Rows
                        If Procesa_Pagos_Fira(x.id_contrato) = 0 Then
                            ProcesaEstadoCuenta(x.id_contrato)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            taCorreos.Insert("PasivoFira@finagil.com.mx", "ecacerest@finagil.com.mx", "Error: " & Args(1), ex.Message, False, Date.Now, "")
        End Try

    End Sub

End Module
