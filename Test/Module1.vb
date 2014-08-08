Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "N6P"
        D.Login = "AR4041"
        D.SSO = True
        D.Password = "hmetal27"

        Dim SC As New SAPConnector
        Dim con = SC.GetSAPConnection(D)

        Dim R As New EKET_Report(con)
        R.IncludeDocument("4504350395")
        R.Execute()


    End Sub


    Friend Class Read_SAP_Logon

        Public Function Get_SAPLogonIni(Optional Path As String = "C:\Windows\saplogon_normal.ini") As DataTable

            Get_SAPLogonIni = Nothing
            If Not My.Computer.FileSystem.FileExists(Path) Then
                Exit Function
            End If

            Dim sr As New System.IO.StreamReader(Path)
            Dim line As String
            Dim category As String = Nothing
            Dim Idx As Integer
            Dim dt_SAPLogonIni As New DataTable

            line = sr.ReadLine
            Do

                If line.StartsWith("[") Then
                    category = line.Replace("[", String.Empty).Replace("]", String.Empty)
                    dt_SAPLogonIni.Columns.Add(category, Type.GetType("System.String"))
                End If
                If line.Contains("Item") Then
                    Idx = CInt(Left(line, line.IndexOf("=")).Replace("Item", String.Empty))
                    If dt_SAPLogonIni.Rows.Count < Idx Then
                        dt_SAPLogonIni.Rows.Add(New Object() {line.Substring(line.IndexOf("=") + 1, line.Length - (line.IndexOf("=") + 1))})
                    Else
                        dt_SAPLogonIni.Rows(Idx - 1)(category) = line.Substring(line.IndexOf("=") + 1, line.Length - (line.IndexOf("=") + 1))
                    End If
                End If

                line = sr.ReadLine
            Loop Until line Is Nothing

            Get_SAPLogonIni = dt_SAPLogonIni

        End Function

    End Class

End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)