Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "GBP"
        D.Login = "AR4041"
        D.SSO = True
        D.Password = "hmetal27"

        Dim SC As New SAPConnector
        Dim con = SC.GetSAPConnection(D)

        Dim R As New SAPExchgRate(con)
        R.ValidFrom = "07/01/2014"
        R.Execute()

        'Dim D As New ConnectionData
        'D.Box = "GBA"
        'D.Login = "BV6358"
        'D.SSO = False
        'D.Password = "PG00005"

        'Dim SC As New SAPConnector
        'Dim Con = SC.GetSAPConnection(D)

        'Dim POC As New POCreator(Con)
        'POC.CreateNewPO("NB")
        'POC.Vendor = "10014823"
        'POC.PurchOrg = "1459"
        'POC.PurchGroup = "DXP"
        'POC.CompanyCode = "002"
        'POC.AppendHeaderText(SAPCOM.SAPTextIDs.HeaderText, "This order replaces TEVA PO 42134506.  The current Terms and Conditions between Teva and your company will govern this PO.")
        'POC.AppendHeaderText(SAPCOM.SAPTextIDs.HeaderNote, "Cut Over TEVA PO")
        'POC.CreateBlankItem("10")
        'POC.Plant = "1725"
        'POC.ItemQuantity = "3377"
        'POC.DeliveryDate = "12/31/2014"
        'POC.Currency("1", "1") = "USD"
        'POC.ItemShortText = "USP Start-up Services"
        'POC.GL_Account = "52970009"
        'POC.AccAsignmentCat = "P"
        'POC.MatGroup = "G41100000"
        'POC.WBS_Element = "N.03976.1719.4.80.92"
        'POC.ItemUOM = "ACT"
        'POC.ChangeTaxCode("3A")
        'POC.PO_Incoterm = "FOB"
        'POC.PO_Incoterm_Desc = "Free On Board"
        'POC.GenerateOutput = False
        'POC.CommitChanges()

        '4501191601

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