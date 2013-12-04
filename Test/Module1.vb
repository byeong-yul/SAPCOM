Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "F7P"
        D.Login = "BJ7774"
        D.SSO = false
        D.Password = "t0d4y55"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim R As New ZFIX_T03_Report(Con)
        R.IncludeCustomParam("REFNR", "18149964")
        R.AddCustomField("LEDGERTEXT")
        R.AddCustomField("WERKS")
        R.Execute()

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)