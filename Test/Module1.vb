Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "GEP"
        D.Login = "BV7795"
        D.SSO = False
        D.Password = "114137"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim R As New ZBBP_SC_Data_Report(Con)
        R.Include_TransNo("113942819")
        R.Execute()

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)