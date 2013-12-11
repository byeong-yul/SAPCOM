Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "L6P"
        D.Login = "AR4041"
        D.SSO = True

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim R As New LFB1_Report(Con)
        R.Include_CCode("811")
        R.IncludeVendor("15081993")
        R.Execute()

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)