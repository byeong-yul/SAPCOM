Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "F6P"
        D.Login = "AR4041"
        D.SSO = True
        'D.Password = "114137"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim E As New EKKO_Report(Con)
        E.IncludeDocument("5500062561")
        E.Execute()

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)