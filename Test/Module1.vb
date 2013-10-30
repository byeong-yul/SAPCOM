Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim D As New ConnectionData
        D.Box = "L7P"
        D.Login = "AR4041"
        D.SSO = True
        'D.Password = "hmetal25"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim OA As New OAChanges("L7P", "AR4041", Nothing, "4600005655")

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)