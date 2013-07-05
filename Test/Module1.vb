Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim S As String = ConfControlKeys.Confirmations

        Dim D As New ConnectionData
        D.Box = "N6A"
        D.Login = "AR4041"
        D.SSO = True
        'D.Password = "hmetal23"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim PO As New POChanges(Con, "4503442320")
        PO.BlockInd("10") = True
        PO.CommitChanges()

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)