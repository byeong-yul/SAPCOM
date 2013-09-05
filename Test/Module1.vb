﻿Imports SAPCOM
Imports Common_Functions

Module Module1

    Sub Main()

        Dim S As String = ConfControlKeys.Confirmations

        Dim D As New ConnectionData
        D.Box = "GBP"
        D.Login = "AR4041"
        D.SSO = True
        D.Password = "hmetal25"

        Dim SC As New SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim MAKT As New SAPCOM.MAKT_Report(Con)
        'MARA.AddCustomField("MAKTG")


        MAKT.IncludeMaterial("10058703")

        MAKT.Execute()


    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)