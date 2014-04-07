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

        Dim E As New EORD_Report(Con)
        E.IncludeOA("5500062561")
        E.Execute()

        'E.Data.Columns.Remove("Number")
        E.Data.Columns.Remove("Created On")
        E.Data.Columns.Remove("Created By")
        E.Data.Columns.Remove("Vendor")
        E.Data.Columns.Remove("Fixed Vendor")
        E.Data.Columns.Remove("Procurement Plant")
        E.Data.Columns.Remove("Fixed Issuing Plant")
        E.Data.Columns.Remove("MPN Material")
        E.Data.Columns.Remove("Purch Org")
        E.Data.Columns.Remove("Doc Category")
        E.Data.Columns.Remove("Control Ind")
        E.Data.Columns.Remove("UOM")
        E.Data.Columns.Remove("Logical System")
        E.Data.Columns.Remove("Special Stock")
        E.Data.Columns.Remove("Central Contract")
        E.Data.Columns.Remove("Central Contract Item")
        E.Data.Columns("Valid From").ColumnName = "VFrom"
        E.Data.Columns("Valid To").ColumnName = "VTo"
        E.Data.Columns("Agreement Item").ColumnName = "Item"
        E.Data.Columns("Materials Planning").ColumnName = "MRP"
        E.Data.Columns("Fixed Agreement Item").ColumnName = "Fixed"

        Dim QBase = From R In E.Data Group R By Client = R("Client"), OA = R("Agreement"), Item = R("Item") Into G = Group Select Client, OA, Item, Num = G.Max(Function(R) (R("Number")))

        Dim Q = From QB In QBase Join R In E.Data On QB.Client Equals R("Client") And QB.OA Equals R("Agreement") And QB.Item Equals R("Item") And QB.Num Equals R("Number")
              Select New With { _
                  .Client = QB.Client,
                  .Material = R("Material"),
                  .Plant = R("Plant"),
                  .VFrom = R("VFrom"),
                  .VTo = R("VTo"),
                  .Agreement = QB.OA,
                  .Item = QB.Item,
                  .Fixed = R("Fixed"),
                  .Blocked = R("Blocked"),
                  .MRP = R("MRP")
              }

        Dim X As New MyFunctions_Class
        Dim Dt As DataTable = X.LinQToDataTable(Q)

    End Sub


End Module


'Dim BAPI As BusinessObjectMethod = Con.CreateBapi("IncomingInvoice", "ReleaseSingle")
'BAPI.Exports("InvoiceDocNumber").ParamValue = "5132834592"
'BAPI.Exports("FiscalYear").ParamValue = "2011"
'BAPI.Execute()
'BAPI.CommitWork(True)