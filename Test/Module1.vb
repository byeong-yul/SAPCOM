Imports SAPCOM
Imports ERPConnect

Module Module1

    Sub Main()

        Dim C As New ConnectionData
        C.Box = "GBP"
        C.Login = "AR4041"
        C.SSO = True
        C.Password = "hmetal27"
        Dim SC As New SAPConnector
        Dim con = SC.GetSAPConnection(C)

        Dim E As New EKKO_Report(con)
        E.IncludeDocument("4501213524")
        E.Execute()

        'Dim R As New EKPO_Report(con)
        'R.IncludeDocument("4503401042")
        'R.Execute()

        'Dim BAPI = con.CreateBapi("PurchSchedAgreement", "Change")
        'BAPI.Exports("PurchasingDocument").ParamValue = "5500009030"

        'Dim TRow = BAPI.Tables("Item_Condition").AddRow
        'TRow("ITEM_NO") = "00230"
        'TRow("SERIAL_ID") = "2044793331"
        'TRow("COND_COUNT") = "01"
        'TRow("DELETION_IND") = " "
        'TRow("COND_TYPE") = "PB00"
        'TRow("SCALE_TYPE") = "A"
        'TRow("SCALE_BASE_TY") = ""
        'TRow("SCALE_UNIT") = ""
        'TRow("SCALE_UNIT_ISO") = ""
        'TRow("SCALE_CURR") = ""
        'TRow("SCALE_CURR_ISO") = ""
        'TRow("CALCTYPCON") = "C"
        'TRow("COND_VALUE") = "10"
        'TRow("CURRENCY") = "USD"
        'TRow("CURRENCY_ISO") = "USD"
        'TRow("COND_P_UNT") = "1000"
        'TRow("COND_UNIT") = "EA"
        'TRow("COND_UNIT_ISO") = "EA"
        'TRow("NUMERATOR") = "1"
        'TRow("DENOMINATOR") = "1"
        'TRow("BASE_UOM") = "EA"
        'TRow("BASE_UOM_ISO") = "EA"
        'TRow("LOWERLIMIT") = "0.00"
        'TRow("UPPERLIMIT") = "0.00"
        'TRow("VENDOR_NO") = ""
        'TRow("CHANGE_ID") = ""

        'Dim TRowX = BAPI.Tables("Item_ConditionX").AddRow
        'TRowX("COND_VALUE") = "X"
        'TRowX("DELETION_IND") = "X"
        'TRowX("CURRENCY") = "X"
        'TRowX("COND_P_UNT") = "X"
        'TRowX("ITEM_NO") = "00230"

        'TRow = BAPI.Tables("Item_Condition").AddRow
        'TRow("ITEM_NO") = "00230"
        'TRow("SERIAL_ID") = "2044793331"
        'TRow("COND_COUNT") = "02"
        'TRow("COND_TYPE") = "ZHC3"
        'TRow("COND_VALUE") = "8"
        'TRow("CURRENCY") = "USD"
        'TRow("COND_UNIT") = "EA"
        'TRow("COND_P_UNT") = "1000"
        'TRow("NUMERATOR") = "1"
        'TRow("DENOMINATOR") = "1"

        'TRowX = BAPI.Tables("Item_ConditionX").AddRow
        'TRowX("ITEM_NO") = "00230"
        ''TRowX("SERIAL_ID") = "X"
        ''TRowX("COND_COUNT") = "X"
        ''TRowX("COND_TYPE") = "X"
        ''TRowX("COND_VALUE") = "X"
        ''TRowX("CURRENCY") = "X"
        ''TRowX("COND_UNIT") = "X"
        ''TRowX("COND_P_UNT") = "X"
        ''TRowX("NUMERATOR") = "X"
        ''TRowX("DENOMINATOR") = "X"


        'TRow = BAPI.Tables("Item_Cond_Validity").AddRow
        'TRow("ITEM_NO") = "00230"
        'TRow("SERIAL_ID") = "2044793331"
        'TRow("PLANT") = ""
        'TRow("VALID_FROM") = "20140101"
        'TRow("VALID_TO") = "99991231"

        'TRowX = BAPI.Tables("Item_Cond_ValidityX").AddRow
        'TRowX("VALID_FROM") = "X"
        'TRowX("VALID_TO") = "X"
        'TRowX("ITEM_NO") = "00230"
        'TRowX("SERIAL_ID") = "X"
        'TRowX("PLANT") = "X"

        'BAPI.Execute()
        'BAPI.CommitWork(True)

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