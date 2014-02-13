Imports ERPConnect
Imports ERPConnect.Queries
Imports ERPConnect.Utils
Imports ERPConnect.ConversionUtils
Imports System.IO
Imports System.Security.Cryptography

#Region "SAPCOM Public Classes"

#Region "Enumerators"

Public NotInheritable Class SAPTextIDs

    Public Const HeaderText = "F01"
    Public Const HeaderNote = "F02"
    Public Const TermsOfDelivery = "F05"
    Public Const ContractRiders = "F11"
    Public Const ItemText = "F01"
    Public Const InforecordPOText = "F02"
    Public Const TermsOfPayment = "F07"

    Public Const PRItemText = "B01"

    Public Const OAReleaseOrderText = "K00"
    Public Const OAHeaderText = "K01"
    Public Const OAHeaderNote = "K02"

End Class

Public NotInheritable Class ConfControlKeys

    Public Const Confirmations As String = "0001"
    Public Const Rough_GR As String = "0002"
    Public Const Shipg_not_rough_GR As String = "0003"
    Public Const Shipping_notificat As String = "0004"
    Public Const Shipment_Tracking As String = "9000"
    Public Const Trans_Tracking As String = "9001"
    Public Const Order_Ack_and_ASN As String = "9002"
    Public Const Lead_Time_Tracking As String = "9003"
    Public Const CCK_for_TAR As String = "9012"
    Public Const Shipping_Milestone As String = "9013"
    Public Const Ship_Mstone_No_ASN As String = "9014"
    Public Const PO_Viewed As String = "9020"
    Public Const PO_Viewed_Ack As String = "9021"

End Class

Public Enum RepairsLevels As Byte

    DoNotProcess = 0
    ExcludeRepairs = 1
    IncludeRepairs = 2
    OnlyRepairs = 3

End Enum

Public Enum ProcStatus
    Not_Processed = 0
    Successfully_Processed = 1
    Incorrectly_Processed = 2
End Enum

#End Region

#Region "SAP Connectivity"

Public NotInheritable Class ConnectionData

    Public Box As String
    Public Login As String
    Public SSO As Boolean
    Public Password As String

    Public Sub New()
    End Sub

    Public Sub New(aBox As String, aLogin As String, Optional aSSO As Boolean = True, Optional aPassword As String = Nothing)

        Box = aBox
        Login = aLogin
        SSO = aSSO
        Password = aPassword

    End Sub

End Class

Public NotInheritable Class SAPConnector

    Private SM As String = Nothing
    Private Servers As DataTable = Nothing

    Private Structure HostData

        Public Validate As Boolean
        Public Host As String
        Public Number As Int16
        Public Client As String
        Public Group As String
        Public SNC_PN As String
        Public Server As String
        Public Balancing As Boolean

    End Structure

    Public Sub New()

        Servers = New DataTable
        Servers.ReadXml(New IO.StringReader(My.Resources.Servers))

    End Sub

    Public ReadOnly Property Status() As String

        Get
            Status = SM
        End Get

    End Property

    Public ReadOnly Property BoxList() As Object()

        Get
            BoxList = Nothing
            If Not Servers Is Nothing Then
                Dim Q = From Server In Servers Select Server("Box")
                BoxList = Q.ToArray
            End If
        End Get

    End Property

    Public Function SNC_ConnString(ByVal Box As String, ByVal User As String) As String

        SNC_ConnString = Nothing
        Dim HD As HostData = GetHostData(Box)
        If HD.Validate AndAlso HD.SNC_PN <> "" Then
            If Box = "L7P" Or Box = "A7P" Then
                SNC_ConnString = "ASHOST=" & HD.Server & " sysnr=" & HD.Number & " client=" & HD.Client & " group=" & HD.Group & " user=" & User & " snc_mode=1 snc_partnername=" & HD.SNC_PN
            Else
                SNC_ConnString = "r3name=" & HD.Host & " sysnr=" & HD.Number & " client=" & HD.Client & " group=" & HD.Group & " user=" & User & " snc_mode=1 snc_partnername=" & HD.SNC_PN
            End If
        End If

    End Function

    Public Function TestConnection(ByVal D As ConnectionData) As Boolean

        TestConnection = False
        If D.Password = "" And Not D.SSO Then
            MsgBox("You must provide the password when SSO is not been used", MsgBoxStyle.Exclamation, "SAP Connection Test")
            Exit Function
        End If
        Dim Con = GetSAPConnection(D)
        If Not Con Is Nothing Then
            TestConnection = True
            Con.Close()
        End If

    End Function

    Public Function GetSAPConnection(ByVal D As ConnectionData) As Object

        GetSAPConnection = GetConnection(D.Box, D.Login, D.SSO, D.Password)

    End Function

    Public Function GetSAPConnection(ByVal Box As String, ByVal User As String, ByVal App As String) As Object

        Dim D As ConnectionData = GetConnectionData(Box, User, App)
        GetSAPConnection = GetConnection(D.Box, D.Login, D.SSO, D.Password)

    End Function

    Public Function GetConnectionData(ByVal Box As String, ByVal User As String, ByVal App As String) As ConnectionData

        Dim D As New ConnectionData
        Dim E As New Simple3Des
        Dim RV

        D.Box = Box
        D.Login = User
        D.SSO = True
        D.Password = Nothing

        RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\L", Box, User)
        If Not RV Is Nothing Then
            D.Login = RV
        End If
        RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\SSO", Box, True)
        If Not RV Is Nothing Then
            D.SSO = RV
        End If
        RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\P", Box, Nothing)
        If Not RV Is Nothing Then
            D.Password = E.DecryptData(RV)
        End If

        GetConnectionData = D

    End Function

    Public Sub SaveConnectionData(ByVal App As String, ByVal D As ConnectionData)

        Dim E As New Simple3Des
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\L", D.Box, D.Login)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\SSO", D.Box, D.SSO)
        Dim P As String = E.EncryptData(D.Password)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\SAPCOM\" & App & "\P", D.Box, P)

    End Sub

    Private Function GetConnection(ByVal Box As String, ByVal User As String, ByVal SSO As Boolean, Optional ByVal Password As String = Nothing) As Object

        GetConnection = Nothing

        If Password Is Nothing And Not SSO Then
            SM = "(E) Password not provided for non SSO login"
            Exit Function
        End If

        ERPConnect.LIC.SetLic(ERPConnect_LicNo)

        Dim C As New R3Connection

        If SSO Then
            Dim CS As String = SNC_ConnString(Box, User)
            If Not CS Is Nothing Then
                Try
                    C.SkipGetSystemInfo = True
                    C.Open(CS)
                Catch ex As Exception
                    SM = "(E) " & ex.Message
                End Try
            End If
        Else
            Dim HD As HostData = GetHostData(Box)
            If HD.Validate Then
                C.Host = HD.Host
                C.SID = HD.Host
                C.SystemNumber = HD.Number
                C.Client = HD.Client
                C.MessageServer = HD.Server
                C.LogonGroup = HD.Group
                C.UserName = User
                C.Password = Password
                C.Language = "EN"
                Try
                    C.Open(HD.Balancing)
                Catch ex As Exception
                    SM = "(E) " & ex.Message
                End Try
            End If
        End If

        If C.Ping Then
            GetConnection = C
        End If

    End Function

    Private Function GetHostData(ByVal Box As String) As HostData

        Dim HD As New HostData
        HD.Validate = False
        If Not Servers Is Nothing Then
            Dim DR As DataRow() = Servers.Select("Box = '" & Box & "'")
            If DR.Length > 0 Then
                HD.Validate = True
                HD.Host = DR(0)("Host")
                HD.Number = DR(0)("Number")
                HD.Client = DR(0)("Client")
                HD.Server = DR(0)("MessageServer")
                HD.Group = DR(0)("LogonGroup")
                HD.Balancing = DR(0)("UseLoadBalancing")
                HD.SNC_PN = DR(0)("SNC_PN")
            Else
                SM = "No Host Data available for SAP System " & Box
            End If
        End If
        GetHostData = HD

    End Function

End Class

#End Region

#Region "Document Information (BAPI)"

Public MustInherit Class SC_BAPI_Base

    Friend BAPI As BusinessObjectMethod = Nothing
    Friend R() As String = Nothing
    Friend Errors As Boolean = False
    Friend IE As Boolean = False
    Friend Con As R3Connection = Nothing
    Friend RF As Boolean = False
    Friend Sts As String = Nothing
    Friend DN As String = Nothing
    Friend BRT As DataTable = Nothing

    Friend VC As String = Nothing  'Vendor Code
    Friend CC As String = Nothing  'Compay Code

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim SC As New SAPConnector
        RF = True
        Con = SC.GetSAPConnection(Box, User, App)
        If Con Is Nothing Then
            RF = False
            Sts = SC.Status
        End If

    End Sub

    Public Sub New(ByVal Connection As Object)

        If Not Connection Is Nothing AndAlso Connection.Ping Then
            Con = Connection
            RF = True
        Else
            Sts = "Connection already closed"
        End If

    End Sub

    Public ReadOnly Property Status() As String

        Get
            Status = Sts
        End Get

    End Property

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = RF
        End Get

    End Property

    Public ReadOnly Property ResultString(ByVal IncludeWarnings As Boolean) As String

        Get
            Dim S As String
            ResultString = ""
            If Not R Is Nothing Then
                For Each S In R
                    If ((Left(S, 3) = "(W)" Or Left(S, 3) = "(I)") And IncludeWarnings) Or (Left(S, 3) <> "(W)" And Left(S, 3) <> "(I)") Then
                        ResultString = ResultString & S
                        If S <> R(UBound(R)) Then
                            ResultString = ResultString & " \ "
                        End If
                    End If
                Next
                If Right(ResultString, 3) = " \ " Then
                    ResultString = Left(ResultString, Len(ResultString) - 3)
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Results()

        Get
            Results = R
        End Get

    End Property

    Public ReadOnly Property BAPIReturn() As DataTable

        Get
            BAPIReturn = BRT
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = Not (Errors)
        End Get

    End Property

    Public Property IgnoreErrors() As Boolean

        Get
            IgnoreErrors = IE
        End Get
        Set(ByVal value As Boolean)
            IE = value
        End Set

    End Property

    Public Overridable Sub CommitChanges()

        If BAPI Is Nothing Then Exit Sub

        Try
            Errors = False
            BAPI.Execute()
            GetResults()
            If Not Errors Or IE Then
                BAPI.CommitWork(True)
            End If
        Catch ex As Exception
            Errors = True
            ReDim R(0)
            R(0) = ex.Message
        End Try
        BAPI = Nothing
        RF = False

    End Sub

    Friend Function GetItemIndex(ByVal TableName As String, ByVal ColumnName As String, ByVal Item As String) As Integer

        Dim F As Boolean = False
        Dim I As Integer = 0

        If BAPI.Tables(TableName).Rows.Count > 0 Then
            I = 0
            Do While Not F And Not I = BAPI.Tables(TableName).Rows.Count
                If (Val(BAPI.Tables(TableName).Rows(I).Item(ColumnName)) = Val(Item)) Then
                    F = True
                Else
                    I += 1
                End If
            Loop
        End If

        If F Then
            GetItemIndex = I
        Else
            GetItemIndex = -1
        End If

    End Function

    Friend Function GetItemIndex_S0(ByVal TableName As String, ByVal ColumnName As String, ByVal Item As String) As Integer

        Dim F As Boolean = False
        Dim I As Integer = 0

        If BAPI.Tables(TableName).Rows.Count > 0 Then
            I = 0
            Do While Not F And Not I = BAPI.Tables(TableName).Rows.Count
                If (Val(BAPI.Tables(TableName).Rows(I).Item(ColumnName)) = Val(Item)) And Val(BAPI.Tables(TableName).Rows(I).Item("SERIAL_NO")) = 0 Then
                    F = True
                Else
                    I += 1
                End If
            Loop
        End If

        If F Then
            GetItemIndex_S0 = I
        Else
            GetItemIndex_S0 = -1
        End If

    End Function

    Friend Sub GetResults()

        Dim TRow
        Dim S As String
        Dim A() As String = Nothing
        Dim I As Integer = 0

        Try
            If IsNumeric(BAPI.Imports("EXP_HEADER").ParamValue(0).ToString) Then
                DN = BAPI.Imports("EXP_HEADER").ParamValue(0).ToString.Trim
            End If
        Catch ex As Exception
        End Try

        For Each TRow In BAPI.Tables("RETURN").Rows
            If TRow.Item("Type") = "E" Then
                Errors = True
            End If
            If Not TRow.Item("Message") Like "Error transferring ExtensionIn*" _
            And Not TRow.Item("Message") Like "*No instance of object type*" _
            And Not TRow.item("Message") Like "*could not be changed" Then
                S = "(" & TRow.Item("Type") & ") " & TRow.Item("Message")
                If A Is Nothing OrElse Not A.Contains(S) Then
                    ReDim Preserve A(I)
                    A(I) = S
                    I += 1
                End If
            End If
        Next
        If BAPI.Tables("RETURN").Rows.Count > 0 Then BRT = BAPI.Tables("RETURN").ToADOTable()
        R = A

    End Sub

    Friend Overridable Sub SetTableRow(ByVal TableName As String, ByVal Item As String, ByRef TRow As Object, ByRef TRowX As Object)

        Dim F As Boolean = False
        Dim I As Integer = 0
        Dim IL As String

        If Not Item.StartsWith("0") Then Item = Item.PadLeft(5, "0")

        Select Case TableName.ToUpper
            Case "POITEM", "POSCHEDULE", "POACCOUNT", "PO_ITEM_SCHEDULES", "PO_ITEM_ACCOUNT_ASSIGNMENT"
                IL = "PO_ITEM"
            Case "ITEM"
                IL = "ITEM_NO"
            Case Else
                IL = "ITM_NUMBER"
        End Select

        If BAPI.Tables(TableName).Rows.Count > 0 Then
            I = 0
            Do While Not F And Not I = BAPI.Tables(TableName).Rows.Count
                If (BAPI.Tables(TableName).Rows(I).Item(IL) = Item) Then
                    F = True
                Else
                    I += 1
                End If
            Loop
        End If

        If Not F Then
            TRow = BAPI.Tables(TableName).AddRow
            TRowX = BAPI.Tables(TableName & "X").AddRow
        Else
            TRow = BAPI.Tables(TableName).Item(I)
            TRowX = BAPI.Tables(TableName & "X").Item(I)
        End If

        TRow(IL) = Item
        TRowX(IL) = Item

    End Sub

    Friend Function BAPITextSplit(ByVal S As String) As String()

        Dim A() As String = Nothing
        Dim I As Integer = 1
        Dim C As Integer = 1
        Const Chunk = 70

        If Len(S) > Chunk Then
            Do While I <= Len(S)
                ReDim Preserve A(C)
                A(C) = Mid(S, I, Chunk)
                I = I + Chunk
                C = C + 1
            Loop
        Else
            ReDim A(1)
            A(1) = S
        End If

        BAPITextSplit = A

    End Function

    Friend Function HeaderTextArray(ByVal ID As String, ByVal TableName As String) As SAPText()

        Dim S As SAPText() = Nothing
        Dim TRow
        Dim I As Integer = 0

        HeaderTextArray = Nothing
        If Not BAPI Is Nothing Then
            For Each TRow In BAPI.Tables(TableName).Rows
                If TRow(1) = ID Then
                    ReDim Preserve S(I)
                    S(I).Format = TRow("TEXT_FORM").ToString
                    S(I).Text = TRow("TEXT_LINE").ToString
                    I += 1
                End If
            Next
            HeaderTextArray = S
        End If

    End Function

    Friend Function BAPI_Ret_Table() As DataTable

        Dim DT As New DataTable
        DT.Columns.Add("TYPE", Type.GetType("System.String"))
        DT.Columns.Add("ID", Type.GetType("System.String"))
        DT.Columns.Add("NUMBER", Type.GetType("System.String"))
        DT.Columns.Add("MESSAGE", Type.GetType("System.String"))
        DT.Columns.Add("LOG_NO", Type.GetType("System.String"))
        DT.Columns.Add("LOG_MSG_NO", Type.GetType("System.String"))
        DT.Columns.Add("MESSAGE_V1", Type.GetType("System.String"))
        DT.Columns.Add("MESSAGE_V2", Type.GetType("System.String"))
        DT.Columns.Add("MESSAGE_V3", Type.GetType("System.String"))
        DT.Columns.Add("MESSAGE_V4", Type.GetType("System.String"))
        DT.Columns.Add("PARAMETER", Type.GetType("System.String"))
        DT.Columns.Add("ROW", Type.GetType("System.String"))
        DT.Columns.Add("FIELD", Type.GetType("System.String"))
        DT.Columns.Add("SYSTEM", Type.GetType("System.String"))
        BAPI_Ret_Table = DT

    End Function

    Friend Function Check_Vendor_LE_Link() As Boolean

        If VC Is Nothing Or CC Is Nothing Then
            Sts = "There are no "
            If VC Is Nothing Then Sts = Sts & "Vendor Code"
            If CC Is Nothing Then
                If VC Is Nothing Then
                    Sts = Sts & " nor Company Code"
                Else
                    Sts = Sts & "Company Code"
                End If
            End If
            Sts = Sts & " data available for link verification!"
            Return True
        End If

        Check_Vendor_LE_Link = False
        Dim R As New LFB1_Report(Con)
        R.Include_CCode(CC)
        R.IncludeVendor(VC)
        R.Execute()
        If R.Data.Rows.Count > 0 Then
            If R.Data.Rows(0)("Block") <> "X" Then
                Check_Vendor_LE_Link = True
            End If
        End If

    End Function

End Class

Public MustInherit Class Contract_Info : Inherits SC_BAPI_Base

    Friend Contract_Type As Byte = ContractType.Unknown
    Private DocNum As String = Nothing
    Private TD As Boolean = True
    Private AD As Boolean = True
    Private CD As Boolean = True

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property AccountData() As Boolean

        Get
            AccountData = AD
        End Get

        Set(ByVal value As Boolean)
            AD = value
        End Set

    End Property

    Public Property TextData() As Boolean

        Get
            TextData = TD
        End Get

        Set(ByVal value As Boolean)
            TD = value
        End Set

    End Property

    Public Property ConditionData() As Boolean

        Get
            ConditionData = CD
        End Get

        Set(ByVal value As Boolean)
            CD = value
        End Set

    End Property

    Friend Property Number() As String

        Get
            Number = DocNum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing And Contract_Type <> ContractType.Unknown Then
                DocNum = value.Trim.PadLeft(10, "0")
                Try
                    BAPI = Con.CreateBapi(BAPI_Name, "GetDetail")
                    BAPI.Exports("PurchasingDocument").ParamValue = DocNum
                    If TD Then BAPI.Exports("Text_Data").ParamValue = "X"
                    If AD Then BAPI.Exports("Account_Data").ParamValue = "X"
                    If CD Then BAPI.Exports("Condition_Data").ParamValue = "X"
                    BAPI.Execute()
                    If BAPI.Imports("Header").ParamValue("Number") = DocNum Then
                        RF = True
                    Else
                        BAPI = Nothing
                        RF = False
                        Sts = "Document " & value & " does not exist"
                    End If
                Catch ex As Exception
                    BAPI = Nothing
                    RF = False
                    Sts = ex.Message
                End Try
            End If

        End Set

    End Property

    Public ReadOnly Property HeaderText(ByVal ID As String) As String

        Get
            HeaderText = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim HText As String = ""
                For Each TRow In BAPI.Tables("Header_Text").Rows
                    If TRow.Item(0) = ID Then
                        If TRow.Item(2).ToString <> "" Then
                            HText = HText & TRow.Item(2).ToString
                        Else
                            HText = HText & vbCr & vbCr
                        End If
                    End If
                Next
                HeaderText = HText
            End If
        End Get

    End Property

    Public ReadOnly Property PurchGroup() As String

        Get
            If Not BAPI Is Nothing Then
                PurchGroup = BAPI.Imports("HEADER").ParamValue("PUR_GROUP")
            Else
                PurchGroup = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Vendor() As String

        Get
            If Not BAPI Is Nothing Then
                If IsNumeric(BAPI.Imports("HEADER").ParamValue("VENDOR")) Then
                    Vendor = CStr(CDbl(BAPI.Imports("HEADER").ParamValue("VENDOR")))
                Else
                    Vendor = BAPI.Imports("HEADER").ParamValue("VENDOR")
                End If
            Else
                Vendor = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Currency() As String

        Get
            If Not BAPI Is Nothing Then
                Currency = BAPI.Imports("HEADER").ParamValue("CURRENCY")
            Else
                Currency = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Item_Interval() As String

        Get
            If Not BAPI Is Nothing Then
                Item_Interval = BAPI.Imports("HEADER").ParamValue("ITEM_INTVL")
            Else
                Item_Interval = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Company_Code() As String

        Get
            If Not BAPI Is Nothing Then
                Company_Code = BAPI.Imports("HEADER").ParamValue("COMP_CODE")
            Else
                Company_Code = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property ValidityStart() As String

        Get
            If Not BAPI Is Nothing Then
                If IsNumeric(BAPI.Imports("HEADER").ParamValue("VPER_START")) Then
                    If Val(BAPI.Imports("HEADER").ParamValue("VPER_START")) > 0 Then
                        ValidityStart = ConversionUtils.SAPDate2NetDate(BAPI.Imports("HEADER").ParamValue("VPER_START")).ToShortDateString
                    Else
                        ValidityStart = Nothing
                    End If
                Else
                    ValidityStart = Nothing
                End If
            Else
                ValidityStart = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property ValidityEnd() As String

        Get
            If Not BAPI Is Nothing Then
                If IsNumeric(BAPI.Imports("HEADER").ParamValue("VPER_END")) Then
                    If Val(BAPI.Imports("HEADER").ParamValue("VPER_END")) > 0 Then
                        ValidityEnd = ConversionUtils.SAPDate2NetDate(BAPI.Imports("HEADER").ParamValue("VPER_END")).ToShortDateString
                    Else
                        ValidityEnd = Nothing
                    End If
                Else
                    ValidityEnd = Nothing
                End If
            Else
                ValidityEnd = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Material_IsDeleted(ByVal Material As String) As Boolean

        Get
            Material_IsDeleted = False
            If Not BAPI Is Nothing Then
                Dim TRow
                For Each TRow In BAPI.Tables("ITEM").Rows
                    If TRow("MATERIAL") = Material.PadLeft(18, "0") And (TRow("DELETE_IND") = "L" Or TRow("DELETE_IND") = "S") Then
                        Material_IsDeleted = True
                    End If
                Next
            End If
        End Get

    End Property

    Public ReadOnly Property Material_ItemNumbers(ByVal Material As String) As String()

        Get
            Material_ItemNumbers = Nothing
            If Not BAPI Is Nothing Then
                Dim IA() As String = Nothing
                Dim TRow
                For Each TRow In BAPI.Tables("ITEM").Rows
                    If TRow("DELETE_IND") <> "L" And TRow("DELETE_IND") <> "S" Then
                        If TRow("MATERIAL") = Material.PadLeft(18, "0") Then
                            If IA Is Nothing Then
                                ReDim IA(0)
                            Else
                                ReDim Preserve IA(IA.GetUpperBound(0) + 1)
                            End If
                            IA(IA.GetUpperBound(0)) = CStr(Val(TRow.Item("ITEM_NO")))
                        End If
                    End If
                Next
                Material_ItemNumbers = IA
            End If
        End Get

    End Property

    Public ReadOnly Property ItemNumbers() As String()

        Get
            Dim S() As String = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim I As Integer = 0
                For Each TRow In BAPI.Tables("ITEM").Rows
                    ReDim Preserve S(I)
                    S(I) = CStr(Val(TRow.Item("ITEM_NO")))
                    I += 1
                Next
            End If
            ItemNumbers = S
        End Get

    End Property

    Public ReadOnly Property ActiveItems() As String()

        Get
            Dim S() As String = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim I As Integer = 0
                For Each TRow In BAPI.Tables("ITEM").Rows
                    If TRow("DELETE_IND") <> "L" And TRow("DELETE_IND") <> "S" Then
                        ReDim Preserve S(I)
                        S(I) = CStr(Val(TRow("ITEM_NO")))
                        I += 1
                    End If
                Next
            End If
            ActiveItems = S
        End Get

    End Property

    Public ReadOnly Property ItemPrice(ByVal Item As String) As String

        Get
            ItemPrice = Nothing
            If Not BAPI Is Nothing Then
                ItemPrice = Nothing
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If I <> -1 Then
                    ItemPrice = BAPI.Tables("ITEM").Rows(I).Item("NET_PRICE")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemPriceUnit(ByVal Item As String) As String

        Get
            ItemPriceUnit = Nothing
            If Not BAPI Is Nothing Then
                ItemPriceUnit = Nothing
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If I <> -1 Then
                    ItemPriceUnit = BAPI.Tables("ITEM").Rows(I).Item("PRICE_UNIT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemPO_UOM(ByVal Item As String) As String

        Get
            ItemPO_UOM = Nothing
            If Not BAPI Is Nothing Then
                ItemPO_UOM = Nothing
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If I <> -1 Then
                    ItemPO_UOM = BAPI.Tables("ITEM").Rows(I).Item("PO_UNIT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Material(ByVal Item As String) As String

        Get
            Material = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If I <> -1 Then
                    Material = BAPI.Tables("ITEM").Rows(I).Item("MATERIAL")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Plant() As String

        Get
            Plant = Nothing
            If Not BAPI Is Nothing Then
                Plant = BAPI.Tables("ITEM").Rows(0).Item("PLANT")
            End If
        End Get

    End Property

    Public ReadOnly Property Item_Condition_Value(ByVal Item As String, ByVal Condition As String, Optional ByVal VStart As String = Nothing) As String

        Get
            Item_Condition_Value = "N/A"
            If Not BAPI Is Nothing Then
                Dim IC
                If VStart Is Nothing Then
                    IC = ItemLastConditions(Item)
                Else
                    IC = ItemConditions(Item, VStart)
                End If
                If Not IC Is Nothing Then
                    For Each ICR In IC
                        If ICR("COND_TYPE") = Condition Then
                            Item_Condition_Value = ICR("COND_VALUE")
                        End If
                    Next
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Item_Cond_Currency(ByVal Item As String, ByVal Condition As String, Optional ByVal VStart As String = Nothing) As String

        Get
            Item_Cond_Currency = "N/A"
            If Not BAPI Is Nothing Then
                Dim IC
                If VStart Is Nothing Then
                    IC = ItemLastConditions(Item)
                Else
                    IC = ItemConditions(Item, VStart)
                End If
                If Not IC Is Nothing Then
                    For Each ICR In IC
                        If ICR("COND_TYPE") = Condition Then
                            Item_Cond_Currency = ICR("CURRENCY")
                        End If
                    Next
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Item_Condition_Per(ByVal Item As String, ByVal Condition As String, Optional ByVal VStart As String = Nothing) As String

        Get
            Item_Condition_Per = "N/A"
            If Not BAPI Is Nothing Then
                Dim IC
                If VStart Is Nothing Then
                    IC = ItemLastConditions(Item)
                Else
                    IC = ItemConditions(Item, VStart)
                End If
                If Not IC Is Nothing Then
                    For Each ICR In IC
                        If ICR("COND_TYPE") = Condition Then
                            Item_Condition_Per = ICR("COND_P_UNT")
                        End If
                    Next
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Item_Condition_UOM(ByVal Item As String, ByVal Condition As String, Optional ByVal VStart As String = Nothing) As String

        Get
            Item_Condition_UOM = "N/A"
            If Not BAPI Is Nothing Then
                Dim IC
                If VStart Is Nothing Then
                    IC = ItemLastConditions(Item)
                Else
                    IC = ItemConditions(Item, VStart)
                End If
                If Not IC Is Nothing Then
                    For Each ICR In IC
                        If ICR("COND_TYPE") = Condition Then
                            Item_Condition_UOM = ICR("COND_UNIT")
                        End If
                    Next
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Item_LastValidity_Start(ByVal Item As String) As Date

        Get
            Item_LastValidity_Start = Nothing
            Dim ILV = ItemLastValidity(Item)
            If Not ILV Is Nothing Then
                Item_LastValidity_Start = SAPDate2NetDate(ILV("VALID_FROM"))
            End If
        End Get

    End Property

    Public ReadOnly Property Item_LastValidity_End(ByVal Item As String) As Date

        Get
            Item_LastValidity_End = Nothing
            Dim ILV = ItemLastValidity(Item)
            If Not ILV Is Nothing Then
                Item_LastValidity_End = SAPDate2NetDate(ILV("VALID_TO"))
            End If
        End Get

    End Property

    Public ReadOnly Property ItemConditions(ByVal Item As String, ByVal VStart As Date) As Object

        Get
            ItemConditions = Nothing
            Dim VID As String = "0"
            If Not BAPI Is Nothing Then
                Dim TRow
                For Each TRow In BAPI.Tables("Item_Cond_Validity").Rows
                    If Val(TRow("ITEM_NO")) = Val(Item) Then
                        If TRow("VALID_FROM") = NetDate2SAPDate(VStart) Then
                            If CInt(VID) < CInt(TRow("SERIAL_ID")) Then
                                VID = TRow("SERIAL_ID")
                            End If
                        End If
                    End If
                Next
                If VID <> "0" Then
                    Dim OC As New RFCStructureCollection(BAPI.Tables("Item_Condition").Columns)
                    Dim CR
                    For Each CR In BAPI.Tables("Item_Condition").Rows
                        If CR("SERIAL_ID") = VID Then
                            OC.Add(CR)
                        End If
                    Next
                    ItemConditions = OC
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemLastConditions(ByVal Item As String) As Object

        Get
            ItemLastConditions = Nothing
            If Not BAPI Is Nothing Then
                Dim ILV = ItemLastValidity(Item)
                If Not ILV Is Nothing Then
                    Dim LSID As String = ILV("SERIAL_ID")
                    Dim OC As New RFCStructureCollection(BAPI.Tables("Item_Condition").Columns)
                    Dim CR
                    For Each CR In BAPI.Tables("Item_Condition").Rows
                        If CR("SERIAL_ID") = LSID Then
                            OC.Add(CR)
                        End If
                    Next
                    ItemLastConditions = OC
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property LastValidityID() As String

        Get
            LastValidityID = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim I As Int64
                Dim S As String = Nothing
                For Each TRow In BAPI.Tables("Item_Cond_Validity").Rows
                    If CInt(TRow("SERIAL_ID")) > I Then
                        I = CInt(TRow("SERIAL_ID"))
                        S = TRow("SERIAL_ID")
                    End If
                Next
                LastValidityID = S
            End If
        End Get

    End Property

    Public ReadOnly Property ItemLastValidity(ByVal Item As String) As Object

        Get
            ItemLastValidity = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                For Each TRow In BAPI.Tables("Item_Cond_Validity").Rows
                    If Val(TRow("ITEM_NO")) = Val(Item) Then
                        ItemLastValidity = TRow
                    End If
                Next
            End If
        End Get

    End Property

    Private Function BAPI_Name() As String

        BAPI_Name = Nothing
        Select Case Contract_Type
            Case ContractType.OutlineAgreement
                BAPI_Name = "PurchasingContract"
            Case ContractType.SchedulingAgreement
                BAPI_Name = "PurchSchedAgreement"
        End Select

    End Function

End Class

Public NotInheritable Class OAInfo : Inherits Contract_Info

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal Number As String = Nothing)

        MyBase.New(Box, User, App)
        Contract_Type = ContractType.OutlineAgreement
        If RF And Not Number Is Nothing Then
            OANumber = Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal Number As String = Nothing)

        MyBase.New(Connection)
        Contract_Type = ContractType.OutlineAgreement
        If RF And Not Number Is Nothing Then
            OANumber = Number
        End If

    End Sub

    Public Property OANumber() As String

        Get
            OANumber = Number
        End Get

        Set(ByVal value As String)
            If RF Then
                Number = value
            End If
        End Set

    End Property

End Class

Public NotInheritable Class SAInfo : Inherits Contract_Info

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal Number As String = Nothing)

        MyBase.New(Box, User, App)
        Contract_Type = ContractType.SchedulingAgreement
        If RF And Not Number Is Nothing Then
            SANumber = Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal Number As String = Nothing)

        MyBase.New(Connection)
        Contract_Type = ContractType.SchedulingAgreement
        If RF And Not Number Is Nothing Then
            SANumber = Number
        End If

    End Sub

    Public Property SANumber() As String

        Get
            SANumber = Number
        End Get

        Set(ByVal value As String)
            If RF Then
                Number = value
            End If
        End Set

    End Property

End Class

Public NotInheritable Class POInfo1 : Inherits SC_BAPI_Base

    Private PONum As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Box, User, App)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Connection)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Property PONumber() As String

        Get
            PONumber = PONum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing Then
                PONum = value
                RF = False
                BAPI = Nothing
                Try
                    BAPI = Con.CreateBapi("PurchaseOrder", "GetDetail1")
                    BAPI.Exports("PURCHASEORDER").ParamValue = PONumber
                    BAPI.Execute()
                    RF = True
                Catch ex As Exception
                    BAPI = Nothing
                    Sts = ex.Message
                End Try
            End If
        End Set

    End Property

    Public ReadOnly Property BrasNCMCode(ByVal Item As String) As String

        Get
            BrasNCMCode = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If I <> -1 Then
                    BrasNCMCode = BAPI.Tables("POITEM").Rows(I).Item("BRAS_NBM")
                    If IsNumeric(BrasNCMCode) Then BrasNCMCode = CDbl(BrasNCMCode).ToString
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property MaterialOrigin(ByVal Item As String) As String

        Get
            MaterialOrigin = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If I <> -1 Then
                    MaterialOrigin = BAPI.Tables("POITEM").Rows(I).Item("MAT_ORIGIN")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property MaterialUsage(ByVal Item As String) As String

        Get
            MaterialUsage = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If I <> -1 Then
                    MaterialUsage = BAPI.Tables("POITEM").Rows(I).Item("MATL_USAGE")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Condition_Value(ByVal Item As String, ByVal Condition As String) As String

        Get
            Condition_Value = Nothing
            If Not BAPI Is Nothing Then
                For Each ICR In BAPI.Tables("POCOND").Rows
                    If ICR("ITM_NUMBER") = Item.PadLeft(6, "0") And ICR("COND_TYPE") = Condition Then
                        Condition_Value = ICR("COND_VALUE")
                        Exit For
                    End If
                Next
            End If
        End Get

    End Property

    Public ReadOnly Property Condition_No(ByVal Item As String, ByVal Condition As String) As String

        Get
            Condition_No = Nothing
            If Not BAPI Is Nothing Then
                For Each ICR In BAPI.Tables("POCOND").Rows
                    If ICR("ITM_NUMBER") = Item.PadLeft(6, "0") And ICR("COND_TYPE") = Condition Then
                        Condition_No = ICR("CONDITION_NO")
                        Exit For
                    End If
                Next
            End If
        End Get

    End Property

    Public ReadOnly Property Condition_StepNo(ByVal Item As String, ByVal Condition As String) As String

        Get
            Condition_StepNo = Nothing
            If Not BAPI Is Nothing Then
                For Each ICR In BAPI.Tables("POCOND").Rows
                    If ICR("ITM_NUMBER") = Item.PadLeft(6, "0") And ICR("COND_TYPE") = Condition Then
                        Condition_StepNo = ICR("COND_ST_NO")
                        Exit For
                    End If
                Next
            End If
        End Get

    End Property

    Public ReadOnly Property Condition_Count(ByVal Item As String, ByVal Condition As String) As String

        Get
            Condition_Count = Nothing
            If Not BAPI Is Nothing Then
                For Each ICR In BAPI.Tables("POCOND").Rows
                    If ICR("ITM_NUMBER") = Item.PadLeft(6, "0") And ICR("COND_TYPE") = Condition Then
                        Condition_Count = ICR("COND_COUNT")
                        Exit For
                    End If
                Next
            End If
        End Get

    End Property

    Public ReadOnly Property Active_Conditions(ByVal Item As String, Optional ByVal Exceptions() As String = Nothing) As String()

        Get
            Active_Conditions = Nothing
            If Not BAPI Is Nothing Then
                Dim A() As String = Nothing
                For Each ICR In BAPI.Tables("POCOND").Rows
                    If ICR("ITM_NUMBER") = Item.PadLeft(6, "0") Then
                        If Exceptions Is Nothing OrElse Not Exceptions.Contains(ICR("COND_TYPE")) Then
                            If ICR("COND_VALUE") <> 0 Then
                                If A Is Nothing Then
                                    ReDim A(0)
                                Else
                                    ReDim Preserve A(A.Length)
                                End If
                                A(A.GetUpperBound(0)) = ICR("COND_TYPE")
                            End If
                        End If
                    End If
                Next
                Active_Conditions = A
            End If
        End Get

    End Property

End Class

Public NotInheritable Class POInfo : Inherits SC_BAPI_Base

    Private PONum As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Box, User, App)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Connection)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Property PONumber() As String

        Get
            PONumber = PONum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing Then
                PONum = value
                RF = False
                BAPI = Nothing
                Try
                    BAPI = Con.CreateBapi("PurchaseOrder", "GetDetail")
                    BAPI.Exports("PURCHASEORDER").ParamValue = PONumber
                    BAPI.Exports("HEADER_TEXTS").ParamValue = "X"
                    BAPI.Exports("ITEM_TEXTS").ParamValue = "X"
                    BAPI.Exports("HISTORY").ParamValue = "X"
                    BAPI.Exports("SCHEDULES").ParamValue = "X"
                    BAPI.Exports("ACCOUNT_ASSIGNMENT").ParamValue = "X"
                    BAPI.Exports("ITEMS").ParamValue = "X"
                    BAPI.Execute()
                    RF = True
                Catch ex As Exception
                    BAPI = Nothing
                    Sts = ex.Message
                End Try
            End If
        End Set

    End Property

    Public ReadOnly Property Vendor() As String

        Get
            If Not BAPI Is Nothing Then
                If IsNumeric(BAPI.Imports("PO_HEADER").ParamValue(11)) Then
                    Vendor = CStr(CDbl(BAPI.Imports("PO_HEADER").ParamValue(11)))
                Else
                    Vendor = "0"
                End If
            Else
                Vendor = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property VendorName() As String

        Get
            If Not BAPI Is Nothing Then
                VendorName = BAPI.Imports("PO_HEADER").ParamValue(65)
            Else
                VendorName = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property PurchOrg() As String

        Get
            If Not BAPI Is Nothing Then
                PurchOrg = BAPI.Imports("PO_HEADER").ParamValue(19)
            Else
                PurchOrg = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property PurchGrp() As String

        Get
            If Not BAPI Is Nothing Then
                PurchGrp = BAPI.Imports("PO_HEADER").ParamValue(20)
            Else
                PurchGrp = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property CompanyCode() As String

        Get
            If Not BAPI Is Nothing Then
                CompanyCode = BAPI.Imports("PO_HEADER").ParamValue(1)
            Else
                CompanyCode = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property CreatedOn() As String

        Get
            If Not BAPI Is Nothing Then
                CreatedOn = ConversionUtils.SAPDate2NetDate(BAPI.Imports("PO_HEADER").ParamValue(7)).ToShortDateString
            Else
                CreatedOn = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property CreatedBy() As String

        Get
            If Not BAPI Is Nothing Then
                CreatedBy = BAPI.Imports("PO_HEADER").ParamValue(8)
            Else
                CreatedBy = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property LastItem() As String

        Get
            If Not BAPI Is Nothing Then
                LastItem = CStr(Val(BAPI.Imports("PO_HEADER").ParamValue(10)))
            Else
                LastItem = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property ItmInterval() As String

        Get
            If Not BAPI Is Nothing Then
                ItmInterval = CStr(Val(BAPI.Imports("PO_HEADER").ParamValue(9)))
            Else
                ItmInterval = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property PmntTerms() As String

        Get
            If Not BAPI Is Nothing Then
                PmntTerms = BAPI.Imports("PO_HEADER").ParamValue(13)
            Else
                PmntTerms = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Currency() As String

        Get
            If Not BAPI Is Nothing Then
                Currency = BAPI.Imports("PO_HEADER").ParamValue(21)
            Else
                Currency = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property DocDate() As String

        Get
            If Not BAPI Is Nothing Then
                DocDate = ConversionUtils.SAPDate2NetDate(BAPI.Imports("PO_HEADER").ParamValue(24)).ToShortDateString
            Else
                DocDate = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Agreement() As String

        Get
            If Not BAPI Is Nothing Then
                Agreement = BAPI.Imports("PO_HEADER").ParamValue(39)
            Else
                Agreement = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property YReference() As String

        Get
            If Not BAPI Is Nothing Then
                YReference = BAPI.Imports("PO_HEADER").ParamValue(34)
            Else
                YReference = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property OReference() As String

        Get
            If Not BAPI Is Nothing Then
                OReference = BAPI.Imports("PO_HEADER").ParamValue(54)
            Else
                OReference = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property SalesPerson() As String

        Get
            If Not BAPI Is Nothing Then
                SalesPerson = BAPI.Imports("PO_HEADER").ParamValue(35)
            Else
                SalesPerson = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Telephone() As String

        Get
            If Not BAPI Is Nothing Then
                Telephone = BAPI.Imports("PO_HEADER").ParamValue(36)
            Else
                Telephone = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property ItemNetPrice(ByVal Item As String) As String

        Get
            ItemNetPrice = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemNetPrice = BAPI.Tables("PO_ITEMS").Rows(I).Item("NET_PRICE")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemQuantity(ByVal Item As String) As String

        Get
            ItemQuantity = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemQuantity = BAPI.Tables("PO_ITEMS").Rows(I).Item("QUANTITY")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemShortText(ByVal Item As String) As String

        Get
            ItemShortText = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemShortText = BAPI.Tables("PO_ITEMS").Rows(I).Item("SHORT_TEXT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemUOM(ByVal Item As String) As String

        Get
            ItemUOM = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemUOM = BAPI.Tables("PO_ITEMS").Rows(I).Item("UNIT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property TaxCode(ByVal Item As String) As String

        Get
            TaxCode = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    TaxCode = BAPI.Tables("PO_ITEMS").Rows(I).Item("TAX_CODE")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property JurisdCode(ByVal Item As String) As String

        Get
            JurisdCode = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    JurisdCode = BAPI.Tables("PO_ITEMS").Rows(I).Item("TAX_JUR_CD")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property MaterialCategory(ByVal Item As String) As String

        Get
            MaterialCategory = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    MaterialCategory = BAPI.Tables("PO_ITEMS").Rows(I).Item("MAT_CAT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property FinalInvoice(ByVal Item As String) As Boolean

        Get
            FinalInvoice = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("FINAL_INV") = "X" Then
                        FinalInvoice = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property GoodsReceiptInd(ByVal Item As String) As Boolean

        Get
            GoodsReceiptInd = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("GR_IND") = "X" Then
                        GoodsReceiptInd = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property InvoiceReceiptInd(ByVal Item As String) As Boolean

        Get
            InvoiceReceiptInd = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("IR_IND") = "X" Then
                        InvoiceReceiptInd = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property GRBasedIV(ByVal Item As String) As Boolean

        Get
            GRBasedIV = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("GR_BASEDIV") = "X" Then
                        GRBasedIV = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ERS_Flag(ByVal Item As String) As Boolean

        Get
            ERS_Flag = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("ERS") = "X" Then
                        ERS_Flag = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property DeliveryCompleted(ByVal Item As String) As Boolean

        Get
            DeliveryCompleted = False
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PO_ITEMS").Rows(I).Item("DEL_COMPL") = "X" Then
                        DeliveryCompleted = True
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property AccAsignmentCat(ByVal Item As String) As String

        Get
            AccAsignmentCat = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    AccAsignmentCat = BAPI.Tables("PO_ITEMS").Rows(I).Item("ACCTASSCAT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemPriceUnit(ByVal Item As String) As String

        Get
            ItemPriceUnit = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemPriceUnit = BAPI.Tables("PO_ITEMS").Rows(I).Item("PRICE_UNIT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property Item_OrderPriceUnit(ByVal Item As String) As String

        Get
            Item_OrderPriceUnit = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    Item_OrderPriceUnit = BAPI.Tables("PO_ITEMS").Rows(I).Item("ORDERPR_UN")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property MatGroup(ByVal Item As String) As String

        Get
            MatGroup = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEMS", "PO_ITEM", Item)
                If I <> -1 Then
                    MatGroup = BAPI.Tables("PO_ITEMS").Rows(I).Item("MAT_GRP")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property DeliveryDate(ByVal Item As String) As String

        Get
            DeliveryDate = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("Po_Item_Schedules", "PO_ITEM", Item)
                If I <> -1 Then
                    DeliveryDate = SAPDate2NetDate(BAPI.Tables("Po_Item_Schedules").Rows(I).Item("DELIV_DATE"))
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property InternalOrder(ByVal Item As String) As String

        Get
            InternalOrder = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEM_ACCOUNT_ASSIGNMENT", "PO_ITEM", Item)
                If I <> -1 Then
                    InternalOrder = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(I).Item("ORDER_NO")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property CostCenter(ByVal Item As String) As String

        Get
            CostCenter = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEM_ACCOUNT_ASSIGNMENT", "PO_ITEM", Item)
                If I <> -1 Then
                    CostCenter = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(I).Item("COST_CTR")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property GL_Account(ByVal Item As String) As String

        Get
            GL_Account = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PO_ITEM_ACCOUNT_ASSIGNMENT", "PO_ITEM", Item)
                If I <> -1 Then
                    GL_Account = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(I).Item("G_L_ACCT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemQuantityGR(ByVal Item As String) As String

        Get
            ItemQuantityGR = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex_S0("Po_Item_History_Totals", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemQuantityGR = BAPI.Tables("Po_Item_History_Totals").Rows(I).Item("DELIV_QTY")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemQuantityIR(ByVal Item As String) As String

        Get
            ItemQuantityIR = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex_S0("Po_Item_History_Totals", "PO_ITEM", Item)
                If I <> -1 Then
                    ItemQuantityIR = BAPI.Tables("Po_Item_History_Totals").Rows(I).Item("IV_QTY")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemNumbers() As String()

        Get
            Dim S() As String = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim I As Integer = 0
                For Each TRow In BAPI.Tables("PO_ITEMS").Rows
                    ReDim Preserve S(I)
                    S(I) = CStr(Val(TRow.Item("PO_ITEM")))
                    I += 1
                Next
            End If
            ItemNumbers = S
        End Get

    End Property

    Public ReadOnly Property ActiveItems() As String()

        Get
            Dim S() As String = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim I As Integer = 0
                For Each TRow In BAPI.Tables("PO_ITEMS").Rows
                    If TRow("DELETE_IND") <> "L" And TRow("DELETE_IND") <> "S" Then
                        ReDim Preserve S(I)
                        S(I) = CStr(Val(TRow("PO_ITEM")))
                        I += 1
                    End If
                Next
            End If
            ActiveItems = S
        End Get

    End Property

    Public ReadOnly Property HeaderText(ByVal ID As String) As String

        Get
            HeaderText = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim HText As String = ""
                For Each TRow In BAPI.Tables("PO_HEADER_TEXTS").Rows
                    If TRow.Item(1) = ID Then
                        If TRow.Item(3).ToString <> "" Then
                            HText = HText & TRow.Item(3).ToString
                        Else
                            HText = HText & vbCr & vbCr
                        End If
                    End If
                Next
                HeaderText = HText
            End If
        End Get

    End Property

    Public ReadOnly Property ItemText(ByVal Item As String, ByVal ID As String) As String

        Get
            ItemText = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim IText As String = ""
                For Each TRow In BAPI.Tables("PO_ITEM_TEXTS").Rows
                    If TRow("TEXT_ID") = ID And Val(TRow("PO_ITEM")) = Val(Item) Then
                        If TRow("TEXT_LINE").ToString <> "" Then
                            IText = IText & TRow("TEXT_LINE").ToString
                        Else
                            IText = IText & vbCr & vbCr
                        End If
                    End If
                Next
                ItemText = IText
            End If
        End Get

    End Property

    Friend Function LastSchedLine() As Integer

        LastSchedLine = 0
        If BAPI Is Nothing Then Exit Function

        Dim TRow

        For Each TRow In BAPI.Tables("PO_ITEM_SCHEDULES").Rows
            LastSchedLine = Val(TRow("SERIAL_NO"))
        Next

    End Function

End Class

Public NotInheritable Class PRInfo : Inherits SC_BAPI_Base

    Private ReqNum As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal Req_Number As String = Nothing)

        MyBase.New(Box, User, App)
        If RF And Not Req_Number Is Nothing Then
            ReqNumber = Req_Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal Req_Number As String = Nothing)

        MyBase.New(Connection)
        If RF And Not Req_Number Is Nothing Then
            ReqNumber = Req_Number
        End If

    End Sub

    Public Property ReqNumber() As String

        Get
            ReqNumber = ReqNum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing Then
                Try
                    BAPI = Con.CreateBapi("PurchaseRequisition", "GetDetail1")
                    BAPI.Exports("Number").ParamValue = value.PadLeft(10, "0")
                    BAPI.Exports("Header_Text").ParamValue = "X"
                    BAPI.Exports("Account_Assignment").ParamValue = "X"
                    BAPI.Exports("Item_Text").ParamValue = "X"
                    BAPI.Exports("Delivery_Address").ParamValue = "X"
                    BAPI.Execute()
                    ReqNum = value
                    RF = True
                Catch ex As Exception
                    BAPI = Nothing
                    RF = False
                    Sts = ex.Message
                End Try
            End If
        End Set

    End Property

    Public ReadOnly Property OrderType() As String

        Get
            If Not BAPI Is Nothing Then
                OrderType = BAPI.Imports("PRHEADER").ParamValue("PR_TYPE")
            Else
                OrderType = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Plant(ByVal Item As String) As String

        Get
            Plant = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PRITEM", "PREQ_ITEM", Item)
                If I <> -1 Then
                    Plant = BAPI.Tables("PRITEM").Rows(I).Item("PLANT")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemInterval() As String

        Get
            If Not BAPI Is Nothing Then
                ItemInterval = BAPI.Imports("PRHEADER").ParamValue("ITEM_INTVL")
            Else
                ItemInterval = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property ItemNetPrice(ByVal Item As String) As String

        Get
            ItemNetPrice = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PRITEM", "PREQ_ITEM", Item)
                If I <> -1 Then
                    If BAPI.Tables("PRITEM").Rows(I).Item("VALUE_ITEM") <> 0 And BAPI.Tables("PRITEM").Rows(I).Item("QUANTITY") <> 0 Then
                        ItemNetPrice = BAPI.Tables("PRITEM").Rows(I).Item("VALUE_ITEM") / BAPI.Tables("PRITEM").Rows(I).Item("QUANTITY")
                    Else
                        ItemNetPrice = "0"
                    End If
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemTotalPrice(ByVal Item As String) As String

        Get
            ItemTotalPrice = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PRITEM", "PREQ_ITEM", Item)
                If I <> -1 Then
                    ItemTotalPrice = BAPI.Tables("PRITEM").Rows(I).Item("VALUE_ITEM")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemQuantity(ByVal Item As String) As String

        Get
            ItemQuantity = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("PRITEM", "PREQ_ITEM", Item)
                If I <> -1 Then
                    ItemQuantity = BAPI.Tables("PRITEM").Rows(I).Item("QUANTITY")
                End If
            End If
        End Get

    End Property

    Public ReadOnly Property ItemText(ByVal Item As String, ByVal ID As String) As String

        Get
            ItemText = Nothing
            If Not BAPI Is Nothing Then
                Dim TRow
                Dim HText As String = ""
                For Each TRow In BAPI.Tables("PrItemText").Rows
                    If TRow("TEXT_ID") = ID And Val(TRow("PREQ_ITEM")) = Val(Item) Then
                        If TRow("TEXT_LINE").ToString <> "" Then
                            HText = HText & TRow("TEXT_LINE").ToString
                        Else
                            HText = HText & vbCr & vbCr
                        End If
                    End If
                Next
                ItemText = HText
            End If
        End Get

    End Property

    Public Function ItemNumbers() As String()

        Dim S As String() = Nothing
        Dim TRow
        Dim I As Integer = 0

        If Not BAPI Is Nothing Then
            For Each TRow In BAPI.Tables("PRITEM").Rows
                ReDim Preserve S(I)
                S(I) = TRow.Item("PREQ_ITEM")
                I += 1
            Next
        End If

        ItemNumbers = S

    End Function

End Class

#End Region

#Region "Document Changes/Creation"

Public MustInherit Class Contract_Creator : Inherits SC_BAPI_Base

    Friend Contract_Type As Byte = ContractType.Unknown
    Private LI As String = "0"
    Private LII As Integer = 0
    Private CCI As String = Nothing
    Private S_ID As Integer = 0

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub CreateNew(ByVal Vendor As String, ByVal Type As String, ByVal PurchOrg As String, ByVal PurchGrp As String, ByVal CompCode As String, ByVal Currency As String, _
                         Optional ByVal VStart As String = Nothing, Optional ByVal VEnd As String = Nothing)

        If Not RF Or Contract_Type = ContractType.Unknown Then Exit Sub

        Try
            RF = True
            BAPI = Con.CreateBapi(BAPI_Name, "Create")
            BAPI.Exports("Header").ParamValue("VENDOR") = Vendor.PadLeft(10, "0")
            BAPI.Exports("HeaderX").ParamValue("VENDOR") = "X"
            BAPI.Exports("Header").ParamValue("DOC_TYPE") = Type
            BAPI.Exports("HeaderX").ParamValue("DOC_TYPE") = "X"
            BAPI.Exports("Header").ParamValue("PURCH_ORG") = PurchOrg
            BAPI.Exports("HeaderX").ParamValue("PURCH_ORG") = "X"
            BAPI.Exports("Header").ParamValue("PUR_GROUP") = PurchGrp
            BAPI.Exports("HeaderX").ParamValue("PUR_GROUP") = "X"
            BAPI.Exports("Header").ParamValue("COMP_CODE") = CompCode
            BAPI.Exports("HeaderX").ParamValue("COMP_CODE") = "X"
            BAPI.Exports("Header").ParamValue("ITEM_INTVL") = "10"
            BAPI.Exports("HeaderX").ParamValue("ITEM_INTVL") = "X"
            BAPI.Exports("Header").ParamValue("CURRENCY") = Currency
            BAPI.Exports("HeaderX").ParamValue("CURRENCY") = "X"
            If Not VStart Is Nothing AndAlso IsDate(VStart) Then
                BAPI.Exports("Header").ParamValue("VPER_START") = NetDate2SAPDate(VStart)
                BAPI.Exports("HeaderX").ParamValue("VPER_START") = "X"
            End If
            If VEnd Is Nothing OrElse Not IsDate(VEnd) Then VEnd = My.Computer.Clock.LocalTime.ToShortDateString
            BAPI.Exports("Header").ParamValue("VPER_END") = NetDate2SAPDate(VEnd)
            BAPI.Exports("HeaderX").ParamValue("VPER_END") = "X"
        Catch ex As Exception
            ReDim R(0)
            R(0) = ex.Message
            RF = False
        End Try

    End Sub

    Public ReadOnly Property Number() As String

        Get
            Number = DN
        End Get

    End Property

    Public Property Item_Interval() As String

        Get
            If Not BAPI Is Nothing Then
                Item_Interval = BAPI.Exports("Header").ParamValue("ITEM_INTVL")
            Else
                Item_Interval = Nothing
            End If
        End Get
        Set(ByVal value As String)
            If IsNumeric(value) Then
                BAPI.Exports("Header").ParamValue("ITEM_INTVL") = value
                BAPI.Exports("HeaderX").ParamValue("ITEM_INTVL") = "X"
            End If
        End Set

    End Property

    Public Property CompanyCode() As String

        Get
            If Not BAPI Is Nothing Then
                CompanyCode = BAPI.Exports("Header").ParamValue("COMP_CODE")
            Else
                CompanyCode = Nothing
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Exports("Header").ParamValue("COMP_CODE") = value
            BAPI.Exports("HeaderX").ParamValue("COMP_CODE") = "X"
        End Set

    End Property

    Public Property Incoterm() As String

        Get
            If Not BAPI Is Nothing Then
                Incoterm = BAPI.Exports("Header").ParamValue("INCOTERMS1")
            Else
                Incoterm = Nothing
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Exports("Header").ParamValue("INCOTERMS1") = value
            BAPI.Exports("HeaderX").ParamValue("INCOTERMS1") = "X"
        End Set

    End Property

    Public Property Incoterm_Desc() As String

        Get
            If Not BAPI Is Nothing Then
                Incoterm_Desc = BAPI.Exports("Header").ParamValue("INCOTERMS2")
            Else
                Incoterm_Desc = Nothing
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Exports("Header").ParamValue("INCOTERMS2") = value
            BAPI.Exports("HeaderX").ParamValue("INCOTERMS2") = "X"
        End Set

    End Property

    Public Function Item_Add_New(ByVal Material As String, ByVal Quantity As String, ByVal NetPrice As String, ByVal Plant As String, _
                            Optional ByVal Per As Object = Nothing, Optional ByVal UOM As Object = Nothing) As String

        Item_Add_New = Nothing
        If BAPI Is Nothing Then Exit Function

        Dim TRow
        Dim TRowX

        Dim Item As String = NewItemNumber()
        Item_Add_New = Item

        TRow = BAPI.Tables("ITEM").AddRow
        TRowX = BAPI.Tables("ITEMX").AddRow

        TRow("ITEM_NO") = Item
        TRow("MATERIAL") = Material.PadLeft(18, "0")
        TRow("TARGET_QTY") = Quantity
        TRow("NET_PRICE") = NetPrice
        TRow("PLANT") = Plant

        TRowX("ITEM_NO") = Item
        TRowX("MATERIAL") = "X"
        TRowX("TARGET_QTY") = "X"
        TRowX("NET_PRICE") = "X"
        TRowX("PLANT") = "X"

        If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) AndAlso UOM <> "" Then
            TRow("PO_UNIT") = UOM.ToUpper
            TRowX("PO_UNIT") = "X"
        End If
        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) AndAlso Per <> "" Then
            TRow("PRICE_UNIT") = Per
            TRowX("PRICE_UNIT") = "X"
        End If

        LII = BAPI.Tables("ITEM").RowCount - 1

    End Function

    Public Sub Item_NewValidity_Period(ByVal VStart As Date, ByVal VEnd As Date)

        If BAPI Is Nothing Then Exit Sub
        Try
            Dim TRow
            Dim TRowX
            TRow = BAPI.Tables("Item_Cond_Validity").AddRow
            TRowX = BAPI.Tables("Item_Cond_ValidityX").AddRow
            TRow("ITEM_NO") = LI
            TRow("SERIAL_ID") = ValidityId()
            TRow("PLANT") = ""
            TRow("VALID_FROM") = ERPConnect.ConversionUtils.NetDate2SAPDate(VStart)
            TRow("VALID_TO") = ERPConnect.ConversionUtils.NetDate2SAPDate(VEnd)
            TRowX("ITEM_NO") = LI
            TRowX("SERIAL_ID") = "X"
            TRowX("PLANT") = "X"
            TRowX("VALID_FROM") = "X"
            TRowX("VALID_TO") = "X"
        Catch ex As Exception
            Sts = ex.Message
        End Try

    End Sub

    Public Sub Item_Add_Condition(ByVal Condition As String, ByVal Value As String, ByVal Currency As String, ByVal UOM As Object, _
                                  Optional ByVal Per As Object = Nothing, Optional ByVal Vendor As Object = Nothing)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow = Nothing
        Dim TRowX = Nothing

        TRow = BAPI.Tables("Item_Condition").AddRow
        TRowX = BAPI.Tables("Item_ConditionX").AddRow

        TRow("ITEM_NO") = LI
        TRow("SERIAL_ID") = ValidityId()
        TRow("COND_COUNT") = NewCondCount()
        TRow("COND_TYPE") = Condition.ToUpper
        TRow("COND_VALUE") = Value
        TRow("NUMERATOR") = "1"
        TRow("DENOMINATOR") = "1"
        TRow("CURRENCY") = Currency.ToUpper
        If UOM Is Nothing OrElse DBNull.Value.Equals(UOM) Then UOM = ""
        TRow("COND_UNIT") = UOM.ToUpper

        TRowX("ITEM_NO") = LI
        TRowX("SERIAL_ID") = "X"
        TRowX("COND_COUNT") = "X"
        TRowX("COND_TYPE") = "X"
        TRowX("COND_VALUE") = "X"
        TRowX("NUMERATOR") = "X"
        TRowX("DENOMINATOR") = "X"
        TRowX("CURRENCY") = "X"
        TRowX("COND_UNIT") = "X"

        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) AndAlso Per <> "" Then
            TRow("COND_P_UNT") = Per
            TRowX("COND_P_UNT") = "X"
        End If
        If Not Vendor Is Nothing AndAlso Not DBNull.Value.Equals(Vendor) AndAlso Vendor <> "" Then
            TRow("VENDOR_NO") = Vendor.PadLeft(10, "0")
            TRowX("VENDOR_NO") = "X"
        End If

    End Sub

    Public Property Item_ConfControl() As String

        Get
            Item_ConfControl = Nothing
            If Not BAPI Is Nothing Then
                Item_ConfControl = BAPI.Tables("ITEM")(LII)("CONF_CTRL")
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Tables("ITEM")(LII)("CONF_CTRL") = value
            BAPI.Tables("ITEMX")(LII)("CONF_CTRL") = "X"
        End Set

    End Property

    Public Property Item_UnderDelTolerance() As String

        Get
            Item_UnderDelTolerance = Nothing
            If Not BAPI Is Nothing Then
                Item_UnderDelTolerance = BAPI.Tables("ITEM")(LII)("UNDER_DLV_TOL")
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Tables("ITEM")(LII)("UNDER_DLV_TOL") = value
            BAPI.Tables("ITEMX")(LII)("UNDER_DLV_TOL") = "X"
        End Set

    End Property

    Public Property Item_OverDelTolerance() As String

        Get
            Item_OverDelTolerance = Nothing
            If Not BAPI Is Nothing Then
                Item_OverDelTolerance = BAPI.Tables("ITEM")(LII)("OVER_DLV_TOL")
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Tables("ITEM")(LII)("OVER_DLV_TOL") = value
            BAPI.Tables("ITEMX")(LII)("OVER_DLV_TOL") = "X"
        End Set

    End Property

    Private Function BAPI_Name() As String

        BAPI_Name = Nothing
        Select Case Contract_Type
            Case ContractType.OutlineAgreement
                BAPI_Name = "PurchasingContract"
            Case ContractType.SchedulingAgreement
                BAPI_Name = "PurchSchedAgreement"
        End Select

    End Function

    Private Function NewItemNumber() As String

        Dim NIN As String
        NewItemNumber = Nothing
        If BAPI Is Nothing Then Exit Function
        NIN = (Val(LI) + Val(BAPI.Exports("Header").ParamValue("ITEM_INTVL"))).ToString
        NewItemNumber = NIN
        LI = NIN
        S_ID += 1
        CCI = "1"

    End Function

    Private Function ValidityId() As String

        ValidityId = S_ID.ToString.PadLeft(10, "0")

    End Function

    Private Function NewCondCount() As String

        NewCondCount = CCI.ToString.PadLeft(2, "0")
        CCI += 1

    End Function

End Class

Public NotInheritable Class OACreator : Inherits Contract_Creator

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)
        Contract_Type = ContractType.OutlineAgreement

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)
        Contract_Type = ContractType.OutlineAgreement

    End Sub

End Class

Public NotInheritable Class SACreator : Inherits Contract_Creator

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)
        Contract_Type = ContractType.SchedulingAgreement

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)
        Contract_Type = ContractType.SchedulingAgreement

    End Sub

End Class

Public MustInherit Class Contract_Changes : Inherits SC_BAPI_Base

    Private Contract_Type As Byte = ContractType.Unknown
    Private Info_Type As Type
    Private DocNum As String = Nothing
    Private OInfo = Nothing
    Private S_ID As Integer = -1
    Private LI As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, ByVal CT As Byte)

        MyBase.New(Box, User, App)
        Contract_Type = CT

    End Sub

    Public Sub New(ByVal Connection As Object, ByVal CT As Byte)

        MyBase.New(Connection)
        Contract_Type = CT

    End Sub

    Public Property Info() As Object

        Get
            If OInfo Is Nothing And Not Contract_Type = ContractType.Unknown And Not DocNum Is Nothing Then
                SetupInfoObject()
            End If
            Info = OInfo
        End Get

        Set(ByVal value)
            If ValidateInfoObject(value) Then
                OInfo = value
            End If
        End Set

    End Property

    Friend Property Number() As String

        Get
            Number = DocNum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing And Contract_Type <> ContractType.Unknown Then
                DocNum = value.Trim
                Try
                    RF = True
                    BAPI = Con.CreateBapi(BAPI_Name, "Change")
                    BAPI.Exports("PurchasingDocument").ParamValue = DocNum.PadLeft(10, "0")
                    OInfo = Nothing
                    S_ID = -1
                Catch ex As Exception
                    ReDim R(0)
                    R(0) = ex.Message
                    RF = False
                End Try
            End If
        End Set

    End Property

    Public Property Payment_Terms() As String

        Get
            Payment_Terms = Nothing
            If Not BAPI Is Nothing Then
                Payment_Terms = BAPI.Exports("HEADER").ParamValue("PMNTTRMS")
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("HEADER").ParamValue("PMNTTRMS") = value
                BAPI.Exports("HEADERX").ParamValue("PMNTTRMS") = "X"
            End If
        End Set

    End Property

    Public Property Purchasing_Group() As String

        Get
            Purchasing_Group = Nothing
            If Not BAPI Is Nothing Then
                Purchasing_Group = BAPI.Exports("HEADER").ParamValue("PUR_GROUP")
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("HEADER").ParamValue("PUR_GROUP") = value
                BAPI.Exports("HEADERX").ParamValue("PUR_GROUP") = "X"
            End If
        End Set

    End Property

    Public Property Validty_End() As String

        Get
            Validty_End = Nothing
            If Not BAPI Is Nothing Then
                If IsNumeric(BAPI.Exports("HEADER").ParamValue("VPER_END")) Then
                    Validty_End = SAPDate2NetDate(BAPI.Exports("HEADER").ParamValue("VPER_END")).ToShortDateString
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                If IsDate(value) Then
                    BAPI.Exports("HEADER").ParamValue("VPER_END") = NetDate2SAPDate(value)
                    BAPI.Exports("HEADERX").ParamValue("VPER_END") = "X"
                End If
            End If
        End Set

    End Property

    Public Property DeletionInd_Deleted(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    If BAPI.Tables("ITEM").Rows(I).Item("DELETE_IND") = "L" Then
                        DeletionInd_Deleted = True
                    Else
                        DeletionInd_Deleted = False
                    End If
                Else
                    DeletionInd_Deleted = False
                End If
            Else
                DeletionInd_Deleted = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                If value Then
                    TRow("DELETE_IND") = "L"
                Else
                    TRow("DELETE_IND") = " "
                End If
                TRowX("DELETE_IND") = "X"
            End If
        End Set

    End Property

    Public Property DeletionInd_Blocked(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    If BAPI.Tables("ITEM").Rows(I).Item("DELETE_IND") = "S" Then
                        DeletionInd_Blocked = True
                    Else
                        DeletionInd_Blocked = False
                    End If
                Else
                    DeletionInd_Blocked = False
                End If
            Else
                DeletionInd_Blocked = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                If value Then
                    TRow("DELETE_IND") = "S"
                Else
                    TRow("DELETE_IND") = " "
                End If
                TRowX("DELETE_IND") = "X"
            End If
        End Set

    End Property

    Public Property PrintPrice(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    If BAPI.Tables("ITEM").Rows(I).Item("PRNT_PRICE") = "X" Then
                        PrintPrice = True
                    Else
                        PrintPrice = False
                    End If
                Else
                    PrintPrice = False
                End If
            Else
                PrintPrice = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                If value Then
                    TRow("PRNT_PRICE") = "X"
                Else
                    TRow("PRNT_PRICE") = " "
                End If
                TRowX("PRNT_PRICE") = "X"
            End If
        End Set

    End Property

    Public Property PDT(ByVal Item As String) As String

        Get
            PDT = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    PDT = BAPI.Tables("ITEM").Rows(I).Item("PLAN_DEL")
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("PLAN_DEL") = value
                TRowX("PLAN_DEL") = "X"
            End If
        End Set

    End Property

    Public Property StorageLocation(ByVal Item As String) As String

        Get
            StorageLocation = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    StorageLocation = BAPI.Tables("ITEM").Rows(I).Item("STGE_LOC")
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("STGE_LOC") = value
                TRowX("STGE_LOC") = "X"
            End If
        End Set

    End Property

    Public Property OverDelTolerance(ByVal Item As String) As String

        Get
            OverDelTolerance = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    OverDelTolerance = BAPI.Tables("ITEM").Rows(I).Item("OVER_DLV_TOL")
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("OVER_DLV_TOL") = value
                TRowX("OVER_DLV_TOL") = "X"
            End If
        End Set

    End Property

    Public Property UnderDelTolerance(ByVal Item As String) As String

        Get
            UnderDelTolerance = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    UnderDelTolerance = BAPI.Tables("ITEM").Rows(I).Item("UNDER_DLV_TOL")
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("UNDER_DLV_TOL") = value
                TRowX("UNDER_DLV_TOL") = "X"
            End If
        End Set

    End Property

    Public Property ConfControl(ByVal Item As String) As String

        Get
            ConfControl = Nothing
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    ConfControl = BAPI.Tables("ITEM").Rows(I).Item("CONF_CTRL")
                End If
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("CONF_CTRL") = value
                TRowX("CONF_CTRL") = "X"
            End If
        End Set

    End Property

    Public Property TaxCode(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("ITEM", "ITEM_NO", Item)
                If Not I < 0 Then
                    TaxCode = BAPI.Tables("ITEM").Rows(I).Item("TAX_CODE")
                Else
                    TaxCode = Nothing
                End If
            Else
                TaxCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("ITEM", Item, TRow, TRowX)
                TRow("TAX_CODE") = value.ToUpper
                TRowX("TAX_CODE") = "X"
            End If
        End Set

    End Property

    Public Function AddMaterial(ByVal Material As String, ByVal Quantity As String, ByVal NetPrice As String, ByVal Plant As String, _
                                Optional ByVal Per As Object = Nothing, Optional ByVal UOM As Object = Nothing) As String

        AddMaterial = Nothing
        If BAPI Is Nothing Then Exit Function

        Dim TRow
        Dim TRowX

        Dim Item As String = NewItemNumber()

        TRow = BAPI.Tables("ITEM").AddRow
        TRowX = BAPI.Tables("ITEMX").AddRow

        TRow("ITEM_NO") = Item.PadLeft(5, "0")
        TRow("MATERIAL") = Material.PadLeft(18, "0")
        TRow("TARGET_QTY") = Quantity
        TRow("NET_PRICE") = NetPrice
        TRow("PLANT") = Plant

        TRowX("ITEM_NO") = Item.PadLeft(5, "0")
        TRowX("MATERIAL") = "X"
        TRowX("TARGET_QTY") = "X"
        TRowX("NET_PRICE") = "X"
        TRowX("PLANT") = "X"

        If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) AndAlso UOM <> "" Then
            TRow("PO_UNIT") = UOM.ToUpper
            TRowX("PO_UNIT") = "X"
        End If
        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) AndAlso Per <> "" Then
            TRow("PRICE_UNIT") = Per
            TRowX("PRICE_UNIT") = "X"
        End If

        AddMaterial = Item

    End Function

    Public Function Update_Item_Last_Validity(ByVal Item As String, Optional ByVal Price As Object = Nothing, Optional ByVal Per As Object = Nothing, Optional ByVal UOM As Object = Nothing) As String

        Update_Item_Last_Validity = Nothing
        If BAPI Is Nothing Then Exit Function
        If Not Get_Updating_VP_ID(Item) Is Nothing Then Exit Function
        If OInfo Is Nothing Then SetupInfoObject()
        Item = Item.PadLeft(5, "0")
        Try
            Dim TRow
            Dim TRowX

            Dim ILV = OInfo.ItemLastValidity(Item)
            TRow = BAPI.Tables("Item_Cond_Validity").AddRow
            TRowX = BAPI.Tables("Item_Cond_ValidityX").AddRow
            TRow("ITEM_NO") = Item
            TRow("SERIAL_ID") = ILV("SERIAL_ID")
            TRow("PLANT") = ILV("PLANT")
            TRow("VALID_FROM") = ILV("VALID_FROM")
            TRow("VALID_TO") = ILV("VALID_TO")
            TRowX("VALID_FROM") = "X"
            TRowX("VALID_TO") = "X"
            TRowX("ITEM_NO") = Item
            TRowX("SERIAL_ID") = "X"
            TRowX("PLANT") = "X"

            Dim CR
            Dim I As Integer
            For Each CR In OInfo.ItemLastConditions(Item)
                If CR("COND_TYPE") = "PB00" Then
                    If Not Price Is Nothing AndAlso Not DBNull.Value.Equals(Price) AndAlso Price <> "" Then
                        CR("COND_VALUE") = Price
                    End If
                    If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) AndAlso Per <> "" Then
                        CR("COND_P_UNT") = Per
                    End If
                    If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) AndAlso UOM <> "" Then
                        CR("COND_UNIT") = UOM.ToUpper
                    End If
                End If
                CR("ITEM_NO") = Item
                TRow = BAPI.Tables("Item_Condition").AddRow
                For I = 0 To CR.columns.count - 1
                    TRow(I) = CR(I)
                Next
                TRowX = BAPI.Tables("Item_ConditionX").AddRow
                For I = 0 To TRowX.columns.count - 1
                    TRowX(I) = "X"
                Next
                TRowX("ITEM_NO") = Item
                Update_Item_Last_Validity = ILV("SERIAL_ID")
            Next
        Catch ex As Exception
            Sts = ex.Message
        End Try

    End Function

    Public Function Item_NewValidity_WithReference(ByVal Item As String, ByVal VStart As Date, ByVal VEnd As Date, ByVal Price As String, _
            Optional ByVal Per As Object = Nothing, Optional ByVal UOM As Object = Nothing) As String

        Item_NewValidity_WithReference = Nothing
        If BAPI Is Nothing Then Exit Function

        Try
            Dim VID As String = NewValidityId()
            Item = Item.PadLeft(5, "0")
            Dim TRow
            Dim TRowX

            TRow = BAPI.Tables("Item_Cond_Validity").AddRow
            TRowX = BAPI.Tables("Item_Cond_ValidityX").AddRow
            TRow("ITEM_NO") = Item
            TRow("SERIAL_ID") = VID
            TRow("PLANT") = ""
            TRow("VALID_FROM") = ERPConnect.ConversionUtils.NetDate2SAPDate(VStart)
            TRow("VALID_TO") = ERPConnect.ConversionUtils.NetDate2SAPDate(VEnd)
            TRowX("ITEM_NO") = Item
            TRowX("SERIAL_ID") = "X"
            TRowX("PLANT") = "X"
            TRowX("VALID_FROM") = "X"
            TRowX("VALID_TO") = "X"

            Dim CR = Nothing
            Dim I As Integer
            For Each CR In OInfo.ItemLastConditions(Item)
                CR("SERIAL_ID") = VID
                If CR("COND_TYPE") = "PB00" Then
                    CR("COND_VALUE") = Price
                    If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) AndAlso Per <> "" Then
                        CR("COND_P_UNT") = Per
                    End If
                    If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) AndAlso UOM <> "" Then
                        CR("COND_UNIT") = UOM.ToUpper
                    End If
                End If
                If CR("SCALE_BASE_TY") <> "" And CR("COND_TYPE") = "ZOA1" Then
                    CR("SCALE_BASE_TY") = ""
                End If
                If (CR("SCALE_BASE_TY") = "C" Or CR("SCALE_BASE_TY") = "B") Then
                    If CR("SCALE_UNIT") = "" Then CR("SCALE_UNIT") = CR("COND_UNIT")
                    If CR("SCALE_CURR") = "" Then CR("SCALE_CURR") = CR("CURRENCY")
                End If
                TRow = BAPI.Tables("Item_Condition").AddRow
                For I = 0 To CR.columns.count - 1
                    TRow(I) = CR(I)
                Next
                TRowX = BAPI.Tables("Item_ConditionX").AddRow
                For I = 0 To TRowX.columns.count - 1
                    TRowX(I) = "X"
                Next
                TRowX("ITEM_NO") = Item
            Next
            Item_NewValidity_WithReference = VID
        Catch ex As Exception
            Sts = ex.Message
        End Try

    End Function

    Public Function Item_NewValidity_Period(ByVal Item As String, ByVal VStart As Date, ByVal VEnd As Date) As String

        Item_NewValidity_Period = Nothing
        If BAPI Is Nothing Then Exit Function
        Try

            'For Each CR In BAPI.Tables("Item_Cond_Validity").Rows
            '    If CR("ITEM_NO") = Item And CR("VALID_FROM") = ERPConnect.ConversionUtils.SAPDate2NetDate(VStart) And CR("VALID_TO") = ERPConnect.ConversionUtils.SAPDate2NetDate(VEnd) Then
            '        Item_NewValidity_Period = CR("SERIAL_ID")
            '        Exit For                                         *********************************************
            '    End If                                               * This is to allow multiple period creation *
            'Next                                                     *********************************************
            'If Not Item_NewValidity_Period Is Nothing Then Exit Function
            Item = Item.PadLeft(5, "0")
            Dim VID As String = Get_Updating_VP_ID(Item)
            If VID Is Nothing Then
                VID = NewValidityId()
            Else
                Item_NewValidity_Period = VID
                Exit Function
            End If

            Dim TRow
            Dim TRowX
            TRow = BAPI.Tables("Item_Cond_Validity").AddRow
            TRowX = BAPI.Tables("Item_Cond_ValidityX").AddRow
            TRow("ITEM_NO") = Item
            TRow("SERIAL_ID") = VID
            TRow("PLANT") = ""
            TRow("VALID_FROM") = ERPConnect.ConversionUtils.NetDate2SAPDate(VStart)
            TRow("VALID_TO") = ERPConnect.ConversionUtils.NetDate2SAPDate(VEnd)
            TRowX("ITEM_NO") = Item
            TRowX("SERIAL_ID") = "X"
            TRowX("PLANT") = "X"
            TRowX("VALID_FROM") = "X"
            TRowX("VALID_TO") = "X"
            Item_NewValidity_Period = VID

        Catch ex As Exception
            Sts = ex.Message
        End Try

    End Function

    Public Sub Item_Condition(ByVal Item As String, ByVal Condition As String, ByVal Value As String, ByVal Currency As Object, ByVal UOM As Object, _
                                  Optional ByVal Per As Object = Nothing, Optional ByVal Vendor As Object = Nothing, _
                                  Optional ByVal VID As String = Nothing, Optional ByVal CC As String = Nothing)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow = Nothing
        Dim TRowX = Nothing
        Item = Item.PadLeft(5, "0")
        Get_Updating_Item_Condition(Item, Condition, TRow, TRowX)
        If TRow Is Nothing Then

            If VID Is Nothing Then
                VID = Get_Updating_VP_ID(Item)
                If VID Is Nothing Then
                    Update_Item_Last_Validity(Item)
                    VID = Get_Updating_VP_ID(Item)
                End If
            End If

            If CC Is Nothing Then
                Dim Temp_CC As String = Last_Cond_Count(Item, Condition)
                CC = CStr(Val(Temp_CC) + 1)
            End If
            CC = CC.PadLeft(2, "0")

            TRow = BAPI.Tables("Item_Condition").AddRow
            TRowX = BAPI.Tables("Item_ConditionX").AddRow

            TRow("ITEM_NO") = Item
            TRow("SERIAL_ID") = VID
            TRow("COND_COUNT") = CC
            TRow("COND_TYPE") = Condition.Trim.ToUpper
            TRow("NUMERATOR") = "1"
            TRow("DENOMINATOR") = "1"

            TRowX("ITEM_NO") = Item
            TRowX("SERIAL_ID") = "X"
            TRowX("COND_COUNT") = "X"
            TRowX("COND_TYPE") = "X"
            TRowX("NUMERATOR") = "X"
            TRowX("DENOMINATOR") = "X"

        End If

        TRow("COND_VALUE") = Value
        If Currency Is Nothing OrElse DBNull.Value.Equals(Currency) Then Currency = ""
        TRow("CURRENCY") = Currency.ToUpper
        If UOM Is Nothing OrElse DBNull.Value.Equals(UOM) Then UOM = ""
        TRow("COND_UNIT") = UOM.ToUpper

        TRowX("COND_VALUE") = "X"
        TRowX("CURRENCY") = "X"
        TRowX("COND_UNIT") = "X"

        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) Then
            TRow("COND_P_UNT") = Per
            TRowX("COND_P_UNT") = "X"
        End If
        If Not Vendor Is Nothing AndAlso Not DBNull.Value.Equals(Vendor) Then
            TRow("VENDOR_NO") = Vendor.ToString.PadLeft(10, "0")
            TRowX("VENDOR_NO") = "X"
        End If

    End Sub

    Private Function Last_Cond_Count(ByVal Item As String, ByVal Condition As String) As String

        Last_Cond_Count = "0"
        For Each CR In BAPI.Tables("Item_Condition").Rows
            If CR("ITEM_NO") = Item Then
                If CR("COND_TYPE") = Condition Then Exit For
                Last_Cond_Count = CR("COND_COUNT")
            End If
        Next

    End Function

    Private Function Get_Updating_VP_ID(ByVal Item As String) As String

        Get_Updating_VP_ID = Nothing
        For Each CR In BAPI.Tables("Item_Cond_Validity").Rows
            If CR("ITEM_NO") = Item Then
                Get_Updating_VP_ID = CR("SERIAL_ID")
            End If
        Next

    End Function

    Private Function NewValidityId() As String

        If Con Is Nothing Or DocNum Is Nothing Then
            Return Nothing
        End If

        If S_ID = -1 Then
            If OInfo Is Nothing Then SetupInfoObject()
            S_ID = CInt(OInfo.LastValidityID) + 1
        End If

        NewValidityId = S_ID.ToString.PadLeft(10, "0")
        S_ID += 1

    End Function

    Private Function NewItemNumber() As String

        Dim NIN As String
        If LI Is Nothing Then
            If OInfo Is Nothing Then SetupInfoObject()
            LI = CStr(Val(OInfo.ItemNumbers(OInfo.ItemNumbers.GetUpperBound(0))))
        End If
        NIN = (Val(LI) + Val(OInfo.Item_Interval)).ToString
        NewItemNumber = NIN
        LI = NIN

    End Function

    Private Function BAPI_Name() As String

        BAPI_Name = Nothing
        Select Case Contract_Type
            Case ContractType.OutlineAgreement
                BAPI_Name = "PurchasingContract"
            Case ContractType.SchedulingAgreement
                BAPI_Name = "PurchSchedAgreement"
        End Select

    End Function

    Private Function ValidateInfoObject(ByVal O As Object) As Boolean

        ValidateInfoObject = False
        Select Case Contract_Type
            Case ContractType.OutlineAgreement
                ValidateInfoObject = O.GetType.Equals(GetType(OAInfo))
            Case ContractType.SchedulingAgreement
                ValidateInfoObject = O.GetType.Equals(GetType(SAInfo))
        End Select

    End Function

    Private Sub SetupInfoObject()

        Select Case Contract_Type
            Case ContractType.OutlineAgreement
                OInfo = New OAInfo(Con, DocNum)
            Case ContractType.SchedulingAgreement
                OInfo = New SAInfo(Con, DocNum)
        End Select

    End Sub

    Private Sub Get_Updating_Item_Condition(ByVal Item As String, ByVal Condition As String, ByRef TRow As RFCStructure, ByRef TRowX As RFCStructure)

        For Each CR In BAPI.Tables("Item_Condition").Rows
            If CR("ITEM_NO") = Item And CR("COND_TYPE") = Condition.Trim.ToUpper Then
                TRow = CR
            End If
        Next

        If Not TRow Is Nothing Then
            For Each CR In BAPI.Tables("Item_ConditionX").Rows
                If CR("ITEM_NO") = Item Then
                    TRowX = CR
                End If
            Next
        End If

    End Sub

    Public Sub NewHeaderText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim A As String()
        Dim I As Integer
        Dim TRow
        Dim SR As New StringReader(Text)
        Dim S As String

        While True
            S = SR.ReadLine
            If S Is Nothing Then
                Exit While
            Else
                A = BAPITextSplit(S)
                I = 1
                Do While I <= UBound(A)
                    TRow = BAPI.Tables("Header_Text").AddRow
                    TRow("TEXT_ID") = TextID
                    If I = 1 Then
                        TRow("TEXT_FORM") = "*"
                    Else
                        TRow("TEXT_FORM") = ""
                    End If
                    TRow("TEXT_LINE") = A(I)
                    I = I + 1
                Loop
            End If
        End While

    End Sub

    Public Sub InsertHeaderText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow
        Dim SR As New StringReader(Text)

        NewHeaderText(TextID, Text)

        If OInfo Is Nothing Then SetupInfoObject()
        Dim HTA As SAPText() = CType(OInfo, Contract_Info).HeaderTextArray(TextID, "Header_Text")
        Dim HT As SAPText
        If Not HTA Is Nothing Then
            For Each HT In HTA
                TRow = BAPI.Tables("Header_Text").AddRow
                TRow("TEXT_ID") = TextID
                TRow("TEXT_FORM") = HT.Format
                TRow("TEXT_LINE") = HT.Text
            Next
        End If

    End Sub

End Class

Public NotInheritable Class OAChanges : Inherits Contract_Changes

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal CNumber As String = Nothing)

        MyBase.New(Box, User, App, ContractType.OutlineAgreement)
        If RF And Not CNumber Is Nothing Then
            Number = CNumber
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal OInfo As Object = Nothing, Optional ByVal CNumber As String = Nothing)

        MyBase.New(Connection, ContractType.OutlineAgreement)
        If RF And Not CNumber Is Nothing Then
            Number = CNumber
        End If
        If Not OInfo Is Nothing Then
            Info = OInfo
        End If

    End Sub

    Public Property OANumber() As String

        Get
            OANumber = Number
        End Get

        Set(ByVal value As String)
            If RF Then
                Number = value
            End If
        End Set

    End Property

End Class

Public NotInheritable Class SAChanges : Inherits Contract_Changes

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal CNumber As String = Nothing)

        MyBase.New(Box, User, App, ContractType.SchedulingAgreement)
        If RF And Not CNumber Is Nothing Then
            Number = CNumber
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal OInfo As Object = Nothing, Optional ByVal CNumber As String = Nothing)

        MyBase.New(Connection, ContractType.SchedulingAgreement)
        If RF And Not CNumber Is Nothing Then
            Number = CNumber
        End If
        If Not OInfo Is Nothing Then
            Info = OInfo
        End If

    End Sub

    Public Property SANumber() As String

        Get
            SANumber = Number
        End Get

        Set(ByVal value As String)
            If RF Then
                Number = value
            End If
        End Set

    End Property

End Class

Public NotInheritable Class POChanges : Inherits SC_BAPI_Base

    Private PO As POInfo = Nothing
    Private P1 As POInfo1 = Nothing
    Private PONum As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Box, User, App)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Sub New(ByVal Connection As Object, Optional ByVal PO_Number As String = Nothing)

        MyBase.New(Connection)
        If RF And Not PO_Number Is Nothing Then
            PONumber = PO_Number
        End If

    End Sub

    Public Property Info() As POInfo

        Get
            Info = PO
        End Get
        Set(ByVal value As POInfo)
            If value.GetType.Equals(GetType(POInfo)) Then PO = value
        End Set

    End Property

    Public Property Info1() As POInfo1

        Get
            Info1 = P1
        End Get
        Set(ByVal value As POInfo1)
            If value.GetType.Equals(GetType(POInfo1)) Then P1 = value
        End Set

    End Property

    Public Property PONumber() As String

        Get
            PONumber = PONum
        End Get

        Set(ByVal value As String)
            If Not Con Is Nothing Then
                PONum = value
                Try
                    RF = True
                    BAPI = Con.CreateBapi("PurchaseOrder", "Change")
                    BAPI.Exports("PURCHASEORDER").ParamValue = PONumber
                Catch ex As Exception
                    ReDim R(0)
                    R(0) = ex.Message
                    RF = False
                End Try
            End If
        End Set

    End Property

    Public Property GenerateOutput() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Exports("NO_MESSAGING").ParamValue = "X" Then
                    GenerateOutput = False
                Else
                    GenerateOutput = True
                End If
            Else
                GenerateOutput = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Exports("NO_MESSAGING").ParamValue = " "
                    BAPI.Exports("NO_MESSAGE_REQ").ParamValue = " "
                Else
                    BAPI.Exports("NO_MESSAGING").ParamValue = "X"
                    BAPI.Exports("NO_MESSAGE_REQ").ParamValue = "X"
                End If
            End If
        End Set

    End Property

    Public Property PaimentTerms() As String

        Get
            If Not BAPI Is Nothing Then
                PaimentTerms = BAPI.Exports("POHEADER").ParamValue(11)
            Else
                PaimentTerms = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(11) = value
                BAPI.Exports("POHEADERX").ParamValue(11) = "X"
            End If
        End Set

    End Property

    Public Property PurchGroup() As String

        Get
            If Not BAPI Is Nothing Then
                PurchGroup = BAPI.Exports("POHEADER").ParamValue(18)
            Else
                PurchGroup = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(18) = value
                BAPI.Exports("POHEADERX").ParamValue(18) = "X"
            End If
        End Set

    End Property

    Public Property PurchOrg() As String

        Get
            If Not BAPI Is Nothing Then
                PurchOrg = BAPI.Exports("POHEADER").ParamValue(17)
            Else
                PurchOrg = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(17) = value
                BAPI.Exports("POHEADERX").ParamValue(17) = "X"
            End If
        End Set

    End Property

    Public Property OurReference() As String

        Get
            If Not BAPI Is Nothing Then
                OurReference = BAPI.Exports("POHEADER").ParamValue(41)
            Else
                OurReference = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(41) = value
                BAPI.Exports("POHEADERX").ParamValue(41) = "X"
            End If
        End Set

    End Property

    Public Property YourReference() As String

        Get
            If Not BAPI Is Nothing Then
                YourReference = BAPI.Exports("POHEADER").ParamValue(29)
            Else
                YourReference = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(29) = value
                BAPI.Exports("POHEADERX").ParamValue(29) = "X"
            End If
        End Set

    End Property

    Public Property InvoicingParty() As String

        Get
            If Not BAPI Is Nothing Then
                InvoicingParty = BAPI.Exports("POHEADER").ParamValue(40)
            Else
                InvoicingParty = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue(40) = value
                BAPI.Exports("POHEADERX").ParamValue(40) = "X"
            End If
        End Set

    End Property

    Public Property PO_Incoterm() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Incoterm = BAPI.Exports("POHEADER").ParamValue("INCOTERMS1")
            Else
                PO_Incoterm = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("INCOTERMS1") = Left(value, 3)
                BAPI.Exports("POHEADERX").ParamValue("INCOTERMS1") = "X"
            End If
        End Set

    End Property

    Public Property PO_Incoterm_Desc() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Incoterm_Desc = BAPI.Exports("POHEADER").ParamValue("INCOTERMS2")
            Else
                PO_Incoterm_Desc = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("INCOTERMS2") = Left(value, 28)
                BAPI.Exports("POHEADERX").ParamValue("INCOTERMS2") = "X"
            End If
        End Set

    End Property

    Public Property PO_Currency() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Currency = BAPI.Exports("POHEADER").ParamValue("CURRENCY")
            Else
                PO_Currency = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("CURRENCY") = Left(value, 3).ToUpper
                BAPI.Exports("POHEADERX").ParamValue("CURRENCY") = "X"
            End If
        End Set

    End Property

    Public Property ItemQuantity(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POSCHEDULE", "PO_ITEM", Item)
                If Not I < 0 Then
                    ItemQuantity = BAPI.Tables("POSCHEDULE").Rows(I).Item("QUANTITY")
                Else
                    ItemQuantity = Nothing
                End If
            Else
                ItemQuantity = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POSCHEDULE", Item, TRow, TRowX)
                TRow("SCHED_LINE") = "1"
                TRow("QUANTITY") = value
                TRowX("SCHED_LINE") = "1"
                TRowX("QUANTITY") = "X"
            End If
        End Set

    End Property

    Public Property DeliveryDate(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POSCHEDULE", "PO_ITEM", Item)
                If Not I < 0 Then
                    DeliveryDate = BAPI.Tables("POSCHEDULE").Rows(I).Item("DELIVERY_DATE")
                Else
                    DeliveryDate = Nothing
                End If
            Else
                DeliveryDate = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POSCHEDULE", Item, TRow, TRowX)
                TRow("SCHED_LINE") = "1"
                TRow("DELIVERY_DATE") = ConversionUtils.NetDate2SAPDate(value)
                TRowX("SCHED_LINE") = "1"
                TRowX("DELIVERY_DATE") = "X"
            End If
        End Set

    End Property

    Public Property StatDeliveryDate(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POSCHEDULE", "PO_ITEM", Item)
                If Not I < 0 Then
                    StatDeliveryDate = BAPI.Tables("POSCHEDULE").Rows(I).Item("STAT_DATE")
                Else
                    StatDeliveryDate = Nothing
                End If
            Else
                StatDeliveryDate = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POSCHEDULE", Item, TRow, TRowX)
                TRow("SCHED_LINE") = "1"
                TRow("STAT_DATE") = ConversionUtils.NetDate2SAPDate(value)
                TRowX("SCHED_LINE") = "1"
                TRowX("STAT_DATE") = "X"
            End If
        End Set

    End Property

    Public Property TaxCode(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    TaxCode = BAPI.Tables("POITEM").Rows(I).Item("TAX_CODE")
                Else
                    TaxCode = Nothing
                End If
            Else
                TaxCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("TAX_CODE") = value.ToUpper
                TRowX("TAX_CODE") = "X"
            End If
        End Set

    End Property

    Public Property JurisdCode(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    JurisdCode = BAPI.Tables("POITEM").Rows(I).Item("TAXJURCODE")
                Else
                    JurisdCode = Nothing
                End If
            Else
                JurisdCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("TAXJURCODE") = value
                TRowX("TAXJURCODE") = "X"
            End If
        End Set

    End Property

    Public Property BrasNCMCode(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    BrasNCMCode = BAPI.Tables("POITEM").Rows(I).Item("BRAS_NBM")
                Else
                    BrasNCMCode = Nothing
                End If
            Else
                BrasNCMCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("BRAS_NBM") = value
                TRowX("BRAS_NBM") = "X"
            End If
        End Set

    End Property

    Public Property MaterialUsage(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    MaterialUsage = BAPI.Tables("POITEM").Rows(I).Item("MATL_USAGE")
                Else
                    MaterialUsage = Nothing
                End If
            Else
                MaterialUsage = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("MATL_USAGE") = value
                TRowX("MATL_USAGE") = "X"
            End If
        End Set

    End Property

    Public Property MaterialOrigin(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    MaterialOrigin = BAPI.Tables("POITEM").Rows(I).Item("MAT_ORIGIN")
                Else
                    MaterialOrigin = Nothing
                End If
            Else
                MaterialOrigin = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("MAT_ORIGIN") = value
                TRowX("MAT_ORIGIN") = "X"
            End If
        End Set

    End Property

    Public Property MaterialCategory(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    MaterialCategory = BAPI.Tables("POITEM").Rows(I).Item("INDUS3")
                Else
                    MaterialCategory = Nothing
                End If
            Else
                MaterialCategory = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("INDUS3") = value
                TRowX("INDUS3") = "X"
            End If
        End Set

    End Property

    Public Property ItemNetPrice(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    ItemNetPrice = BAPI.Tables("POITEM").Rows(I).Item("NET_PRICE")
                Else
                    ItemNetPrice = Nothing
                End If
            Else
                ItemNetPrice = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("NET_PRICE") = value
                TRowX("NET_PRICE") = "X"
            End If
        End Set

    End Property

    Public Property ItemPriceUnit(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    ItemPriceUnit = BAPI.Tables("POITEM").Rows(I).Item("PRICE_UNIT")
                Else
                    ItemPriceUnit = Nothing
                End If
            Else
                ItemPriceUnit = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("PRICE_UNIT") = value
                TRowX("PRICE_UNIT") = "X"
            End If
        End Set

    End Property

    Public Property ItemUnitOfMeasure(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    ItemUnitOfMeasure = BAPI.Tables("POITEM").Rows(I).Item("PO_UNIT")
                Else
                    ItemUnitOfMeasure = Nothing
                End If
            Else
                ItemUnitOfMeasure = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("PO_UNIT") = value.ToUpper
                TRowX("PO_UNIT") = "X"
            End If
        End Set

    End Property

    Public Property Item_OrderPriceUnit(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    Item_OrderPriceUnit = BAPI.Tables("POITEM").Rows(I).Item("ORDERPR_UN")
                Else
                    Item_OrderPriceUnit = Nothing
                End If
            Else
                Item_OrderPriceUnit = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("ORDERPR_UN") = value.ToUpper
                TRowX("ORDERPR_UN") = "X"
            End If
        End Set

    End Property

    Public Property GL_Account(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POACCOUNT", "PO_ITEM", Item)
                If Not I < 0 Then
                    GL_Account = BAPI.Tables("POACCOUNT").Rows(I).Item("GL_ACCOUNT")
                Else
                    GL_Account = Nothing
                End If
            Else
                GL_Account = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                Dim AN As String
                SetTableRow("POACCOUNT", Item, TRow, TRowX)
                If Len(value) < 10 Then
                    AN = Left("0000000000", 10 - Len(value)) & value
                Else
                    AN = value
                End If
                TRow("SERIAL_NO") = "1"
                TRow("GL_ACCOUNT") = AN
                TRowX("SERIAL_NO") = "1"
                TRowX("GL_ACCOUNT") = "X"
            End If
        End Set

    End Property

    Public Property InternalOrder(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POACCOUNT", "PO_ITEM", Item)
                If Not I < 0 Then
                    InternalOrder = BAPI.Tables("POACCOUNT").Rows(I).Item("ORDERID")
                Else
                    InternalOrder = Nothing
                End If
            Else
                InternalOrder = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                Dim AN As String
                SetTableRow("POACCOUNT", Item, TRow, TRowX)
                If Len(value) < 10 Then
                    AN = Left("0000000000", 10 - Len(value)) & value
                Else
                    AN = value
                End If
                TRow("SERIAL_NO") = "1"
                TRow("ORDERID") = AN
                TRowX("SERIAL_NO") = "1"
                TRowX("ORDERID") = "X"
            End If
        End Set

    End Property

    Public Property DeliveryCompleted(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("NO_MORE_GR") = "X" Then
                        DeliveryCompleted = True
                    Else
                        DeliveryCompleted = False
                    End If
                Else
                    DeliveryCompleted = False
                End If
            Else
                DeliveryCompleted = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("NO_MORE_GR") = "X"
                Else
                    TRow("NO_MORE_GR") = " "
                End If
                TRowX("NO_MORE_GR") = "X"
            End If
        End Set

    End Property

    Public Property GRNonValuated(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("GR_NON_VAL") = "X" Then
                        GRNonValuated = True
                    Else
                        GRNonValuated = False
                    End If
                Else
                    GRNonValuated = False
                End If
            Else
                GRNonValuated = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("GR_NON_VAL") = "X"
                Else
                    TRow("GR_NON_VAL") = " "
                End If
                TRowX("GR_NON_VAL") = "X"
            End If
        End Set

    End Property

    Public Property UnlimitedDelivery(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("UNLIMITED_DLV") = "X" Then
                        UnlimitedDelivery = True
                    Else
                        UnlimitedDelivery = False
                    End If
                Else
                    UnlimitedDelivery = False
                End If
            Else
                UnlimitedDelivery = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("UNLIMITED_DLV") = "X"
                Else
                    TRow("UNLIMITED_DLV") = " "
                End If
                TRowX("UNLIMITED_DLV") = "X"
            End If
        End Set

    End Property

    Public Property FinalInvoice(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("FINAL_INV") = "X" Then
                        FinalInvoice = True
                    Else
                        FinalInvoice = False
                    End If
                Else
                    FinalInvoice = False
                End If
            Else
                FinalInvoice = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("FINAL_INV") = "X"
                Else
                    TRow("FINAL_INV") = " "
                End If
                TRowX("FINAL_INV") = "X"
            End If
        End Set

    End Property

    Public Property GoodsReceiptInd(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("GR_IND") = "X" Then
                        GoodsReceiptInd = True
                    Else
                        GoodsReceiptInd = False
                    End If
                Else
                    GoodsReceiptInd = False
                End If
            Else
                GoodsReceiptInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("GR_IND") = "X"
                Else
                    TRow("GR_IND") = " "
                End If
                TRowX("GR_IND") = "X"
            End If
        End Set

    End Property

    Public Property InvoiceReceiptInd(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("IR_IND") = "X" Then
                        InvoiceReceiptInd = True
                    Else
                        InvoiceReceiptInd = False
                    End If
                Else
                    InvoiceReceiptInd = False
                End If
            Else
                InvoiceReceiptInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("IR_IND") = "X"
                Else
                    TRow("IR_IND") = " "
                End If
                TRowX("IR_IND") = "X"
            End If
        End Set

    End Property

    Public Property MatGroup(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    MatGroup = BAPI.Tables("POITEM").Rows(I).Item("MATL_GROUP")
                Else
                    MatGroup = Nothing
                End If
            Else
                MatGroup = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("MATL_GROUP") = value
                TRowX("MATL_GROUP") = "X"
            End If
        End Set

    End Property

    Public Property ShortText(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    ShortText = BAPI.Tables("POITEM").Rows(I).Item("SHORT_TEXT")
                Else
                    ShortText = Nothing
                End If
            Else
                ShortText = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("SHORT_TEXT") = value
                TRowX("SHORT_TEXT") = "X"
            End If
        End Set

    End Property

    Public Property Item_AcctAssCat(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    Item_AcctAssCat = BAPI.Tables("POITEM").Rows(I).Item("ACCTASSCAT")
                Else
                    Item_AcctAssCat = Nothing
                End If
            Else
                Item_AcctAssCat = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("ACCTASSCAT") = value
                TRowX("ACCTASSCAT") = "X"
            End If
        End Set

    End Property

    Public Property BlockInd(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("DELETE_IND") = "S" Then
                        BlockInd = True
                    Else
                        BlockInd = False
                    End If
                Else
                    BlockInd = False
                End If
            Else
                BlockInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("DELETE_IND") = "S"
                Else
                    TRow("DELETE_IND") = " "
                End If
                TRowX("DELETE_IND") = "X"
            End If
        End Set

    End Property

    Public Property DeletionInd(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("DELETE_IND") = "X" Then
                        DeletionInd = True
                    Else
                        DeletionInd = False
                    End If
                Else
                    DeletionInd = False
                End If
            Else
                DeletionInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("DELETE_IND") = "X"
                Else
                    TRow("DELETE_IND") = " "
                End If
                TRowX("DELETE_IND") = "X"
            End If
        End Set

    End Property

    Public Property OrderAck(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    OrderAck = BAPI.Tables("POITEM").Rows(I).Item("ACKNOWL_NO")
                Else
                    OrderAck = Nothing
                End If
            Else
                OrderAck = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("ACKNOWL_NO") = value
                TRowX("ACKNOWL_NO") = "X"
            End If
        End Set

    End Property

    Public Property AckRequired(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("ACKN_REQD") = "X" Then
                        AckRequired = True
                    Else
                        AckRequired = False
                    End If
                Else
                    AckRequired = False
                End If
            Else
                AckRequired = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("ACKN_REQD") = "X"
                Else
                    TRow("ACKN_REQD") = " "
                End If
                TRowX("ACKN_REQD") = "X"
            End If
        End Set

    End Property

    Public Property GRBasedIV(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("GR_BASEDIV") = "X" Then
                        GRBasedIV = True
                    Else
                        GRBasedIV = False
                    End If
                Else
                    GRBasedIV = False
                End If
            Else
                GRBasedIV = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("GR_BASEDIV") = "X"
                Else
                    TRow("GR_BASEDIV") = " "
                End If
                TRowX("GR_BASEDIV") = "X"
            End If
        End Set

    End Property

    Public Property Free(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("FREE_ITEM") = "X" Then
                        Free = True
                    Else
                        Free = False
                    End If
                Else
                    Free = False
                End If
            Else
                Free = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("FREE_ITEM") = "X"
                Else
                    TRow("FREE_ITEM") = " "
                End If
                TRowX("FREE_ITEM") = "X"
            End If
        End Set

    End Property

    Public Property ERS_Flag(ByVal Item As String) As Boolean

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    If BAPI.Tables("POITEM").Rows(I).Item("ERS") = "X" Then
                        ERS_Flag = True
                    Else
                        ERS_Flag = False
                    End If
                Else
                    ERS_Flag = False
                End If
            Else
                ERS_Flag = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                If value Then
                    TRow("ERS") = "X"
                Else
                    TRow("ERS") = " "
                End If
                TRowX("ERS") = "X"
            End If
        End Set

    End Property

    Public Property TrackingField(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    TrackingField = BAPI.Tables("POITEM").Rows(I).Item("TRACKINGNO")
                Else
                    TrackingField = Nothing
                End If
            Else
                TrackingField = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("TRACKINGNO") = value
                TRowX("TRACKINGNO") = "X"
            End If
        End Set

    End Property

    Public Property ConfControl(ByVal Item As String) As String

        Get
            If Not BAPI Is Nothing Then
                Dim I As Integer = GetItemIndex("POITEM", "PO_ITEM", Item)
                If Not I < 0 Then
                    ConfControl = BAPI.Tables("POITEM").Rows(I).Item("CONF_CTRL")
                Else
                    ConfControl = Nothing
                End If
            Else
                ConfControl = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POITEM", Item, TRow, TRowX)
                TRow("CONF_CTRL") = value
                TRowX("CONF_CTRL") = "X"
            End If
        End Set

    End Property

    Public Sub Item_Add_Condition(ByVal Item As String, ByVal Condition As String, ByVal Value As String, Optional ByVal Currency As Object = Nothing, _
                          Optional ByVal UOM As Object = Nothing, Optional ByVal Per As Object = Nothing)

        If BAPI Is Nothing Then Exit Sub
        If Value Is Nothing OrElse Value = "" Then Value = "0"
        If Not IsNumeric(Value) Then
            Sts = "Condition Value must be numeric"
            Exit Sub
        End If

        Item = Item.PadLeft(6, "0")
        Condition = Condition.ToUpper.Trim
        If Not Currency Is Nothing Then Currency = Currency.ToString.ToUpper.Trim
        If Not UOM Is Nothing Then UOM = UOM.ToString.ToUpper.Trim
        If Not Per Is Nothing Then Per = Per.ToString.ToUpper.Trim

        Dim TRow = BAPI.Tables("POCOND").AddRow
        Dim TRowX = BAPI.Tables("POCONDX").AddRow

        TRow("ITM_NUMBER") = Item.PadLeft(6, "0")
        TRowX("ITM_NUMBER") = Item.PadLeft(6, "0")
        TRow("COND_TYPE") = Condition
        TRowX("COND_TYPE") = "X"
        TRow("COND_VALUE") = CDbl(Value)
        TRowX("COND_VALUE") = "X"
        TRow("CHANGE_ID") = "I"
        TRowX("CHANGE_ID") = "X"

        If Not Currency Is Nothing AndAlso Not DBNull.Value.Equals(Currency) Then
            If Currency <> "" Then
                TRow("CURRENCY") = Currency
                TRowX("CURRENCY") = "X"
            End If
        End If

        If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) Then
            If UOM <> "" Then
                TRow("COND_UNIT") = UOM
                TRowX("COND_UNIT") = "X"
            End If
        End If

        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) Then
            If Per <> "" Then
                TRow("COND_P_UNT") = Per
                TRowX("COND_P_UNT") = "X"
            End If
        End If

    End Sub

    Public Sub Item_Update_Condition(ByVal Item As String, ByVal Condition As String, ByVal Value As String, Optional ByVal Currency As Object = Nothing, _
                          Optional ByVal UOM As Object = Nothing, Optional ByVal Per As Object = Nothing)

        If BAPI Is Nothing Then Exit Sub
        If Value Is Nothing OrElse Value = "" Then Value = "0"
        If Not IsNumeric(Value) Then
            Sts = "Condition Value must be numeric"
            Exit Sub
        End If

        Item = Item.PadLeft(6, "0")
        Condition = Condition.ToUpper.Trim
        If Not Currency Is Nothing Then Currency = Currency.ToString.ToUpper.Trim
        If Not UOM Is Nothing Then UOM = UOM.ToString.ToUpper.Trim
        If Not Per Is Nothing Then Per = Per.ToString.ToUpper.Trim

        If P1 Is Nothing Then
            P1 = New POInfo1(Con, PONum)
        End If

        If P1.Condition_Value(Item, Condition) Is Nothing Then Exit Sub

        Dim TRow = BAPI.Tables("POCOND").AddRow
        Dim TRowX = BAPI.Tables("POCONDX").AddRow

        TRow("ITM_NUMBER") = Item.PadLeft(6, "0")
        TRowX("ITM_NUMBER") = Item.PadLeft(6, "0")
        TRow("COND_TYPE") = Condition
        TRowX("COND_TYPE") = "X"
        TRow("CONDITION_NO") = P1.Condition_No(Item, Condition)
        TRowX("CONDITION_NO") = "X"
        TRow("COND_ST_NO") = P1.Condition_StepNo(Item, Condition)
        TRowX("COND_ST_NO") = "X"
        TRow("COND_COUNT") = P1.Condition_Count(Item, Condition)
        TRowX("COND_COUNT") = "X"
        TRow("COND_VALUE") = CDbl(Value)
        TRowX("COND_VALUE") = "X"

        TRow("NUMCONVERT") = "1"
        TRowX("NUMCONVERT") = "X"
        TRow("DENOMINATO") = "1"
        TRowX("DENOMINATO") = "X"

        TRow("CHANGE_ID") = "U"
        TRowX("CHANGE_ID") = "X"

        If Not Currency Is Nothing AndAlso Not DBNull.Value.Equals(Currency) Then
            If Currency <> "" Then
                TRow("CURRENCY") = Currency
                TRowX("CURRENCY") = "X"
            End If
        End If

        If Not UOM Is Nothing AndAlso Not DBNull.Value.Equals(UOM) Then
            If UOM <> "" Then
                TRow("COND_UNIT") = UOM
                TRowX("COND_UNIT") = "X"
            End If
        End If

        If Not Per Is Nothing AndAlso Not DBNull.Value.Equals(Per) Then
            If Per <> "" Then
                TRow("COND_P_UNT") = Per
                TRowX("COND_P_UNT") = "X"
            End If
        End If

    End Sub

    Public Sub Reverse_PriceQuantity(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        ItemNetPrice(Item) = PO.ItemQuantity(Item)
        ItemQuantity(Item) = PO.ItemNetPrice(Item)

    End Sub

    Public Sub CloseItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        FinalInvoice(Item) = True
        DeliveryCompleted(Item) = True

    End Sub

    Public Sub ReOpenItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        FinalInvoice(Item) = False
        DeliveryCompleted(Item) = False

    End Sub

    Public Sub ClosePO()

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        Dim Item As String
        For Each Item In PO.ItemNumbers()
            CloseItem(Item)
        Next

    End Sub

    Public Sub ReOpenPO()

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        Dim Item As String
        For Each Item In PO.ItemNumbers
            ReOpenItem(Item)
        Next

    End Sub

    Public Sub DeleteItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        FinalInvoice(Item) = True
        DeliveryCompleted(Item) = True
        DeletionInd(Item) = True

    End Sub

    Public Sub UnDeleteItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        FinalInvoice(Item) = False
        DeliveryCompleted(Item) = False
        DeletionInd(Item) = False

    End Sub

    Public Sub DeletePO()

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        Dim Item As String
        For Each Item In PO.ItemNumbers()
            DeleteItem(Item)
        Next

    End Sub

    Public Sub UnDeletePO()

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        Dim Item As String
        For Each Item In PO.ItemNumbers
            UnDeleteItem(Item)
        Next

    End Sub

    Public Sub NewHeaderText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim A As String()
        Dim I As Integer
        Dim TRow
        Dim SR As New StringReader(Text)
        Dim S As String

        While True
            S = SR.ReadLine
            If S Is Nothing Then
                Exit While
            Else
                A = BAPITextSplit(S)
                I = 1
                Do While I <= UBound(A)
                    TRow = BAPI.Tables("POTEXTHEADER").AddRow
                    TRow.Item("PO_NUMBER") = BAPI.Exports("PURCHASEORDER").ParamValue
                    TRow.Item("TEXT_ID") = TextID
                    If I = 1 Then
                        TRow.Item("TEXT_FORM") = "*"
                    Else
                        TRow.Item("TEXT_FORM") = ""
                    End If
                    TRow.Item("TEXT_LINE") = A(I)
                    I = I + 1
                Loop
            End If
        End While

    End Sub

    Public Sub InsertHeaderText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow
        Dim SR As New StringReader(Text)

        NewHeaderText(TextID, Text)

        If PO Is Nothing Then
            PO = New POInfo(BAPI.Connection, BAPI.Exports("PURCHASEORDER").ParamValue)
        End If

        Dim HTA As SAPText() = PO.HeaderTextArray(TextID, "PO_HEADER_TEXTS")
        Dim HT As SAPText
        If Not HTA Is Nothing Then
            For Each HT In HTA
                TRow = BAPI.Tables("POTEXTHEADER").AddRow
                TRow.Item("PO_NUMBER") = BAPI.Exports("PURCHASEORDER").ParamValue
                TRow.Item("TEXT_ID") = TextID
                TRow.Item("TEXT_FORM") = HT.Format
                TRow.Item("TEXT_LINE") = HT.Text
            Next
        End If

    End Sub

    Public Sub NewItemText(ByVal Item As String, ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim A As String()
        Dim I As Integer
        Dim TRow
        Dim SR As New StringReader(Text)
        Dim S As String

        While True
            S = SR.ReadLine
            If S Is Nothing Then
                Exit While
            Else
                A = BAPITextSplit(S)
                I = 1
                Do While I <= UBound(A)
                    TRow = BAPI.Tables("POTEXTITEM").AddRow
                    TRow.Item(1) = Item
                    TRow.Item(2) = TextID
                    If I = 1 Then
                        TRow.Item(3) = "*"
                    Else
                        TRow.Item(3) = ""
                    End If
                    TRow.Item(4) = A(I)
                    I = I + 1
                Loop
            End If
        End While

    End Sub

    Public Sub AppendRequisition(ByVal ReqNumber As String)

        If BAPI Is Nothing Then Exit Sub

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If
        Dim II As Integer = Val(PO.ItmInterval)
        Dim LI As Integer = Val(PO.LastItem)
        Dim Req As New PRInfo(Con, ReqNumber)
        Req.ReqNumber = ReqNumber
        Dim RI As String
        For Each RI In Req.ItemNumbers
            Dim TRow
            TRow = BAPI.Tables("POITEM").AddRow
            LI = LI + II
            TRow("PO_ITEM") = CStr(LI)
            TRow("PREQ_NO") = Left("0000000000", 10 - Len(ReqNumber)) & ReqNumber
            TRow("PREQ_ITEM") = RI
            Dim TRowX As RFCStructure = BAPI.Tables("POITEMX").AddRow
            TRowX("PO_ITEM") = CStr(LI)
            TRowX("PREQ_NO") = "X"
            TRowX("PREQ_ITEM") = "X"
        Next

    End Sub

    Public Sub IncreaseLineItem(ByVal Item As String, ByVal ReqNumber As String)

        If BAPI Is Nothing Then Exit Sub

        Dim IT As String = Nothing
        Dim Req As New PRInfo(Con, ReqNumber)

        If PO Is Nothing Then
            PO = New POInfo(Con, PONum)
        End If

        Dim SL As Integer = PO.LastSchedLine + 1
        Dim TRow
        Dim TRowX
        For Each IT In Req.ItemNumbers
            TRow = BAPI.Tables("POSCHEDULE").AddRow
            TRowX = BAPI.Tables("POSCHEDULEX").AddRow
            TRow("PO_ITEM") = Item
            TRow("SCHED_LINE") = CStr(SL)
            TRow("PREQ_NO") = Left("0000000000", 10 - Len(ReqNumber)) & ReqNumber
            TRow("PREQ_ITEM") = IT
            TRowX("PO_ITEM") = Item
            TRowX("SCHED_LINE") = CStr(SL)
            TRowX("PREQ_NO") = "X"
            TRowX("PREQ_ITEM") = "X"
            SL += 1
        Next

    End Sub

    Public Sub PO_Incoterms(ByVal Part1 As String, Optional ByVal Part2 As String = Nothing)

        PO_Incoterm = Part1
        If Not Part2 Is Nothing Then
            PO_Incoterm_Desc = Part2
        End If

    End Sub

End Class

Public NotInheritable Class POCreator1 : Inherits SC_BAPI_Base

    Private LIN As Integer = 0
    Private LII As Integer = 0
    Private PONum As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public ReadOnly Property PONumber() As String

        Get
            PONumber = PONum
        End Get

    End Property

    Public Property PO_Type() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Type = BAPI.Exports("PO_HEADER").ParamValue("DOC_TYPE")
            Else
                PO_Type = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("PO_HEADER").ParamValue("DOC_TYPE") = value
            End If
        End Set

    End Property

    Public Property Vendor() As String

        Get
            If Not BAPI Is Nothing Then
                Vendor = BAPI.Exports("PO_HEADER").ParamValue("VENDOR")
            Else
                Vendor = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("PO_HEADER").ParamValue("VENDOR") = value.PadLeft(10, "0")
                VC = value
            End If
        End Set

    End Property

    Public Property PurchGroup() As String

        Get
            If Not BAPI Is Nothing Then
                PurchGroup = BAPI.Exports("PO_HEADER").ParamValue("PUR_GROUP")
            Else
                PurchGroup = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("PO_HEADER").ParamValue("PUR_GROUP") = value
            End If
        End Set

    End Property

    Public Property PurchOrg() As String

        Get
            If Not BAPI Is Nothing Then
                PurchOrg = BAPI.Exports("PO_HEADER").ParamValue("PURCH_ORG")
            Else
                PurchOrg = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("PO_HEADER").ParamValue("PURCH_ORG") = value
            End If
        End Set

    End Property

    Public Property AccAsignmentCat() As String

        Get
            If Not BAPI Is Nothing Then
                AccAsignmentCat = BAPI.Tables("PO_ITEMS").Rows(LII)("ACCTASSCAT")
            Else
                AccAsignmentCat = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("PO_ITEMS").Rows(LII)("ACCTASSCAT") = value
            End If
        End Set

    End Property

    Public Property Plant() As String

        Get
            If Not BAPI Is Nothing Then
                Plant = BAPI.Tables("PO_ITEMS").Rows(LII)("PLANT")
            Else
                Plant = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("PO_ITEMS").Rows(LII)("PLANT") = value
                Dim T As New T001K_Report(Con)
                T.IncludePlant(value)
                T.Execute()
                If T.Success AndAlso T.Data.Rows.Count > 0 Then
                    CC = T.Data.Rows(0)("CCode")
                End If
            End If
        End Set

    End Property

    Public Property Material() As String

        Get
            If Not BAPI Is Nothing Then
                Material = BAPI.Tables("PO_ITEMS").Rows(LII)("PUR_MAT")
            Else
                Material = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("PO_ITEMS").Rows(LII)("PUR_MAT") = value.PadLeft(18, "0")
            End If
        End Set

    End Property

    Public Property ItemUOM() As String

        Get
            If Not BAPI Is Nothing Then
                ItemUOM = BAPI.Tables("PO_ITEMS").Rows(LII)("UNIT")
            Else
                ItemUOM = Nothing
            End If
        End Get
        Set(ByVal value As String)
            BAPI.Tables("PO_ITEMS").Rows(LII)("UNIT") = value.ToUpper
        End Set

    End Property

    Public Property ItemNetPrice() As String

        Get
            If Not BAPI Is Nothing Then
                ItemNetPrice = BAPI.Tables("PO_ITEMS").Rows(LII)("NET_PRICE")
            Else
                ItemNetPrice = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("PO_ITEMS").Rows(LII)("NET_PRICE") = value
            End If
        End Set

    End Property

    Public Property ItemShortText() As String

        Get
            If Not BAPI Is Nothing Then
                ItemShortText = BAPI.Tables("PO_ITEMS").Rows(LII)("SHORT_TEXT")
            Else
                ItemShortText = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("PO_ITEMS").Rows(LII)("SHORT_TEXT") = value
            End If
        End Set

    End Property

    Public Property ItemQuantity() As String

        Get
            If Not BAPI Is Nothing Then
                ItemQuantity = BAPI.Tables("Po_Item_Schedules").Rows(LII)("QUANTITY")
            Else
                ItemQuantity = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("Po_Item_Schedules", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("QUANTITY") = value
                BAPI.Tables("PO_ITEMS").Rows(LII)("DISP_QUAN") = value
            End If
        End Set

    End Property

    Public Property DeliveryDate() As String

        Get
            If Not BAPI Is Nothing Then
                DeliveryDate = BAPI.Tables("Po_Item_Schedules").Rows(LII)("DELIV_DATE")
            Else
                DeliveryDate = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("Po_Item_Schedules", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("DELIV_DATE") = ConversionUtils.NetDate2SAPDate(value)
            End If
        End Set

    End Property

    Public Property InternalOrder() As String

        Get
            If Not BAPI Is Nothing Then
                InternalOrder = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(LII)("ORDER_NO")
            Else
                InternalOrder = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("PO_ITEM_ACCOUNT_ASSIGNMENT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("ORDER_NO") = value.PadLeft(12, "0")
            End If
        End Set

    End Property

    Public Property GL_Account() As String

        Get
            If Not BAPI Is Nothing Then
                GL_Account = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(LII)("G_L_ACCT")
            Else
                GL_Account = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("PO_ITEM_ACCOUNT_ASSIGNMENT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("G_L_ACCT") = value.PadLeft(10, "0")
            End If
        End Set

    End Property

    Public Property CostCenter() As String

        Get
            If Not BAPI Is Nothing Then
                CostCenter = BAPI.Tables("PO_ITEM_ACCOUNT_ASSIGNMENT").Rows(LII)("COST_CTR")
            Else
                CostCenter = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("PO_ITEM_ACCOUNT_ASSIGNMENT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("COST_CTR") = value
            End If
        End Set

    End Property

    Public Sub CreateNewPO(ByVal Type As String)

        If Con Is Nothing OrElse Not Con.Ping Then Exit Sub

        Try
            GC.Collect()
            BAPI = Con.CreateBapi("PurchaseOrder", "CreateFromData")
            If Not Type Is Nothing Then
                If Type <> "" Then
                    PO_Type = Type
                End If
            End If
            LII = 0
            LIN = 0
            Errors = False
        Catch ex As Exception
            Sts = ex.Message
            BAPI = Nothing
            RF = False
        End Try

    End Sub

    Public Sub CreateBlankItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow
        LIN = Val(Item)
        TRow = BAPI.Tables("PO_ITEMS").AddRow
        TRow("PO_ITEM") = Item
        LII = BAPI.Tables("PO_ITEMS").Rows.Count - 1

    End Sub

    Public Overloads Sub CommitChanges()

        If BAPI Is Nothing Then Exit Sub

        If Not Check_Vendor_LE_Link() Then
            ReDim R(0)
            R(0) = "(E) Vendor not linked to Company Code"
            Errors = True
            BAPI = Nothing
            Exit Sub
        End If

        BAPI.Execute()
        GetResult()
        If Not Errors Then
            BAPI.CommitWork(True)
        End If
        BAPI = Nothing

    End Sub

    Friend Overrides Sub SetTableRow(ByVal TableName As String, ByVal Item As String, ByRef TRow As Object, ByRef TRowX As Object)

        Dim F As Boolean = False
        Dim I As Integer = 0
        Dim IL As String

        If Not Item.StartsWith("0") Then Item = Item.PadLeft(5, "0")

        IL = "PO_ITEM"

        If BAPI.Tables(TableName).Rows.Count > 0 Then
            I = 0
            Do While Not F And Not I = BAPI.Tables(TableName).Rows.Count
                If (BAPI.Tables(TableName).Rows(I).Item(IL) = Item) Then
                    F = True
                Else
                    I += 1
                End If
            Loop
        End If

        If Not F Then
            TRow = BAPI.Tables(TableName).AddRow
        Else
            TRow = BAPI.Tables(TableName).Item(I)
        End If

        TRow(IL) = Item

    End Sub

    Friend Sub GetResult()

        Dim TRow
        Dim S As String
        Dim A() As String = Nothing
        Dim I As Integer = 0

        For Each TRow In BAPI.Tables("RETURN").Rows
            If TRow.Item("Type") = "E" Then
                Errors = True
            End If
            If Not TRow.Item("Message") Like "Error transferring ExtensionIn*" _
            And Not TRow.Item("Message") Like "*No instance of object type*" _
            And Not TRow.item("Message") Like "*could not be changed" Then
                S = "(" & TRow.Item("Type") & ") " & TRow.Item("Message")
                If A Is Nothing OrElse Not A.Contains(S) Then
                    ReDim Preserve A(I)
                    A(I) = S
                    I += 1
                End If
            End If
            If TRow.Item("CODE") = "06017" Then
                PONum = TRow.Item("MESSAGE_V2")
            End If
        Next
        If BAPI.Tables("RETURN").Rows.Count > 0 Then BRT = BAPI.Tables("RETURN").ToADOTable()
        R = A

    End Sub

End Class

Public NotInheritable Class POCreator : Inherits SC_BAPI_Base

    Private Req As PRInfo = Nothing
    Private LIN As Integer = 0
    Private LII As Integer = 0
    Private PONum As String = Nothing
    Private TaxPP As Boolean = False
    Private Taxes() As String = Nothing
    Private JurPP As Boolean = False
    Private JurCo() As String = Nothing
    Private ItmItvl As Integer = -1
    Private URDD As Boolean = False

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Public Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public ReadOnly Property PONumber() As String

        Get
            PONumber = PONum
        End Get

    End Property

    Public Property UpdateRealisticDeliveryDates() As Boolean

        Get
            UpdateRealisticDeliveryDates = URDD
        End Get

        Set(ByVal value As Boolean)
            URDD = value
        End Set

    End Property

    Public Property GenerateOutput() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Exports("NO_MESSAGING").ParamValue = "X" Then
                    GenerateOutput = False
                Else
                    GenerateOutput = True
                End If
            Else
                GenerateOutput = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Exports("NO_MESSAGING").ParamValue = " "
                    BAPI.Exports("NO_MESSAGE_REQ").ParamValue = " "
                Else
                    BAPI.Exports("NO_MESSAGING").ParamValue = "X"
                    BAPI.Exports("NO_MESSAGE_REQ").ParamValue = "X"
                End If
            End If
        End Set

    End Property

    Public Property ItemInterval() As Integer

        Get
            ItemInterval = ItmItvl
        End Get

        Set(ByVal value As Integer)
            If value <> 0 Then
                ItmItvl = value
            End If
        End Set

    End Property

    Public Property PO_Type() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Type = BAPI.Exports("POHEADER").ParamValue("DOC_TYPE")
            Else
                PO_Type = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("DOC_TYPE") = value
                BAPI.Exports("POHEADERX").ParamValue("DOC_TYPE") = "X"
            End If
        End Set

    End Property

    Public Property Vendor() As String

        Get
            If Not BAPI Is Nothing Then
                Vendor = BAPI.Exports("POHEADER").ParamValue("VENDOR")
            Else
                Vendor = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("VENDOR") = Left("0000000000", 10 - Len(value)) & value
                BAPI.Exports("POHEADERX").ParamValue("VENDOR") = "X"
                VC = value
            End If
        End Set

    End Property

    Public Property PO_Currency() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Currency = BAPI.Exports("POHEADER").ParamValue("CURRENCY")
            Else
                PO_Currency = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("CURRENCY") = Left(value, 3).ToUpper
            End If
        End Set

    End Property

    Public Property PurchGroup() As String

        Get
            If Not BAPI Is Nothing Then
                PurchGroup = BAPI.Exports("POHEADER").ParamValue("PUR_GROUP")
            Else
                PurchGroup = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("PUR_GROUP") = value
                BAPI.Exports("POHEADERX").ParamValue("PUR_GROUP") = "X"
            End If
        End Set

    End Property

    Public Property PurchOrg() As String

        Get
            If Not BAPI Is Nothing Then
                PurchOrg = BAPI.Exports("POHEADER").ParamValue("PURCH_ORG")
            Else
                PurchOrg = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("PURCH_ORG") = value
                BAPI.Exports("POHEADERX").ParamValue("PURCH_ORG") = "X"
            End If
        End Set

    End Property

    Public Property CompanyCode() As String

        Get
            If Not BAPI Is Nothing Then
                CompanyCode = BAPI.Exports("POHEADER").ParamValue("COMP_CODE")
            Else
                CompanyCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("COMP_CODE") = value
                BAPI.Exports("POHEADERX").ParamValue("COMP_CODE") = "X"
                CC = value
            End If
        End Set

    End Property

    Public Property OurReference() As String

        Get
            If Not BAPI Is Nothing Then
                OurReference = BAPI.Exports("POHEADER").ParamValue("OUR_REF")
            Else
                OurReference = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("OUR_REF") = value
                BAPI.Exports("POHEADERX").ParamValue("OUR_REF") = "X"
            End If
        End Set

    End Property

    Public Property YourReference() As String

        Get
            If Not BAPI Is Nothing Then
                YourReference = BAPI.Exports("POHEADER").ParamValue("REF_1")
            Else
                YourReference = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("REF_1") = value
                BAPI.Exports("POHEADERX").ParamValue("REF_1") = "X"
            End If
        End Set

    End Property

    Public Property PmntTerms() As String

        Get
            If Not BAPI Is Nothing Then
                PmntTerms = BAPI.Exports("POHEADER").ParamValue("PMNTTRMS")
            Else
                PmntTerms = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("PMNTTRMS") = value
                BAPI.Exports("POHEADERX").ParamValue("PMNTTRMS") = "X"
            End If
        End Set

    End Property

    Public Property PO_Incoterm() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Incoterm = BAPI.Exports("POHEADER").ParamValue("INCOTERMS1")
            Else
                PO_Incoterm = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("INCOTERMS1") = Left(value, 3)
                BAPI.Exports("POHEADERX").ParamValue("INCOTERMS1") = "X"
            End If
        End Set

    End Property

    Public Property PO_Incoterm_Desc() As String

        Get
            If Not BAPI Is Nothing Then
                PO_Incoterm_Desc = BAPI.Exports("POHEADER").ParamValue("INCOTERMS2")
            Else
                PO_Incoterm_Desc = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Exports("POHEADER").ParamValue("INCOTERMS2") = Left(value, 28)
                BAPI.Exports("POHEADERX").ParamValue("INCOTERMS2") = "X"
            End If
        End Set

    End Property

    Public Property Storage_Loc() As String

        Get
            If Not BAPI Is Nothing Then
                Storage_Loc = BAPI.Tables("POITEM").Rows(LII)("STGE_LOC")
            Else
                Storage_Loc = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("STGE_LOC") = value
                BAPI.Tables("POITEMX").Rows(LII)("STGE_LOC") = "X"
            End If
        End Set

    End Property

    Public Property AckRequired() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("ACKN_REQD") = "X" Then
                    AckRequired = True
                Else
                    AckRequired = False
                End If
            Else
                AckRequired = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("ACKN_REQD") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("ACKN_REQD") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("ACKN_REQD") = "X"
            End If
        End Set

    End Property

    Public Property Plant() As String

        Get
            If Not BAPI Is Nothing Then
                Plant = BAPI.Tables("POITEM").Rows(LII)("PLANT")
            Else
                Plant = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("PLANT") = value
                BAPI.Tables("POITEMX").Rows(LII)("PLANT") = "X"
                Dim T As New T001K_Report(Con)
                T.IncludePlant(value)
                T.Execute()
                If T.Success AndAlso T.Data.Rows.Count > 0 Then
                    CC = T.Data.Rows(0)("CCode")
                End If
            End If
        End Set

    End Property

    Public Property AccAsignmentCat() As String

        Get
            If Not BAPI Is Nothing Then
                AccAsignmentCat = BAPI.Tables("POITEM").Rows(LII)("ACCTASSCAT")
            Else
                AccAsignmentCat = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("ACCTASSCAT") = value
                BAPI.Tables("POITEMX").Rows(LII)("ACCTASSCAT") = "X"
            End If
        End Set

    End Property

    Public Property ItemUOM() As String

        Get
            If Not BAPI Is Nothing Then
                ItemUOM = BAPI.Tables("POITEM").Rows(LII)("PO_UNIT")
            Else
                ItemUOM = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("PO_UNIT") = value
                BAPI.Tables("POITEMX").Rows(LII)("PO_UNIT") = "X"
            End If
        End Set

    End Property

    Public Property ItemOPU() As String

        Get
            If Not BAPI Is Nothing Then
                ItemOPU = BAPI.Tables("POITEM").Rows(LII)("ORDERPR_UN")
            Else
                ItemOPU = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("ORDERPR_UN") = value
                BAPI.Tables("POITEMX").Rows(LII)("ORDERPR_UN") = "X"
            End If
        End Set

    End Property

    Public Property FreeItem() As Boolean '

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("FREE_ITEM") = "X" Then
                    FreeItem = True
                Else
                    FreeItem = False
                End If
            Else
                FreeItem = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("FREE_ITEM") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("FREE_ITEM") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("FREE_ITEM") = "X"
            End If
        End Set

    End Property

    Public Property GoodsReceiptInd() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("GR_IND") = "X" Then
                    GoodsReceiptInd = True
                Else
                    GoodsReceiptInd = False
                End If
            Else
                GoodsReceiptInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("GR_IND") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("GR_IND") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("GR_IND") = "X"
            End If
        End Set

    End Property

    Public Property Unlimited() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("UNLIMITED_DLV") = "X" Then
                    Unlimited = True
                Else
                    Unlimited = False
                End If
            Else
                Unlimited = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("UNLIMITED_DLV") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("UNLIMITED_DLV") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("UNLIMITED_DLV") = "X"
            End If
        End Set

    End Property

    Public Property InvReceiptInd() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("IR_IND") = "X" Then
                    InvReceiptInd = True
                Else
                    InvReceiptInd = False
                End If
            Else
                InvReceiptInd = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("IR_IND") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("IR_IND") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("IR_IND") = "X"
            End If
        End Set

    End Property

    Public Property FinalInvoice() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("IR_IND") = "X" Then
                    FinalInvoice = True
                Else
                    FinalInvoice = False
                End If
            Else
                FinalInvoice = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("FINAL_INV") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("FINAL_INV") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("FINAL_INV") = "X"
            End If
        End Set

    End Property

    Public Property ERS() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("ERS") = "X" Then
                    ERS = True
                Else
                    ERS = False
                End If
            Else
                ERS = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("ERS") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("ERS") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("ERS") = "X"
            End If
        End Set

    End Property

    Public Property GRBasedIV() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("GR_BASEDIV") = "X" Then
                    GRBasedIV = True
                Else
                    GRBasedIV = False
                End If
            Else
                GRBasedIV = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("GR_BASEDIV") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("GR_BASEDIV") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("GR_BASEDIV") = "X"
            End If
        End Set

    End Property

    Public Property PrintPrice() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("PRNT_PRICE") = "X" Then
                    PrintPrice = True
                Else
                    PrintPrice = False
                End If
            Else
                PrintPrice = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("PRNT_PRICE") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("PRNT_PRICE") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("PRNT_PRICE") = "X"
            End If
        End Set

    End Property

    Public Property EstimatedPrice() As Boolean

        Get
            If Not BAPI Is Nothing Then
                If BAPI.Tables("POITEM").Rows(LII)("EST_PRICE") = "X" Then
                    EstimatedPrice = True
                Else
                    EstimatedPrice = False
                End If
            Else
                EstimatedPrice = False
            End If
        End Get

        Set(ByVal value As Boolean)
            If Not BAPI Is Nothing Then
                If value Then
                    BAPI.Tables("POITEM").Rows(LII)("EST_PRICE") = "X"
                Else
                    BAPI.Tables("POITEM").Rows(LII)("EST_PRICE") = " "
                End If
                BAPI.Tables("POITEMX").Rows(LII)("EST_PRICE") = "X"
            End If
        End Set

    End Property

    Public Property ItemShortText() As String

        Get
            If Not BAPI Is Nothing Then
                ItemShortText = BAPI.Tables("POITEM").Rows(LII)("SHORT_TEXT")
            Else
                ItemShortText = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("SHORT_TEXT") = value
                BAPI.Tables("POITEMX").Rows(LII)("SHORT_TEXT") = "X"
            End If
        End Set

    End Property

    Public Property ConfControl() As String

        Get
            If Not BAPI Is Nothing Then
                ConfControl = BAPI.Tables("POITEM").Rows(LII)("CONF_CTRL")
            Else
                ConfControl = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("CONF_CTRL") = value
                BAPI.Tables("POITEMX").Rows(LII)("CONF_CTRL") = "X"
            End If
        End Set

    End Property

    Public Property OverDelTolerance() As String

        Get
            If Not BAPI Is Nothing Then
                OverDelTolerance = BAPI.Tables("POITEM").Rows(LII)("OVER_DLV_TOL")
            Else
                OverDelTolerance = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("OVER_DLV_TOL") = value
                BAPI.Tables("POITEMX").Rows(LII)("OVER_DLV_TOL") = "X"
            End If
        End Set

    End Property

    Public Property MatGroup() As String

        Get
            If Not BAPI Is Nothing Then
                MatGroup = BAPI.Tables("POITEM").Rows(LII)("MATL_GROUP")
            Else
                MatGroup = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("MATL_GROUP") = value
                BAPI.Tables("POITEMX").Rows(LII)("MATL_GROUP") = "X"
            End If
        End Set

    End Property

    Public Property BrasNCMCode() As String

        Get
            If Not BAPI Is Nothing Then
                BrasNCMCode = BAPI.Tables("POITEM").Rows(LII)("BRAS_NBM")
            Else
                BrasNCMCode = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("BRAS_NBM") = value
                BAPI.Tables("POITEMX").Rows(LII)("BRAS_NBM") = "X"
            End If
        End Set

    End Property

    Public Property Material() As String

        Get
            If Not BAPI Is Nothing Then
                Material = BAPI.Tables("POITEM").Rows(LII)("MATERIAL")
            Else
                Material = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("Material") = Left("000000000000000000", 18 - Len(value)) & value
                BAPI.Tables("POITEMX").Rows(LII)("Material") = "X"
            End If
        End Set

    End Property

    Public Property MatlUsage() As String

        Get
            If Not BAPI Is Nothing Then
                MatlUsage = BAPI.Tables("POITEM").Rows(LII)("MATL_USAGE")
            Else
                MatlUsage = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("MATL_USAGE") = value
                BAPI.Tables("POITEMX").Rows(LII)("MATL_USAGE") = "X"
            End If
        End Set

    End Property

    Public Property MatlOrigin() As String

        Get
            If Not BAPI Is Nothing Then
                MatlOrigin = BAPI.Tables("POITEM").Rows(LII)("MAT_ORIGIN")
            Else
                MatlOrigin = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("MAT_ORIGIN") = value
                BAPI.Tables("POITEMX").Rows(LII)("MAT_ORIGIN") = "X"
            End If
        End Set

    End Property

    Public Property MatlCategory() As String

        Get
            If Not BAPI Is Nothing Then
                MatlCategory = BAPI.Tables("POITEM").Rows(LII)("INDUS3")
            Else
                MatlCategory = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("INDUS3") = value
                BAPI.Tables("POITEMX").Rows(LII)("INDUS3") = "X"
            End If
        End Set

    End Property

    Public Property ItemNetPrice() As String

        Get
            If Not BAPI Is Nothing Then
                ItemNetPrice = BAPI.Tables("POITEM").Rows(LII)("NET_PRICE")
            Else
                ItemNetPrice = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("NET_PRICE") = value
                BAPI.Tables("POITEMX").Rows(LII)("NET_PRICE") = "X"
            End If
        End Set

    End Property

    Public Property ItemCategory() As String

        Get
            If Not BAPI Is Nothing Then
                ItemCategory = BAPI.Tables("POITEM").Rows(LII)("ITEM_CAT")
            Else
                ItemCategory = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("ITEM_CAT") = value
                BAPI.Tables("POITEMX").Rows(LII)("ITEM_CAT") = "X"
            End If
        End Set

    End Property

    Public Property TrackingField() As String

        Get
            If Not BAPI Is Nothing Then
                TrackingField = BAPI.Tables("POITEM").Rows(LII)("TRACKINGNO")
            Else
                TrackingField = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                BAPI.Tables("POITEM").Rows(LII)("TRACKINGNO") = value
                BAPI.Tables("POITEMX").Rows(LII)("TRACKINGNO") = "X"
            End If
        End Set

    End Property

    Public Property ItemQuantity() As String

        Get
            If Not BAPI Is Nothing Then
                ItemQuantity = BAPI.Tables("POSCHEDULE").Rows(LII)("QUANTITY")
            Else
                ItemQuantity = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POSCHEDULE", CStr(LIN), TRow, TRowX)
                TRow("SCHED_LINE") = "1"
                TRow("QUANTITY") = value
                TRowX("SCHED_LINE") = "1"
                TRowX("QUANTITY") = "X"
            End If
        End Set

    End Property

    Public WriteOnly Property Currency(Optional ByVal Price As String = Nothing, Optional ByVal Per As String = Nothing) As String

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then

                Dim TRow = BAPI.Tables("POCOND").AddRow
                Dim TRowX = BAPI.Tables("POCONDX").AddRow
                TRow("ITM_NUMBER") = CStr(LIN)
                TRow("COND_TYPE") = "PBXX"
                TRow("CURRENCY") = value
                TRowX("ITM_NUMBER") = CStr(LIN)
                TRowX("COND_TYPE") = "X"
                TRowX("CURRENCY") = "X"
                TRowX("CHANGE_ID") = "X"
                If Not Price Is Nothing Then
                    TRow("COND_VALUE") = Price
                    TRowX("COND_VALUE") = "X"
                End If
                If Not Per Is Nothing Then
                    TRow("COND_P_UNT") = Per
                    TRowX("COND_P_UNT") = "X"
                End If
                TRow("CHANGE_ID") = "U"

                TRow = BAPI.Tables("POCOND").AddRow
                TRowX = BAPI.Tables("POCONDX").AddRow
                TRow("ITM_NUMBER") = CStr(LIN)
                TRow("COND_TYPE") = "PB00"
                TRow("CURRENCY") = value
                TRowX("ITM_NUMBER") = CStr(LIN)
                TRowX("COND_TYPE") = "X"
                TRowX("CURRENCY") = "X"
                TRowX("CHANGE_ID") = "X"
                If Not Price Is Nothing Then
                    TRow("COND_VALUE") = Price
                    TRowX("COND_VALUE") = "X"
                End If
                If Not Per Is Nothing Then
                    TRow("COND_P_UNT") = Per
                    TRowX("COND_P_UNT") = "X"
                End If
                TRow("CHANGE_ID") = "U"

                BAPI.Exports("POHEADER").ParamValue("CURRENCY") = value
                BAPI.Exports("POHEADERX").ParamValue("CURRENCY") = "X"

            End If
        End Set

    End Property

    Public WriteOnly Property ZHC3_Cond_Price(ByVal Currency As String, ByVal Per As String) As String

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = BAPI.Tables("POCOND").AddRow
                Dim TRowX = BAPI.Tables("POCONDX").AddRow
                TRow("ITM_NUMBER") = CStr(LIN)
                TRow("COND_TYPE") = "ZHC3"
                TRow("COND_VALUE") = value
                TRow("CHANGE_ID") = "U"
                TRow("CURRENCY") = Currency
                TRow("COND_P_UNT") = Per
                TRowX("ITM_NUMBER") = CStr(LIN)
                TRowX("COND_TYPE") = "X"
                TRowX("COND_VALUE") = "X"
                TRowX("CHANGE_ID") = "X"
                TRowX("COND_P_UNT") = "X"
                TRowX("CURRENCY") = "X"
            End If
        End Set

    End Property

    Public Property DeliveryDate() As String

        Get
            If Not BAPI Is Nothing Then
                DeliveryDate = BAPI.Tables("POSCHEDULE").Rows(LII)("DELIVERY_DATE")
            Else
                DeliveryDate = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POSCHEDULE", CStr(LIN), TRow, TRowX)
                TRow("SCHED_LINE") = "1"
                TRow("DELIVERY_DATE") = ConversionUtils.NetDate2SAPDate(value)
                TRow("STAT_DATE") = ConversionUtils.NetDate2SAPDate(value)
                TRowX("SCHED_LINE") = "1"
                TRowX("DELIVERY_DATE") = "X"
                TRow("STAT_DATE") = "X"
            End If
        End Set

    End Property

    Public Property InternalOrder() As String

        Get
            If Not BAPI Is Nothing Then
                InternalOrder = BAPI.Tables("POACCOUNT").Rows(LII)("ORDERID")
            Else
                InternalOrder = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POACCOUNT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("ORDERID") = value.PadLeft(12, "0")
                TRowX("SERIAL_NO") = "1"
                TRowX("ORDERID") = "X"
            End If
        End Set

    End Property

    Public Property GL_Account() As String

        Get
            If Not BAPI Is Nothing Then
                GL_Account = BAPI.Tables("POACCOUNT").Rows(LII)("GL_ACCOUNT")
            Else
                GL_Account = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POACCOUNT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("GL_ACCOUNT") = value.PadLeft(10, "0")
                TRowX("SERIAL_NO") = "1"
                TRowX("GL_ACCOUNT") = "X"
            End If
        End Set

    End Property

    Public Property CostCenter() As String

        Get
            If Not BAPI Is Nothing Then
                CostCenter = BAPI.Tables("POACCOUNT").Rows(LII)("COSTCENTER")
            Else
                CostCenter = Nothing
            End If
        End Get

        Set(ByVal value As String)
            If Not BAPI Is Nothing Then
                Dim TRow = Nothing
                Dim TRowX = Nothing
                SetTableRow("POACCOUNT", CStr(LIN), TRow, TRowX)
                TRow("SERIAL_NO") = "1"
                TRow("COSTCENTER") = value
                TRowX("SERIAL_NO") = "1"
                TRowX("COSTCENTER") = "X"
            End If
        End Set

    End Property

    Public Sub CreateNewPO(ByVal Type As String)

        If Con Is Nothing OrElse Not Con.Ping Then Exit Sub

        Try
            GC.Collect()
            BAPI = Con.CreateBapi("PurchaseOrder", "CreateFromData1")
            If Not Type Is Nothing Then
                If Type <> "" Then
                    PO_Type = Type
                End If
            End If
            LII = 0
            LIN = 0
            Req = Nothing
            Errors = False
            TaxPP = False
            Taxes = Nothing
            JurPP = False
            JurCo = Nothing
        Catch ex As Exception
            Sts = ex.Message
            BAPI = Nothing
            RF = False
        End Try

    End Sub

    Public Sub ChangeTaxCode(ByVal Value As String)

        If BAPI Is Nothing Then Exit Sub

        If Not TaxPP Then
            TaxPP = True
            ReDim Taxes(0)
        Else
            ReDim Preserve Taxes(UBound(Taxes) + 1)
        End If

        Taxes(UBound(Taxes)) = CStr(LIN) & "," & Value

    End Sub

    Public Sub ChangeJurisCode(ByVal Value As String)

        If BAPI Is Nothing Then Exit Sub

        If Not JurPP Then
            JurPP = True
            ReDim JurCo(0)
        Else
            ReDim Preserve JurCo(UBound(JurCo) + 1)
        End If

        JurCo(UBound(JurCo)) = CStr(LIN) & "," & Value

    End Sub

    Public Sub Reverse_PriceQuantity()

        If BAPI Is Nothing Then Exit Sub

        Dim I As Integer = 0
        Dim RI As String = BAPI.Tables("POITEM").Rows(LII)("PREQ_ITEM")
        If Req Is Nothing Then
            Req = New PRInfo(Con, BAPI.Tables("POITEM").Rows(LII)("PREQ_NO"))
        End If
        ItemNetPrice = Req.ItemQuantity(RI)
        ItemQuantity = Req.ItemNetPrice(RI)

    End Sub

    Public Sub CreateBlankItem(ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        Dim TRow
        Dim TRowX

        LIN = Val(Item)
        TRow = BAPI.Tables("POITEM").AddRow
        TRow("PO_ITEM") = Item
        TRowX = BAPI.Tables("POITEMX").AddRow
        TRowX("PO_ITEM") = Item
        LII = BAPI.Tables("POITEM").Rows.Count - 1

    End Sub

    Public Sub AdoptReqItem(ByVal ReqNumber As String, ByVal Item As String)

        If BAPI Is Nothing Then Exit Sub

        If Req Is Nothing OrElse Req.ReqNumber <> ReqNumber Then
            Req = New PRInfo(Con, ReqNumber)
        End If

        Dim II As Integer = 10
        If Req.IsReady Then
            If ItmItvl = -1 Then
                II = Val(Req.ItemInterval)
            Else
                II = ItmItvl
            End If
        End If

        Dim TRow
        LIN = LIN + II
        TRow = BAPI.Tables("POITEM").AddRow
        TRow("PO_ITEM") = CStr(LIN)
        TRow("PREQ_NO") = ReqNumber.PadLeft(10, "0")
        TRow("PREQ_ITEM") = Item
        Dim TRowX As RFCStructure = BAPI.Tables("POITEMX").AddRow
        TRowX("PO_ITEM") = CStr(LIN)
        TRowX("PREQ_NO") = "X"
        TRowX("PREQ_ITEM") = "X"
        LII = BAPI.Tables("POITEM").Rows.Count - 1

        If Req.IsReady Then
            Plant = Req.Plant(Item)
        End If

    End Sub

    Public Sub AppendHeaderText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim A As String()
        Dim I As Integer
        Dim TRow
        Dim SR As New StringReader(Text)
        Dim S As String

        While True
            S = SR.ReadLine
            If S Is Nothing Then
                Exit While
            Else
                A = BAPITextSplit(S)
                I = 1
                Do While I <= UBound(A)
                    TRow = BAPI.Tables("POTEXTHEADER").AddRow
                    TRow.Item(2) = TextID
                    If I = 1 Then
                        TRow.Item(3) = "*"
                    Else
                        TRow.Item(3) = ""
                    End If
                    TRow.Item(4) = A(I)
                    I = I + 1
                Loop
            End If
        End While

    End Sub

    Public Sub AppendItemText(ByVal TextID As String, ByVal Text As String)

        If BAPI Is Nothing Then Exit Sub

        Dim A As String()
        Dim I As Integer
        Dim TRow
        Dim SR As New StringReader(Text)
        Dim S As String

        While True
            S = SR.ReadLine
            If S Is Nothing Then
                Exit While
            Else
                A = BAPITextSplit(S)
                I = 1
                Do While I <= UBound(A)
                    TRow = BAPI.Tables("POTEXTITEM").AddRow
                    TRow.Item(1) = CStr(LIN)
                    TRow.Item(2) = TextID
                    If I = 1 Then
                        TRow.Item(3) = "*"
                    Else
                        TRow.Item(3) = ""
                    End If
                    TRow.Item(4) = A(I)
                    I = I + 1
                Loop
            End If
        End While

    End Sub

    Public Sub PO_Incoterms(ByVal Part1 As String, Optional ByVal Part2 As String = Nothing)

        PO_Incoterm = Part1
        If Not Part2 Is Nothing Then
            PO_Incoterm_Desc = Part2
        End If

    End Sub

    Public Overloads Sub CommitChanges()

        If BAPI Is Nothing Then Exit Sub

        If Not Check_Vendor_LE_Link() Then
            ReDim R(0)
            R(0) = "(E) Vendor not linked to Company Code"
            BAPI = Nothing
            Errors = True
            Exit Sub
        End If

        If BAPI.Exports("POHEADERX").ParamValue(2) <> "X" Then
            Req = New PRInfo(Con, CDbl(BAPI.Tables("POITEM").Item(0)("PREQ_NO")))
            BAPI.Exports("POHEADER").ParamValue(2) = Req.OrderType
            BAPI.Exports("POHEADERX").ParamValue(2) = "X"
        End If

        BAPI.Execute()
        GetResults()
        If Sts Like "There are no*" Then
            ReDim Preserve R(R.Length)
            R(R.GetUpperBound(0)) = "IMPORTANT! :" & Sts
        End If
        PONum = Nothing
        If Not Errors Then
            BAPI.CommitWork(True)
            PONum = BAPI.Imports("ExpPurchaseOrder").ParamValue
            Dim POC As POChanges = Nothing
            If TaxPP Then
                POC = New POChanges(Con, PONum)
                Dim I As Integer
                Dim A
                For I = 0 To UBound(Taxes)
                    A = Split(Taxes(I), ",")
                    POC.TaxCode(A(0)) = A(1)
                Next
            End If
            If JurPP Then
                If POC Is Nothing Then POC = New POChanges(Con, PONum)
                Dim I As Integer
                Dim A
                For I = 0 To UBound(JurCo)
                    A = Split(JurCo(I), ",")
                    POC.JurisdCode(A(0)) = A(1)
                Next
            End If
            If URDD Then
                Dim TRow
                For Each TRow In BAPI.Tables("RETURN").Rows
                    If TRow("MESSAGE") Like "*Realistic delivery date:*" Then
                        If POC Is Nothing Then POC = New POChanges(Con, PONum)
                        POC.StatDeliveryDate(CStr(Val(BAPI.Tables("POSCHEDULE")(CInt(TRow("ROW")) - 1)("PO_ITEM")))) = TRow("MESSAGE_V1")
                        POC.DeliveryDate(CStr(Val(BAPI.Tables("POSCHEDULE")(CInt(TRow("ROW")) - 1)("PO_ITEM")))) = TRow("MESSAGE_V1")
                    End If
                Next
            End If
            If Not POC Is Nothing Then
                POC.GenerateOutput = False
                POC.CommitChanges()
            End If
        End If
        BAPI = Nothing

    End Sub

End Class

#End Region

#Region "SAP Tables Reports"

Public MustInherit Class RTable_Report : Inherits LINQ_Support

    Friend T As RTable = Nothing
    Friend Con As R3Connection = Nothing
    Friend A(,) As String = Nothing
    Friend EM As String = Nothing
    Friend SF As Boolean = False
    Friend RF As Boolean = False

    Private CF(,) As String = Nothing

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim SC As New SAPConnector
        Con = SC.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            T = New RTable(Con)
            RF = True
        Else
            EM = SC.Status
        End If

    End Sub

    Public Sub New(ByVal Connection)

        If Not Connection Is Nothing AndAlso Connection.Ping Then
            Con = Connection
            T = New RTable(Con)
            RF = True
        Else
            EM = "Connection already closed"
        End If

    End Sub

    Public Overridable ReadOnly Property Data() As DataTable

        Get
            If Not T Is Nothing Then
                Data = T.Result
            Else
                Data = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property DataArray() As String(,)

        Get
            If Not Data Is Nothing AndAlso A Is Nothing Then
                Dim R As DataRow
                Dim X As Integer
                Dim I As Integer
                ReDim A(Data.Rows.Count, Data.Columns.Count)
                For I = 1 To Data.Columns.Count
                    A(0, I) = Data.Columns(I - 1).ColumnName
                Next
                I = 1
                For Each R In Data.Rows
                    For X = 1 To Data.Columns.Count
                        A(I, X) = R(X - 1).ToString
                    Next
                    I += 1
                Next
            End If
            DataArray = A
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = RF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public Sub AddCustomField(ByVal FName As String, Optional ByVal FLabel As String = Nothing)

        If CF Is Nothing Then
            ReDim CF(1, 0)
        Else
            ReDim Preserve CF(1, CF.GetUpperBound(1) + 1)
        End If

        CF(0, CF.GetUpperBound(1)) = FName
        If Not FLabel Is Nothing Then
            CF(1, CF.GetUpperBound(1)) = FLabel
        Else
            CF(1, CF.GetUpperBound(1)) = FName
        End If

    End Sub

    Public Sub IncludeCustomParam(ByVal SAP_Name As String, ByVal ParamValue As String)

        T.ParamInclude(SAP_Name, ParamValue)

    End Sub

    Public Sub IncludeCustomParam_Date(ByVal SAP_Name As String, ByVal DateValue As Date)

        T.ParamInclude(SAP_Name, ConversionUtils.NetDate2SAPDate(DateValue))

    End Sub

    Public Sub ExcludeCustomParam(ByVal SAP_Name As String, ByVal ParamValue As String)

        T.ParamExclude(SAP_Name, ParamValue)

    End Sub

    Public Sub ExcludeCustomParam_Date(ByVal SAP_Name As String, ByVal DateValue As Date)

        T.ParamExclude(SAP_Name, ConversionUtils.NetDate2SAPDate(DateValue))

    End Sub

    Public Overridable Sub Execute()

        If Not RF Then Exit Sub

        Dim I As Integer = 0
        If Not CF Is Nothing Then
            For I = 0 To CF.GetUpperBound(1)
                T.AddField(CF(0, I), CF(1, I))
            Next
        End If
        T.Run()
        If Not T.Success Then
            EM = T.ErrMessage
            SF = False
        Else
            SF = True
        End If
        RF = False

    End Sub

    Public Sub ColumnToDateStr(ByVal ColumnName As String)
        T.ColumnToDateStr(ColumnName)
    End Sub

    Public Sub ColumnToDoubleStr(ByVal ColumnName As String)
        T.ColumnToDoubleStr(ColumnName)
    End Sub

    Public Sub ColumnToIntStr(ByVal ColumnName As String)
        T.ColumnToIntStr(ColumnName)
    End Sub

End Class

Public NotInheritable Class EKKO_Report : Inherits RTable_Report

    Private DI = Nothing                    'Deletion Indicator

    '*** From/To
    Private DNF As String = Nothing         'Document Number
    Private DNT As String = Nothing
    Private COF As String = Nothing         'Created On
    Private COT As String = Nothing
    Private DDF As String = Nothing         'Doc Date
    Private DDT As String = Nothing
    Private OAF As String = Nothing         'Outline Agreement
    Private OAT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DocumentFrom() As String

        Get
            DocumentFrom = DNF
        End Get

        Set(ByVal value As String)
            DNF = value
        End Set

    End Property

    Public Property DocumentTo() As String

        Get
            DocumentTo = DNT
        End Get

        Set(ByVal value As String)
            DNT = value
        End Set

    End Property

    Public Property CreatedFrom() As Date

        Get
            CreatedFrom = ConversionUtils.SAPDate2NetDate(COF)
        End Get

        Set(ByVal value As Date)
            COF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property CreatedTo() As Date

        Get
            CreatedTo = ConversionUtils.SAPDate2NetDate(COT)
        End Get

        Set(ByVal value As Date)
            COT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property DocDateFrom() As Date

        Get
            DocDateFrom = ConversionUtils.SAPDate2NetDate(DDF)
        End Get

        Set(ByVal value As Date)
            DDF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property DocDateTo() As Date

        Get
            DocDateTo = ConversionUtils.SAPDate2NetDate(DDT)
        End Get

        Set(ByVal value As Date)
            DDT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property OAFrom() As String

        Get
            OAFrom = OAF
        End Get

        Set(ByVal value As String)
            OAF = value
        End Set

    End Property

    Public Property OATo() As String

        Get
            OATo = OAT
        End Get

        Set(ByVal value As String)
            OAT = value
        End Set

    End Property

    Public Property DeletionIndicator() As Boolean

        Get
            If DI Is Nothing Then
                DeletionIndicator = False
            Else
                DeletionIndicator = DI
            End If
        End Get

        Set(ByVal value As Boolean)
            DI = value
        End Set

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub IncludeDocType(ByVal DType As String)

        If Not T Is Nothing Then
            T.ParamInclude("BSART", DType)
        End If

    End Sub

    Public Sub ExcludeDocType(ByVal DType As String)

        If Not T Is Nothing Then
            T.ParamExclude("BSART", DType)
        End If

    End Sub

    Public Sub IncludeDocCategory(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BSTYP", Code)
        End If

    End Sub

    Public Sub ExcludeDocCategory(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BSTYP", Code)
        End If

    End Sub

    Public Sub IncludePurchOrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKORG", Code)
        End If

    End Sub

    Public Sub ExcludePurchOrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKORG", Code)
        End If

    End Sub

    Public Sub IncludePurchGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKGRP", Code)
        End If

    End Sub

    Public Sub ExcludePurchGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKGRP", Code)
        End If

    End Sub

    Public Sub IncludeDocsDated(ByVal DDate As String)

        If Not T Is Nothing Then
            T.ParamInclude("BEDAT", ConversionUtils.NetDate2SAPDate(DDate))
        End If

    End Sub

    Public Sub ExcludeDocsDated(ByVal DDate As String)

        If Not T Is Nothing Then
            T.ParamExclude("BEDAT", ConversionUtils.NetDate2SAPDate(DDate))
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Doc Number] [Company Code] [Doc Type] [Created On] [Created By] [Vendor] [Language] [POrg] [PGrp] [Currency] [Doc Date]
    '''    [Validity Start] [Validity End] [Y Refer] [Salesperson] [Telephone] [OA] [O Reference]
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EKKO"

        If Not DNF Is Nothing And Not DNT Is Nothing Then
            T.ParamIncludeFromTo("EBELN", DNF, DNT)
        End If
        If Not COF Is Nothing And Not COT Is Nothing Then
            T.ParamIncludeFromTo("AEDAT", COF, COT)
        End If
        If Not DDF Is Nothing And Not DDT Is Nothing Then
            T.ParamIncludeFromTo("BEDAT", DDF, DDT)
        End If
        If Not OAF Is Nothing And Not OAT Is Nothing Then
            T.ParamIncludeFromTo("KONNR", OAF, OAT)
        End If

        If Not DI Is Nothing Then
            If DI Then
                T.ParamExclude("LOEKZ", " ")
            Else
                T.ParamInclude("LOEKZ", " ")
            End If
        End If

        T.AddField("MANDT", "Client")
        T.AddField("EBELN", "Doc Number")
        T.AddField("BUKRS", "Company Code")
        T.AddField("BSART", "Doc Type")
        T.AddField("AEDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("LIFNR", "Vendor")
        T.AddField("SPRAS", "Language")
        T.AddField("EKORG", "POrg")
        T.AddField("EKGRP", "PGrp")
        T.AddField("WAERS", "Currency")
        T.AddField("BEDAT", "Doc Date")
        T.AddField("KDATB", "Validity Start")
        T.AddField("KDATE", "Validity End")
        T.AddField("IHREZ", "Y Refer")
        T.AddField("VERKF", "Salesperson")
        T.AddField("TELF1", "Telephone")
        T.AddField("KONNR", "OA")
        T.AddField("UNSEZ", "O Reference")
        T.AddKeyColumn(0)
        T.AddKeyColumn(1)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Created On")
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDateStr("Doc Date")
            T.ColumnToDateStr("Validity Start")
            T.ColumnToDateStr("Validity End")
        End If

    End Sub

End Class

Public NotInheritable Class EKPO_Report : Inherits RTable_Report

    Private DT As DataTable = Nothing   'Document Numbers DataTable

    Private DI = Nothing                'Deletion Indicator

    '*** From/To
    Private DNF As String = Nothing     'Document Number
    Private DNT As String = Nothing
    Private MNF As String = Nothing     'Material
    Private MNT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DeletionIndicator() As Boolean

        Get
            If DI Is Nothing Then
                DeletionIndicator = False
            Else
                DeletionIndicator = DI
            End If
        End Get

        Set(ByVal value As Boolean)
            DI = value
        End Set

    End Property

    Public ReadOnly Property Documents() As DataTable

        Get
            Documents = DT
        End Get

    End Property

    Public Property DocumentFrom() As String

        Get
            DocumentFrom = DNF
        End Get

        Set(ByVal value As String)
            DNF = value
        End Set

    End Property

    Public Property DocumentTo() As String

        Get
            DocumentTo = DNT
        End Get

        Set(ByVal value As String)
            DNT = value
        End Set

    End Property

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF
        End Get

        Set(ByVal value As String)
            MNF = value
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT
        End Get

        Set(ByVal value As String)
            MNT = value
        End Set

    End Property

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Plant)
        End If

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    Public Sub IncludeAccAssignment(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("KNTTP", Code)
        End If

    End Sub

    Public Sub ExcludeAccAssignment(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("KNTTP", Code)
        End If

    End Sub

    Public Sub IncludeMatGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATKL", Code)
        End If

    End Sub

    Public Sub ExcludeMatGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATKL", Code)
        End If

    End Sub

    Public Sub IncludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub IncludeCompCode(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BUKRS", Code.PadLeft(3, "0"))
        End If

    End Sub

    Public Sub ExcludeCompCode(ByVal Code As String) '

        If Not T Is Nothing Then
            T.ParamExclude("BUKRS", Code.PadLeft(3, "0"))
        End If

    End Sub

    Public Sub IncludeDocCategory(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BSTYP", Code)
        End If

    End Sub

    Public Sub ExcludeDocCategory(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BSTYP", Code)
        End If

    End Sub

    Public Sub IncludeAgreement(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("KONNR", Code)
        End If

    End Sub

    Public Sub ExcludeAgreement(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("KONNR", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Doc Number] [Item Number] [Short Text] [Material] [Plant] [Inforecord] [Quantity] [UOM] [Price] [Tax code] [PDT]
    '''    [Mat Group] [Tracking Fld] [Price Unit]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EKPO"

        If Not DNF Is Nothing And Not DNT Is Nothing Then
            T.ParamIncludeFromTo("EBELN", DNF, DNT)
        End If
        If Not MNF Is Nothing And Not MNT Is Nothing Then
            T.ParamIncludeFromTo("MATNR", MNF.PadLeft(18, "0"), MNT.PadLeft(18, "0"))
        End If

        If Not DI Is Nothing Then
            If DI Then
                T.ParamExclude("LOEKZ", " ")
            Else
                T.ParamInclude("LOEKZ", " ")
            End If
        End If

        T.AddField("MANDT", "Client")
        T.AddField("EBELN", "Doc Number")
        T.AddField("EBELP", "Item Number")
        T.AddField("TXZ01", "Short Text")
        T.AddField("MATNR", "Material")
        T.AddField("WERKS", "Plant")
        T.AddField("INFNR", "Inforecord")
        T.AddField("MENGE", "Quantity")
        T.AddField("MEINS", "UOM")
        T.AddField("NETPR", "Price")
        T.AddField("MWSKZ", "Tax code")
        T.AddField("PLIFZ", "PDT")
        T.AddField("MATKL", "Mat Group")
        T.AddField("BEDNR", "Tracking Fld")
        T.AddField("PEINH", "Price Unit")
        T.AddKeyColumn(0)
        T.AddKeyColumn(1)
        T.AddKeyColumn(2)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            DT = New DataTable("Doc Numbers")
            DT.Columns.Add("Doc Number", System.Type.GetType("System.String"))
            DT.PrimaryKey = New DataColumn() {DT.Columns("Doc Number")}
            For Each R In T.Result.Rows
                DT.LoadDataRow(New Object() {R("Doc Number")}, LoadOption.OverwriteChanges)
            Next
            T.ColumnToDoubleStr("Material")
            T.ColumnToDoubleStr("Inforecord")
            T.ColumnToIntStr("Item Number")
        End If

    End Sub

End Class

Public NotInheritable Class EBAN_Report : Inherits RTable_Report

    Private DI = Nothing        'Deletion Indicator
    Private CI = Nothing        'Close Indicator

    Private DNF As String = Nothing         'Document Number
    Private DNT As String = Nothing
    Private COF As String = Nothing         'Created On
    Private COT As String = Nothing
    Private MNF As String = Nothing         'Material
    Private MNT As String = Nothing
    Private POF As String = Nothing         'PO Date
    Private POT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DocumentFrom() As String

        Get
            DocumentFrom = DNF
        End Get

        Set(ByVal value As String)
            DNF = value
        End Set

    End Property

    Public Property DocumentTo() As String

        Get
            DocumentTo = DNT
        End Get

        Set(ByVal value As String)
            DNT = value
        End Set

    End Property

    Public Property CreatedFrom() As Date

        Get
            CreatedFrom = ConversionUtils.SAPDate2NetDate(COF)
        End Get

        Set(ByVal value As Date)
            COF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property CreatedTo() As Date

        Get
            CreatedTo = ConversionUtils.SAPDate2NetDate(COT)
        End Get

        Set(ByVal value As Date)
            COT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property POCreatedFrom() As Date

        Get
            POCreatedFrom = ConversionUtils.SAPDate2NetDate(POF)
        End Get

        Set(ByVal value As Date)
            POF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property POCreatedTo() As Date

        Get
            POCreatedTo = ConversionUtils.SAPDate2NetDate(POT)
        End Get

        Set(ByVal value As Date)
            POT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF
        End Get

        Set(ByVal value As String)
            MNF = value
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT
        End Get

        Set(ByVal value As String)
            MNT = value
        End Set

    End Property

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Plant)
        End If

    End Sub

    Public Sub IncludePOrg(ByVal POrg As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKORG", POrg)
        End If

    End Sub

    Public Sub ExcludePOrg(ByVal POrg As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKORG", POrg)
        End If

    End Sub

    Public Sub IncludePGrp(ByVal PGrp As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKGRP", PGrp)
        End If

    End Sub

    Public Sub ExcludePGrp(ByVal PGrp As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKGRP", PGrp)
        End If

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BANFN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BANFN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub IncludePurchOrder(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludePurchOrder(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    Public Property DeletionIndicator() As Boolean

        Get
            If DI Is Nothing Then
                DeletionIndicator = False
            Else
                DeletionIndicator = DI
            End If
        End Get

        Set(ByVal value As Boolean)
            DI = value
        End Set

    End Property

    Public Property CloseIndicator() As Boolean

        Get
            If CI Is Nothing Then
                CloseIndicator = False
            Else
                CloseIndicator = CI
            End If
        End Get

        Set(ByVal value As Boolean)
            CI = value
        End Set

    End Property

    ''' <summary>
    ''' Returns:
    '''    [Client] [Req Number] [Req Item] [Req Date] [Purch Org] [Purch Doc] [PO Item]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EBAN"

        T.AddField("MANDT", "Client")
        T.AddField("BANFN", "Req Number")
        T.AddField("BNFPO", "Req Item")
        T.AddField("BADAT", "Req Date")
        T.AddField("EKORG", "Purch Org")
        T.AddField("EBELN", "Purch Doc")
        T.AddField("EBELP", "PO Item")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)
        T.AddKeyColumn(2)

        If Not DNF Is Nothing And Not DNT Is Nothing Then
            T.ParamIncludeFromTo("BANFN", DNF, DNT)
        End If

        If Not COF Is Nothing And Not COT Is Nothing Then
            T.ParamIncludeFromTo("BADAT", COF, COT)
        End If

        If Not MNF Is Nothing And Not MNT Is Nothing Then
            T.ParamIncludeFromTo("MATNR", MNF.PadLeft(18, "0"), MNT.PadLeft(18, "0"))
        End If

        If Not POF Is Nothing And Not POT Is Nothing Then
            T.ParamIncludeFromTo("BEDAT", POF, POT)
        End If

        If Not DI Is Nothing Then
            If DI Then
                T.ParamExclude("LOEKZ", " ")
            Else
                T.ParamInclude("LOEKZ", " ")
            End If
        End If

        If Not CI Is Nothing Then
            If CI Then
                T.ParamExclude("EBAKZ", " ")
            Else
                T.ParamInclude("EBAKZ", " ")
            End If
            T.AddField("EBAKZ", "Closed")
        End If

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Req Date")
            T.ColumnToDoubleStr("Req Number")
            T.ColumnToIntStr("Req Item")
        End If

    End Sub

End Class

Public NotInheritable Class EKET_Report : Inherits RTable_Report

    Private DNF As String = Nothing         'Document Number
    Private DNT As String = Nothing
    Private LQOD = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public ReadOnly Property OpenDeliveries_Data() As DataTable

        Get
            OpenDeliveries_Data = LinQToDataTable(LQOD)
        End Get

    End Property

    Public Property DocumentFrom() As String

        Get
            DocumentFrom = DNF
        End Get

        Set(ByVal value As String)
            DNF = value
        End Set

    End Property

    Public Property DocumentTo() As String

        Get
            DocumentTo = DNT
        End Get

        Set(ByVal value As String)
            DNT = value
        End Set

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    Public Sub Include_Item(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELP", Number.PadLeft(5, "0"))
        End If

    End Sub

    Public Sub Exclude_Item(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELP", Number.PadLeft(5, "0"))
        End If

    End Sub

    Public Sub Include_Req(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BANFN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Exclude_Req(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BANFN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Include_ReqItem(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BNFPO", Number.PadLeft(5, "0"))
        End If

    End Sub

    Public Sub Exclude_ReqItem(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BNFPO", Number.PadLeft(5, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Purch Doc] [Purch Doc Item] [Schedule Line] [Delivery Date] [Stat Del Date] [Scheduled Qty] [Delivered Qty] [Purch Req] [Purch Req Item] [Purch Doc Date] 
    '''    [Quota Arrangement] [Qta Arr Item] [Creation Indicator]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        If Not DNF Is Nothing And Not DNT Is Nothing Then
            T.ParamIncludeFromTo("EBELN", DNF, DNT)
        End If

        T.TableName = "EKET"

        T.AddField("EBELN", "Purch Doc")
        T.AddField("EBELP", "Purch Doc Item")
        T.AddField("ETENR", "Schedule Line")
        T.AddField("EINDT", "Delivery Date")
        T.AddField("SLFDT", "Stat Del Date")
        T.AddField("MENGE", "Scheduled Qty")
        T.AddField("WEMNG", "Delivered Qty")
        T.AddField("BANFN", "Purch Req")
        T.AddField("BNFPO", "Purch Req Item")
        T.AddField("BEDAT", "Purch Doc Date")
        T.AddField("QUNUM", "Quota Arrangement")
        T.AddField("QUPOS", "Qta Arr Item")
        T.AddField("ESTKZ", "Creation Indicator")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)
        T.AddKeyColumn(2)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Delivery Date")
            T.ColumnToDateStr("Stat Del Date")
            T.ColumnToIntStr("Purch Doc Item")
            T.ColumnToDoubleStr("Purch Req")
            T.ColumnToIntStr("Purch Req Item")
            T.ColumnToDateStr("Purch Doc Date")
            T.ColumnToIntStr("Qta Arr Item")
            T.ColumnToDoubleStr("Scheduled Qty")
            T.ColumnToDoubleStr("Delivered Qty")
            LQOD = From EKET In T.Result Where CDate(EKET("Delivery Date")) >= My.Computer.Clock.LocalTime.AddYears(-1) _
                Group EKET By PurchDoc = EKET("Purch Doc"), DocItem = EKET("Purch Doc Item") Into G = Group _
                Select New With { _
                .Document = PurchDoc, _
                .Item = DocItem, _
                .Scheduled = G.Sum(Function(EKET) CDbl(EKET("Scheduled Qty").ToString.Replace("*", ""))), _
                .Delivered = G.Sum(Function(EKET) CDbl(EKET("Delivered Qty").ToString.Replace("*", ""))), _
                .Open = If(.Scheduled > .Delivered, "Yes", "No") _
            }
        End If

    End Sub

End Class

Public NotInheritable Class EKBE_Report : Inherits RTable_Report

    '*** From/To
    Private PDF As String = Nothing         'Posting Date
    Private PDT As String = Nothing
    Private MNF As String = Nothing     'Material
    Private MNT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property PostingDateFrom() As Date

        Get
            PostingDateFrom = ConversionUtils.SAPDate2NetDate(PDF)
        End Get

        Set(ByVal value As Date)
            PDF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property PostingDateTo() As Date

        Get
            PostingDateTo = ConversionUtils.SAPDate2NetDate(PDT)
        End Get

        Set(ByVal value As Date)
            PDT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF
        End Get

        Set(ByVal value As String)
            MNF = value
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT
        End Get

        Set(ByVal value As String)
            MNT = value
        End Set

    End Property

    Public Sub Include_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Exclude_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Include_Item(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELP", Number.PadLeft(5, "0"))
        End If

    End Sub

    Public Sub Exclude_Item(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELP", Number.PadLeft(5, "0"))
        End If

    End Sub

    Public Sub Include_Movement(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BWART", Code)
        End If

    End Sub

    Public Sub Exclude_Movement(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BWART", Code)
        End If

    End Sub

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Plant)
        End If

    End Sub

    Public Sub Include_Reference(ByVal Value As String)

        If Not T Is Nothing Then
            T.ParamInclude("XBLNR", Value)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Purch Doc] [Item] [Transaction] [Movement] [Posting Date] [Quantity] [Amount] [Currency] [Reference] [Material] [Plant] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EKBE"

        If Not PDF Is Nothing And Not PDT Is Nothing Then
            T.ParamIncludeFromTo("BUDAT", PDF, PDT)
        End If

        If Not MNF Is Nothing And Not MNT Is Nothing Then
            T.ParamIncludeFromTo("MATNR", MNF.PadLeft(18, "0"), MNT.PadLeft(18, "0"))
        End If

        T.AddField("EBELN", "Purch Doc")
        T.AddField("EBELP", "Item")
        T.AddField("VGABE", "Transaction")
        T.AddField("BWART", "Movement")
        T.AddField("BUDAT", "Posting Date")
        T.AddField("MENGE", "Quantity")
        T.AddField("WRBTR", "Amount")
        T.AddField("WAERS", "Currency")
        T.AddField("XBLNR", "Reference")
        T.AddField("MATNR", "Material")
        T.AddField("WERKS", "Plant")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToIntStr("Item")
            T.ColumnToDateStr("Posting Date")
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class EORD_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeOA(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludeOA(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    Public Sub IncludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Material] [Plant] [Number] [Created On] [Created By] [Valid From] [Valid To] [Vendor] [Fixed Vendor] [Agreement] [Agreement Item] [Fixed Agreement Item]
    ''' [MPN Material] [Blocked] [Purch Org] [Doc Category] [Control Ind] [Materials Planning] [UOM] [Logical System] [Special Stock] [Central Contract] [Central Contract Item] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EORD"

        T.AddField("MANDT", "Client")
        T.AddField("MATNR", "Material")
        T.AddField("WERKS", "Plant")
        T.AddField("ZEORD", "Number")
        T.AddField("ERDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("VDATU", "Valid From")
        T.AddField("BDATU", "Valid To")
        T.AddField("LIFNR", "Vendor")
        T.AddField("FLIFN", "Fixed Vendor")
        T.AddField("EBELN", "Agreement")
        T.AddField("EBELP", "Agreement Item")
        T.AddField("FEBEL", "Fixed Agreement Item")
        T.AddField("RESWK", "Procurement Plant")
        T.AddField("FRESW", "Fixed Issuing Plant")
        T.AddField("EMATN", "MPN Material")
        T.AddField("NOTKZ", "Blocked")
        T.AddField("EKORG", "Purch Org")
        T.AddField("VRTYP", "Doc Category")
        T.AddField("EORTP", "Control Ind")
        T.AddField("AUTET", "Materials Planning")
        T.AddField("MEINS", "UOM")
        T.AddField("LOGSY", "Logical System")
        T.AddField("SOBKZ", "Special Stock")
        T.AddField("SRM_CONTRACT_ID", "Central Contract")
        T.AddField("SRM_CONTRACT_ITM", "Central Contract Item")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)
        T.AddKeyColumn(10)
        T.AddKeyColumn(11)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Created On")
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDoubleStr("Fixed Vendor")
            T.ColumnToDoubleStr("Agreement Item")
            T.ColumnToDoubleStr("Fixed Agreement Item")
            T.ColumnToDateStr("Valid From")
            T.ColumnToDateStr("Valid To")
            T.ColumnToDoubleStr("Material")
            T.ColumnToDoubleStr("Number")
        End If


    End Sub

End Class

Public NotInheritable Class EBKN_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeReq(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BANFN", Number)
        End If

    End Sub

    Public Sub ExcludeReq(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BANFN", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Req Number] [Req Item] [GL Account]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EBKN"

        T.AddField("MANDT", "Client")
        T.AddField("BANFN", "Req Number")
        T.AddField("BNFPO", "Req Item")
        T.AddField("SAKTO", "GL Account")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)
        T.AddKeyColumn(2)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Req Number")
            T.ColumnToIntStr("Req Item")
        End If

    End Sub

End Class

Public NotInheritable Class EKES_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Doc Number] [Item] [Conf Category] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EKES"

        T.AddField("MANDT", "Client")
        T.AddField("EBELN", "Doc Number")
        T.AddField("EBELP", "Item")
        T.AddField("EBTYP", "Conf Category")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToIntStr("Item")
        End If

    End Sub

End Class

Public NotInheritable Class EIPO_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeFTDNum(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EXNUM", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeFTDNum(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EXNUM", Number.PadLeft(10, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [FTD Number] [Disp Country]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EIPO"

        T.AddField("MANDT", "Client")
        T.AddField("EXNUM", "FTD Number")
        T.AddField("VERLD", "Disp Country")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class EINA_Report : Inherits RTable_Report

    Private DI = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DeletionIndicator() As Boolean

        Get
            If DI Is Nothing Then
                DeletionIndicator = False
            Else
                DeletionIndicator = DI
            End If
        End Get

        Set(ByVal value As Boolean)
            DI = value
        End Set

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("INFNR", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("INFNR", Number)
        End If

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Inforecord] [Material] [Material Group] [Vendor] [Created On] [Created By] [Short Text] [Sort Term] [Order Unit]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EINA"

        If Not DI Is Nothing Then
            If DI Then
                T.ParamExclude("LOEKZ", " ")
            Else
                T.ParamInclude("LOEKZ", " ")
            End If
        End If

        T.AddField("MANDT", "Client")
        T.AddField("INFNR", "Inforecord")
        T.AddField("MATNR", "Material")
        T.AddField("MATKL", "Material Group")
        T.AddField("LIFNR", "Vendor")
        T.AddField("ERDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("TXZ01", "Short Text")
        T.AddField("SORTL", "Sort Term")
        T.AddField("MEINS", "Order Unit")

        MyBase.Execute()

        T.ColumnToDoubleStr("Material")
        T.ColumnToDoubleStr("Vendor")
        T.ColumnToDateStr("Created On")

    End Sub

End Class

Public NotInheritable Class EIPA_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("INFNR", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("INFNR", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Inforecord] [Document] [Item]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EIPA"

        T.AddField("MANDT", "Client")
        T.AddField("INFNR", "Inforecord")
        T.AddField("EKORG", "Purch Org")
        T.AddField("EBELN", "Document")
        T.AddField("EBELP", "Item")

        MyBase.Execute()

        T.ColumnToIntStr("Item")

    End Sub

End Class

Public NotInheritable Class EINE_Report : Inherits RTable_Report

    Private DI = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DeletionIndicator() As Boolean

        Get
            If DI Is Nothing Then
                DeletionIndicator = False
            Else
                DeletionIndicator = DI
            End If
        End Get

        Set(ByVal value As Boolean)
            DI = value
        End Set

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("INFNR", Number)
        End If

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("INFNR", Number)
        End If

    End Sub

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Plant)
        End If

    End Sub

    Public Sub IncludePurchGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKGRP", Code)
        End If

    End Sub

    Public Sub ExcludePurchGroup(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKGRP", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Client] [Inforecord] [Purch Org] [Plant] [Created On] [Created By] [Purch Group] [Currency] [Purch Doc] [Item] [Doc Date] [Net Price]
    '''  [Price Unit] [OPU] [Conf Control] [GR Based IV] [Standard Qty] [Unlimited Ovrdl] [Pricing Date Ctrl] [Qty Conversion 1] [Qty Conversion 2]
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EINE"

        If Not DI Is Nothing Then
            If DI Then
                T.ParamExclude("LOEKZ", " ")
            Else
                T.ParamInclude("LOEKZ", " ")
            End If
        End If

        T.AddField("MANDT", "Client")
        T.AddField("INFNR", "Inforecord")
        T.AddField("EKORG", "Purch Org")
        T.AddField("WERKS", "Plant")
        T.AddField("ERDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("EKGRP", "Purch Group")
        T.AddField("WAERS", "Currency")
        T.AddField("EBELN", "Purch Doc")
        T.AddField("EBELP", "Item")
        T.AddField("DATLB", "Doc Date")
        T.AddField("NETPR", "Net Price")
        T.AddField("PEINH", "Price Unit")
        T.AddField("BPRME", "OPU")
        T.AddField("BSTAE", "Conf Control")
        T.AddField("WEBRE", "GR Based IV")
        T.AddField("NORBM", "Standard Qty")
        T.AddField("UEBTK", "Unlimited Ovrdl")
        T.AddField("MEPRF", "Pricing Date Ctrl")
        T.AddField("BPUMZ", "Qty Conversion 1")
        T.AddField("BPUMN", "Qty Conversion 2")
        T.AddField("MWSKZ", "Tax Code")
        T.AddField("INCO1", "Incoterm")
        T.AddField("INCO2", "Incoterm Desc")
        T.AddField("XERSN", "ERS")
        T.AddField("APLFZ", "PDT")

        MyBase.Execute()

        T.ColumnToDateStr("Created On")
        T.ColumnToDateStr("Doc Date")
        T.ColumnToIntStr("Item")

    End Sub

End Class

Public NotInheritable Class EKBZ_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_DocNumber(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BELNR", Number)
        End If

    End Sub

    Public Sub Exclude_DocNumber(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BELNR", Number)
        End If

    End Sub

    Public Sub Include_PurchasingDoc(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Sub Exclude_PurchasingDoc(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("EBELN", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Purchasing Document] [Item] [Step number] ...
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "EKBZ"

        T.AddField("MANDT", "Client")
        T.AddField("EBELN", "Purchasing Document")
        T.AddField("EBELP", "Item")
        T.AddField("STUNR", "Step number")
        T.AddField("ZAEHK", "Counter")
        T.AddField("VGABE", "Trans./event type")
        T.AddField("GJAHR", "Fiscal Year")
        T.AddField("BELNR", "Document Number")
        T.AddField("BUZEI", "Material Doc.Item")
        T.AddField("BEWTP", "PO History Category")
        T.AddField("BUDAT", "Posting Date")
        T.AddField("MENGE", "Quantity")
        T.AddField("DMBTR", "Amount in LC")
        T.AddField("WRBTR", "Amount")
        T.AddField("WAERS", "Currency")
        T.AddField("AREWR", "GR/IR clearing value in local currency")
        T.AddField("SHKZG", "Debit/Credit Ind.")
        T.AddField("XBLNR", "Reference")
        T.AddField("FRBNR", "Bill of Lading")
        T.AddField("LIFNR", "Vendor")
        T.AddField("CPUDT", "Entry Date")
        T.AddField("CPUTM", "Time of Entry")
        T.AddField("REEWR", "Invoice Value")
        T.AddField("REFWR", "Invoice value in FC")
        T.AddField("BWTAR", "Valuation Type")
        T.AddField("KSCHL", "Condition Type")
        T.AddField("BPMNG", "Qty in OPUn")
        T.AddField("AREWW", "GR/IR clearing value in FC")
        T.AddField("HSWAE", "Local currency")
        T.AddField("VNETW", "Net value")
        T.AddField("ERNAM", "Created by")
        T.AddField("SHKKO", "Debit/Credit Ind._0")
        T.AddField("AREWB", "GR/IR clearing value in FC_0")
        T.AddField("REWRB", "FC invoice amount")
        T.AddField("SAPRL", "SAP Release")
        T.AddField("MENGE_POP", "Quantity_0")
        T.AddField("DMBTR_POP", "Amount in LC_0")
        T.AddField("WRBTR_POP", "Amount_0")
        T.AddField("BPMNG_POP", "Qty in OPUn_0")
        T.AddField("AREWR_POP", "GR/IR clearing value in local currency_0")
        T.AddField("KUDIF", "Exch. Rate Diff. Amt")
        T.AddField("XMACC", "Multiple Acct Assignment")
        T.AddField("WKURS", "Exchange Rate")

        MyBase.Execute()

        If Not T.Result Is Nothing AndAlso T.Result.Rows.Count > 0 Then

            T.ColumnToDateStr("Posting Date")
            T.ColumnToDateStr("Entry Date")
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDoubleStr("Quantity")
            T.ColumnToDoubleStr("Amount in LC")
            T.ColumnToDoubleStr("Amount")
            T.ColumnToDoubleStr("GR/IR clearing value in local currency")
            T.ColumnToDoubleStr("Invoice Value")
            T.ColumnToDoubleStr("Invoice value in FC")
            T.ColumnToDoubleStr("Qty in OPUn")
            T.ColumnToDoubleStr("GR/IR clearing value in FC")
            T.ColumnToDoubleStr("Net value")
            T.ColumnToDoubleStr("GR/IR clearing value in FC_0")
            T.ColumnToDoubleStr("FC invoice amount")
            T.ColumnToDoubleStr("Quantity_0")
            T.ColumnToDoubleStr("Amount in LC_0")
            T.ColumnToDoubleStr("Amount_0")
            T.ColumnToDoubleStr("Qty in OPUn_0")
            T.ColumnToDoubleStr("GR/IR clearing value in local currency_0")
            T.ColumnToDoubleStr("Exch. Rate Diff. Amt")
            T.ColumnToDoubleStr("Exchange Rate")

        End If

    End Sub

End Class

Public NotInheritable Class BKPF_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Key(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("AWKEY", Number)
        End If

    End Sub

    Public Sub Exclude_Key(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("AWKEY", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Doc_Number] [Reference_Key] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "BKPF"

        T.AddField("BELNR", "Doc_Number")
        T.AddField("AWKEY", "Reference_Key")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class BSEG_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BELNR", Number)
        End If

    End Sub

    Public Sub Exclude_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("BELNR", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Doc Number] [Line Item] [BaseLine Date] [Payment Terms] [Days1]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "BSEG"

        T.AddField("BELNR", "Doc Number")
        T.AddField("BUZEI", "Line Item")
        T.AddField("ZFBDT", "BaseLine Date")
        T.AddField("ZTERM", "Payment Terms")
        T.AddField("ZBD1T", "Days1")

        MyBase.Execute()

        If Not T.Result Is Nothing AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("BaseLine Date")
        End If

    End Sub

End Class

Public NotInheritable Class MARC_Report : Inherits RTable_Report

    Private MNF As String = Nothing     'Material From - To
    Private MNT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF.Replace("0", "").Trim
        End Get

        Set(ByVal value As String)
            MNF = value.PadLeft(18, "0")
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT.Replace("0", "").Trim
        End Get

        Set(ByVal value As String)
            MNT = value.PadLeft(18, "0")
        End Set

    End Property

    Public Sub IncludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Code)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Code)
        End If

    End Sub

    Public Sub IncludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Material.Trim.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Material] [Plant] [Purch Group] [PDT] [AutoPO] [Source List] [MRP Group] [Deletion Flag PL] [MRP Controller]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "MARC"

        T.AddField("MATNR", "Material")
        T.AddField("WERKS", "Plant")
        T.AddField("EKGRP", "Purch Group")
        T.AddField("PLIFZ", "PDT")
        T.AddField("KAUTB", "AutoPO")
        T.AddField("KORDB", "Source List")
        T.AddField("DISGR", "MRP Group")
        T.AddField("LVORM", "Deletion Flag PL")
        T.AddField("DISPO", "MRP Controller")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class LFA1_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Vendor] [Name] [Country]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "LFA1"

        T.AddField("MANDT", "Client")
        T.AddField("LIFNR", "Vendor")
        T.AddField("NAME1", "Name")
        T.AddField("LAND1", "Country")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
        End If

    End Sub

End Class

Public NotInheritable Class LFM1_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Include_POrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKORG", Code)
        End If

    End Sub

    Public Sub Exclude_POrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKORG", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Vendor] [POrg] [Created On] [Created By] [Incoterms] [Incoterms Desc] [Block] [Delete]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "LFM1"

        T.AddField("LIFNR", "Vendor")
        T.AddField("EKORG", "POrg")
        T.AddField("ERDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("INCO1", "Incoterms")
        T.AddField("INCO2", "Incoterms Desc")
        T.AddField("SPERM", "Block")
        T.AddField("LOEVM", "Delete")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDateStr("Created On")
        End If

    End Sub

End Class

Public NotInheritable Class LFB1_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Include_CCode(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BUKRS", Code)
        End If

    End Sub

    Public Sub Exclude_CCode(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BUKRS", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Vendor] [CCode] [Created On] [Created By] [Pmnt Terms] [Block] [Delete]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "LFB1"

        T.AddField("LIFNR", "Vendor")
        T.AddField("BUKRS", "CCode")
        T.AddField("ERDAT", "Created On")
        T.AddField("ERNAM", "Created By")
        T.AddField("ZTERM", "Pmnt Terms")
        T.AddField("SPERR", "Block")
        T.AddField("LOEVM", "Delete")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDateStr("Created On")
        End If

    End Sub

End Class

Public NotInheritable Class LFBK_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Vendor] [Bank Account]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "LFBK"

        T.AddField("LIFNR", "Vendor")
        T.AddField("BANKN", "Bank Account")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
        End If

    End Sub

End Class

Public NotInheritable Class T001K_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BWKEY", Code.PadLeft(4, "0"))
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BWKEY", Code.PadLeft(4, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Plant] [CCode] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "T001K"

        T.AddField("BWKEY", "Plant")
        T.AddField("BUKRS", "CCode")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class Manufacturers_Report : Inherits RTable_Report

    Private MN As RTable = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("PWERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("PWERKS", Plant)
        End If

    End Sub

    Public Sub IncludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Plant] [Material] [Manufacturer] [Part No]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZMANP"
        T.IncludeAllFields()
        MN = New RTable(Con)
        MN.TableName = "ZMANUFA10"
        MN.IncludeAllFields()

        T.AddKeyColumn(1)
        T.AddKeyColumn(2)
        T.AddKeyColumn(3)

        MyBase.Execute()
        If Not SF Then Exit Sub

        If T.Result.Rows.Count <> 0 Then
            MN.Run()
            T.Result.Columns(0).ColumnName = "Client"
            T.Result.Columns(1).ColumnName = "Plant"
            T.Result.Columns(2).ColumnName = "Material"
            T.Result.Columns(3).ColumnName = "Manufacturer"
            T.Result.Columns(4).ColumnName = "Part No"
            T.ColumnToDoubleStr("Material")
            Dim R As DataRow
            For Each R In T.Result.Rows
                R("Manufacturer") = MN.Result.Select("MANUF = '" & R("Manufacturer") & "'")(0)("MANUFN")
                R.AcceptChanges()
            Next
        End If

    End Sub

End Class

Public NotInheritable Class MARA_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Material] [Mat Type]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "MARA"

        T.AddField("MATNR", "Material")
        T.AddField("MTART", "Mat Type")

        MyBase.Execute()

        If T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class MAKT_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Material] [Description 1] [Description 2]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "MAKT"

        T.AddField("MATNR", "Material")
        T.AddField("MAKTX", "Description 1")
        T.AddField("MAKTG", "Description 2")

        MyBase.Execute()

        If T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class MARM_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Left("000000000000000000", 18 - Len(Number)) & Number)
        End If

    End Sub

    Public Sub IncludeAltUOM(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("MEINH", Code)
        End If

    End Sub

    Public Sub ExcludeAltUOM(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("MEINH", Code)
        End If

    End Sub

    Public Sub IncludeEAN_UPC(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EAN11", Code)
        End If

    End Sub

    Public Sub ExcludeEAN_UPC(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EAN11", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Material] [Alt UOM] [Numerator] [Denominator]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "MARM"

        T.AddField("MATNR", "Material")
        T.AddField("MEINH", "Alt UOM")
        T.AddField("UMREZ", "Numerator")
        T.AddField("UMREN", "Denominator")

        MyBase.Execute()

        If T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class A016_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Condition(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("KSCHL", Code)
        End If

    End Sub

    Public Sub Include_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EVRTN", Number & "*")
        End If

    End Sub

    Public Sub Include_Item(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EVRTP", Number.PadLeft(5, "0") & "*")
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Cond Type] [Document] [Item] [VFrom] [VTo] [CRNum]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "A016"

        T.AddField("KSCHL", "Cond Type")
        T.AddField("EVRTN", "Document")
        T.AddField("EVRTP", "Item")
        T.AddField("DATAB", "VFrom")
        T.AddField("DATBI", "VTo")
        T.AddField("KNUMH", "CRNum")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToIntStr("Item")
            T.ColumnToDateStr("VFrom")
            T.ColumnToDateStr("VTo")
        End If

    End Sub

End Class

Public NotInheritable Class A025_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Condition(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("KSCHL", Code)
        End If

    End Sub

    Public Sub Include_Inforecord(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("INFNR", Number & "*")
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Cond Type] [Vendor] [Purch Org] [Mat Group] [Inforecord] [Plant] [VTo] [VFrom] [CRNum]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "A025"

        T.AddField("MANDT", "Client")
        T.AddField("KSCHL", "Cond Type")
        T.AddField("LIFNR", "Vendor")
        T.AddField("EKORG", "Purch Org")
        T.AddField("MATKL", "Mat Group")
        T.AddField("INFNR", "Inforecord")
        T.AddField("WERKS", "Plant")
        T.AddField("DATBI", "VTo")
        T.AddField("DATAB", "VFrom")
        T.AddField("KNUMH", "CRNum")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
            T.ColumnToDateStr("VTo")
            T.ColumnToDateStr("VFrom")
        End If

    End Sub

End Class

Public NotInheritable Class CONDH_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Condition(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("KSCHL", Code)
        End If

    End Sub

    Public Sub Include_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("VAKEY", Number & "*")
        End If

    End Sub

    Public Sub IncludeKey(ByVal Document As String, ByVal Item As String)

        If Not T Is Nothing Then
            T.ParamInclude("VAKEY", Document & Item.PadLeft(5, "0"))
        End If

    End Sub

    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "KONH"

        T.AddField("KNUMH", "CRNum")
        T.AddField("VAKEY", "Key")
        T.AddField("DATAB", "VFrom")
        T.AddField("DATBI", "VTo")
        T.AddField("KSCHL", "Cond Type")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.Result.Columns.Add("Document", System.Type.GetType("System.String"), "SUBSTRING(Key,1,10)")
            T.Result.Columns.Add("Item", System.Type.GetType("System.String"), "CONVERT(CONVERT(SUBSTRING(Key,11,5),System.Int32),System.String)")
            T.ColumnToDateStr("VFrom")
            T.ColumnToDateStr("VTo")
        End If

    End Sub

End Class

Public NotInheritable Class CONDI_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeCRNum(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("KNUMH", Left("0000000000", 10 - Len(Number)) & Number)
        End If

    End Sub

    Public Sub Exclude_Condition(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("KSCHL", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [CRNum] [Condition] [Price] [Per] [UOM] [Currency]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "KONP"

        T.AddField("KNUMH", "CRNum")
        T.AddField("KSCHL", "Condition")
        T.AddField("KBETR", "Price")
        T.AddField("KPEIN", "Per")
        T.AddField("KMEIN", "UOM")
        T.AddField("KONWA", "Currency")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class OA_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("EBELN", Number)
        End If

    End Sub

    Public Overrides Sub Execute()

        If T Is Nothing Then Exit Sub

        T.TableName = "EKKO"

        T.AddField("EBELN", "Doc Number")
        T.AddField("KONNR", "OA")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class SAPExchgRate : Inherits RTable_Report

    Private ER As Double = 0
    Private FC As String = Nothing
    Private TC As String = Nothing
    Private VF As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public ReadOnly Property Rate() As Double

        Get
            Rate = ER
        End Get

    End Property

    Public Property FromCurrency() As String

        Get
            FromCurrency = FC
        End Get

        Set(ByVal value As String)
            FC = value
        End Set

    End Property

    Public Property ToCurrency() As String

        Get
            ToCurrency = TC
        End Get

        Set(ByVal value As String)
            TC = value
        End Set

    End Property

    Public Property ValidFrom() As Date

        Get
            ValidFrom = Inverted_Date(VF)
        End Get

        Set(ByVal value As Date)
            VF = Inverted_Date(value)
        End Set

    End Property

    Public Overrides Sub Execute()

        If Not RF Then Exit Sub
        If FC Is Nothing Or TC Is Nothing Then Exit Sub

        If VF Is Nothing Then
            VF = Inverted_Date(My.Computer.Clock.LocalTime.AddDays((My.Computer.Clock.LocalTime.Day * -1) + 1))
        End If

        T.TableName = "TCURR"

        T.ParamInclude("FCURR", FC)
        T.ParamInclude("TCURR", TC)
        T.ParamInclude("GDATU", VF)

        T.AddField("KURST")
        T.AddField("FCURR")
        T.AddField("TCURR")
        T.AddField("GDATU")
        T.AddField("UKURS")
        T.AddField("FFACT")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            Dim Rate As Double
            Dim Ratio As Double
            Dim FR() As DataRow
            If T.Result.Rows.Count = 1 Then
                Rate = CDbl(T.Result.Rows(0)("UKURS"))
                Ratio = CDbl(T.Result.Rows(0)("FFACT"))
            Else
                FR = T.Result.Select("KURST = 'EURX'")
                If FR.Length <> 0 Then
                    Rate = CDbl(FR(0)("UKURS"))
                    Ratio = CDbl(FR(0)("FFACT"))
                End If
                FR = T.Result.Select("KURST = 'LLR'")
                If FR.Length <> 0 Then
                    Rate = CDbl(FR(0)("UKURS"))
                    Ratio = CDbl(FR(0)("FFACT"))
                End If
                FR = T.Result.Select("KURST = 'M'")
                If FR.Length <> 0 Then
                    Rate = CDbl(FR(0)("UKURS"))
                    Ratio = CDbl(FR(0)("FFACT"))
                End If
            End If
            If Rate < 0 Then
                ER = (Ratio * Rate) * -1
            Else
                ER = Ratio / Rate
            End If
            ER = Math.Round(ER, 2)
        End If

    End Sub

    Private Function Inverted_Date(ByVal Value) As Object

        Dim R = Nothing
        Dim SF As String
        If IsDate(Value) Then
            Dim D As Date = Value
            SF = D.Year.ToString & Left("00", 2 - Len(D.Month.ToString)) & D.Month.ToString & Left("00", 2 - Len(D.Day.ToString)) & D.Day.ToString
            R = 99999999 - Val(SF)
        Else
            R = ConversionUtils.SAPDate2NetDate((99999999 - Val(Value)).ToString)
        End If
        Inverted_Date = R

    End Function

End Class

Public NotInheritable Class QInf_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Number.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Number.PadLeft(18, "0"))
        End If

    End Sub

    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "QINF"

        T.AddField("MATNR", "Material")
        T.AddField("LIEFERANT", "Vendor")
        T.AddField("WERK", "Plant")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class Plants_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Code)
        End If

    End Sub

    Public Overrides Sub Execute()

        T.TableName = "T001W"

        T.AddField("WERKS", "Plant")
        T.AddField("NAME1", "Name")
        T.AddField("LAND1", "Country")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class Buyers_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludePGrp(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKGRP", Code)
        End If

    End Sub

    Public Overrides Sub Execute()

        T.TableName = "T024"

        T.AddField("MANDT", "Client")
        T.AddField("EKGRP", "PGrp")
        T.AddField("EKNAM", "Description")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class PTerms_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludePTemr(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("ZTERM", Code)
        End If

    End Sub

    Public Overrides Sub Execute()

        T.TableName = "T052U"

        T.AddField("MANDT", "Client")
        T.AddField("ZTERM", "PTerm")
        T.AddField("TEXT1", "Description")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class NAST_Report : Inherits RTable_Report

    Private COF As String = Nothing         'Created On
    Private COT As String = Nothing
    Private DNF As String = Nothing         'Document Number
    Private DNT As String = Nothing
    Private SAR As Boolean = False          'Show all Records

    Public Property CreatedFrom() As Date
        Get
            Return ConversionUtils.SAPDate2NetDate(COF)
        End Get

        Set(ByVal value As Date)
            COF = ConversionUtils.NetDate2SAPDate(value)
        End Set
    End Property

    Public Property CreatedTo() As Date
        Get
            Return ConversionUtils.SAPDate2NetDate(COT)
        End Get

        Set(ByVal value As Date)
            COT = ConversionUtils.NetDate2SAPDate(value)
        End Set
    End Property

    Public Property DocumentFrom() As String

        Get
            Return DNF
        End Get

        Set(ByVal value As String)
            DNF = value
        End Set

    End Property

    Public Property DocumentTo() As String

        Get
            Return DNT
        End Get

        Set(ByVal value As String)
            DNT = value
        End Set

    End Property

    Public Property Show_All_Records() As Boolean
        Get
            Return SAR
        End Get
        Set(ByVal value As Boolean)
            SAR = value
        End Set
    End Property

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)
        MyBase.New(Box, User, App)
    End Sub

    Sub New(ByVal Connection As Object)
        MyBase.New(Connection)
    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("OBJKY", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Doc Number] [Tr Medium] [Message Type] [Language] [Created On] [Created By] [Proc Status] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()
        If Not RF Then Exit Sub

        T.TableName = "NAST"

        If Not COF Is Nothing And Not COT Is Nothing Then
            T.ParamIncludeFromTo("ERDAT", COF, COT)
        End If

        If Not DNF Is Nothing And Not DNT Is Nothing Then
            T.ParamIncludeFromTo("OBJKY", DNF, DNT)
        End If

        T.AddField("MANDT", "Client")
        T.AddField("OBJKY", "Doc Number")
        T.AddField("NACHA", "Tr Medium")
        T.AddField("KSCHL", "Message Type")
        T.AddField("SPRAS", "Language")
        T.AddField("ERDAT", "Created On")
        T.AddField("USNAM", "Created By")
        T.AddField("VSTAT", "Proc Status")

        If Not SAR Then
            T.AddKeyColumn(0)
            T.AddKeyColumn(1)
        End If

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Created On")
        End If
    End Sub

    Public Sub IncludeProcStatus(ByVal pStatus As ProcStatus)
        If Not T Is Nothing Then
            T.ParamInclude("VSTAT", pStatus)
        End If
    End Sub

    Public Sub ExcludeProcStatus(ByVal pStatus As ProcStatus)
        If Not T Is Nothing Then
            T.ParamExclude("VSTAT", pStatus)
        End If
    End Sub

    Public Sub IncludeTransMedium(ByVal Medium As String)
        If Not T Is Nothing Then
            T.ParamInclude("NACHA", Medium)
        End If
    End Sub

    Public Sub ExcludeTransMedium(ByVal Medium As String)
        If Not T Is Nothing Then
            T.ParamExclude("NACHA", Medium)
        End If
    End Sub

End Class

Public NotInheritable Class ZMXXTLOG_Report : Inherits RTable_Report

    Private MNF As String = Nothing     'Material
    Private MNT As String = Nothing
    Private MCF As String = Nothing
    Private MCT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF
        End Get

        Set(ByVal value As String)
            MNF = value
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT
        End Get

        Set(ByVal value As String)
            MNT = value
        End Set

    End Property

    Public Property MatsCreatedFrom() As String

        Get
            MatsCreatedFrom = MCF
        End Get

        Set(ByVal value As String)
            MCF = value
        End Set

    End Property

    Public Property MatsCreatedTo() As String

        Get
            MatsCreatedTo = MCT
        End Get

        Set(ByVal value As String)
            MCT = value
        End Set

    End Property

    Public Sub IncludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Plant)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Plant)
        End If

    End Sub

    Public Sub Include_SC_Indicator(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("ZSTEP", Code)
        End If

    End Sub

    Public Sub IncludeMatsCreatedOn(ByVal DDate As String)

        If Not T Is Nothing Then
            T.ParamInclude("ERSDA", ConversionUtils.NetDate2SAPDate(DDate))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Client] [Material] [Plant] [Created By]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        If Not MNF Is Nothing And Not MNT Is Nothing Then
            T.ParamIncludeFromTo("MATNR", MNF.PadLeft(18, "0"), MNT.PadLeft(18, "0"))
        End If
        If Not MCF Is Nothing And Not MCT Is Nothing Then
            T.ParamIncludeFromTo("ERSDA", ConversionUtils.NetDate2SAPDate(MCF), ConversionUtils.NetDate2SAPDate(MCT))
        End If

        T.TableName = "ZMXXTLOG"

        T.AddField("MANDT", "Client")
        T.AddField("MATNR", "Material")
        T.AddField("WERKS", "Plant")
        T.AddField("ERNAM", "Created By")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Material")
        End If

    End Sub

End Class

Public NotInheritable Class ZFIX_T03_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("MMBELNR", Number)
        End If

    End Sub

    Public Sub Exclude_Document(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("MMBELNR", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Record] [InvoiceNumber] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZFIX_T03"

        T.AddField("REFNR", "Record")
        T.AddField("MMBELNR", "InvoiceNumber")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class ZFIX_T02_Report : Inherits RTable_Report

    Private DF As String = Nothing
    Private DT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Property DateFrom() As Date

        Get
            DateFrom = ConversionUtils.SAPDate2NetDate(DF)
        End Get

        Set(ByVal value As Date)
            DF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property DateTo() As Date

        Get
            DateTo = ConversionUtils.SAPDate2NetDate(DT)
        End Get

        Set(ByVal value As Date)
            DT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Sub Include_Record(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("REFNR", Number)
        End If

    End Sub

    Public Sub Include_Date(ByVal Current As String)

        If Not T Is Nothing Then
            T.ParamInclude("ZDATE", NetDate2SAPDate(Current))
        End If

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [Record] [Status] [Current Date]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZFIX_T02"

        If Not DF Is Nothing And Not DT Is Nothing Then T.ParamIncludeFromTo("ZDATE", DF, DT)

        T.AddField("REFNR", "Record")
        T.AddField("STATUS", "Status")
        T.AddField("ZDATE", "Current Date")
        T.AddField("ZTIME", "Time")

        MyBase.Execute()

        T.ColumnToDateStr("Current Date")

    End Sub

End Class

Public NotInheritable Class CDHDR_Report : Inherits RTable_Report

    Private DF As String = Nothing
    Private DT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Property DateFrom() As Date

        Get
            DateFrom = ConversionUtils.SAPDate2NetDate(DF)
        End Get

        Set(ByVal value As Date)
            DF = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Property DateTo() As Date

        Get
            DateTo = ConversionUtils.SAPDate2NetDate(DT)
        End Get

        Set(ByVal value As Date)
            DT = ConversionUtils.NetDate2SAPDate(value)
        End Set

    End Property

    Public Sub IncludeCDO(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("OBJECTCLAS", Code)
        End If

    End Sub

    Public Sub IncludeUser(ByVal Name As String)

        If Not T Is Nothing Then
            T.ParamInclude("USERNAME", Name)
        End If

    End Sub

    Public Sub IncludeDate(ByVal ADate As Date)

        If Not T Is Nothing Then
            T.ParamInclude("UDATE", ConversionUtils.NetDate2SAPDate(ADate))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Change doc object] [Name] [Date] [Client] [Object Value] [Document Number] [Time] [Transaction] [Change Type] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "CDHDR"

        If Not DF Is Nothing AndAlso Not DT Is Nothing Then T.ParamIncludeFromTo("UDATE", DF, DT)

        T.AddField("OBJECTCLAS", "Change doc object")
        T.AddField("USERNAME", "Name")
        T.AddField("UDATE", "Date")
        T.AddField("MANDANT", "Client")
        T.AddField("OBJECTID", "Object Value")
        T.AddField("CHANGENR", "Document Number")
        T.AddField("UTIME", "Time")
        T.AddField("TCODE", "Transaction")
        T.AddField("CHANGE_IND", "Change Type")

        MyBase.Execute()

        If T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Date")
        End If

    End Sub

End Class

Public NotInheritable Class CDPOS_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeCDO(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("OBJECTCLAS", Code)
        End If

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("CHANGENR", Number)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Change doc object] [Document Number] [Client] [Object Value] [Table Name] [Table key] 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "CDPOS"

        T.AddField("OBJECTCLAS", "Change doc object")
        T.AddField("CHANGENR", "Document Number")
        T.AddField("MANDANT", "Client")
        T.AddField("OBJECTID", "Object Value")
        T.AddField("TABNAME", "Table Name")
        T.AddField("TABKEY", "Table key")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class ZMEP_PVL_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_Vendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("VENDNO", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub Include_Country(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("COUNTRYCODE", Code)
        End If

    End Sub

    Public Sub Exclude_Country(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("COUNTRYCODE", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Vendor] [Plant] [MatGrp]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZMEP_PVL"

        T.AddField("VENDNO", "Vendor")
        T.AddField("LOCNO", "Plant")
        T.AddField("PRCAT", "MatGrp")

        MyBase.Execute()

        T.ColumnToDoubleStr("Vendor")

    End Sub

End Class

Public NotInheritable Class ZTXXSTLGEX_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("LIFNR", Code.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub IncludeCCode(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("BUKRS", Code)
        End If

    End Sub

    Public Sub ExcludeCCode(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("BUKRS", Code)
        End If

    End Sub

    ''' <summary>
    ''' Returns: [Vendor] [CCode] [ExcepCode]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZTXXSTLGEX"

        T.AddField("LIFNR", "Vendor")
        T.AddField("BUKRS", "Name")
        T.AddField("ZEXCEPTION", "ExcepCode")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDoubleStr("Vendor")
        End If

    End Sub

End Class

Public NotInheritable Class ADR6_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeAddress(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("ADDRNUMBER", Number.PadLeft(10, "0"))
        End If

    End Sub

    Public Sub ExcludeAddress(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamExclude("ADDRNUMBER", Number.PadLeft(10, "0"))
        End If

    End Sub

    ''' <summary>
    ''' Returns: [AddrNum] [eMail1] [eMail2]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ADR6"

        T.AddField("ADDRNUMBER", "AddrNum")
        T.AddField("SMTP_ADDR", "eMail1")
        T.AddField("SMTP_SRCH", "eMail2")

        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class ZBBP_SC_Data_Report : Inherits RTable_Report

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub Include_TransNo(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("OBJECT_ID", Number.PadLeft(10, "0"))
        End If

    End Sub


    ''' <summary>
    ''' Returns: [TransNo] [Posting Date] [Box] [User Name] [Vendor Name]
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "ZBBP_SC_Data"

        T.AddField("OBJECT_ID", "TransNo")
        T.AddField("POSTING_DATE", "Posting Date")
        T.AddField("BE_LOG_SYSTEM", "Box")
        T.AddField("REQUESTOR_NAME", "User Name")
        T.AddField("VENDOR_NAME", "Vendor Name")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            T.ColumnToDateStr("Posting Date")
        End If

    End Sub

End Class

#End Region

#Region "SAP Transactions Reports"

Public NotInheritable Class OTD_Report : Inherits RTable_Report

    Private D As DataTable = Nothing
    Private MNF As String = Nothing
    Private MNT As String = Nothing
    Private MF As String = Nothing
    Private MT As String = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Overrides ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public Property MaterialFrom() As String

        Get
            MaterialFrom = MNF.Replace("0", "").Trim
        End Get

        Set(ByVal value As String)
            MNF = value.PadLeft(18, "0")
        End Set

    End Property

    Public Property MaterialTo() As String

        Get
            MaterialTo = MNT.Replace("0", "").Trim
        End Get

        Set(ByVal value As String)
            MNT = value.PadLeft(18, "0")
        End Set

    End Property

    Public Property MonthFrom() As String

        Get
            MonthFrom = MF
        End Get
        Set(ByVal value As String)
            MF = value
        End Set

    End Property

    Public Property MonthTo() As String

        Get
            MonthTo = MT
        End Get
        Set(ByVal value As String)
            MT = value
        End Set

    End Property

    Public Sub IncludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("WERKS", Code)
        End If

    End Sub

    Public Sub ExcludePlant(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("WERKS", Code)
        End If

    End Sub

    Public Sub IncludePOrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamInclude("EKORG", Code)
        End If

    End Sub

    Public Sub ExcludePOrg(ByVal Code As String)

        If Not T Is Nothing Then
            T.ParamExclude("EKORG", Code)
        End If

    End Sub

    Public Sub IncludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamInclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Sub ExcludeMaterial(ByVal Material As String)

        If Not T Is Nothing Then
            T.ParamExclude("MATNR", Material.PadLeft(18, "0"))
        End If

    End Sub

    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "S012"

        If Not MNF Is Nothing And Not MNT Is Nothing Then
            T.ParamIncludeFromTo("MATNR", MNF, MNT)
        End If
        If Not MF Is Nothing And Not MT Is Nothing Then
            T.ParamIncludeFromTo("SPMON", MF, MT)
        End If

        T.AddField("MATNR")
        T.AddField("WERKS")
        T.AddField("TABW1")
        T.AddField("TABW2")
        T.AddField("TABW3")
        T.AddField("TABW4")
        T.AddField("TABW5")

        T.AddKeyColumn(0)
        T.AddKeyColumn(1)

        MyBase.Execute()
        If Not SF Then Exit Sub

        If T.Result.Rows.Count <> 0 Then
            D = New DataTable
            D.Columns.Add("Material", System.Type.GetType("System.String"))
            D.Columns.Add("Del1", System.Type.GetType("System.Int16"))
            D.Columns.Add("Del2", System.Type.GetType("System.Int16"))
            D.Columns.Add("Del3", System.Type.GetType("System.Int16"))
            D.Columns.Add("Del4", System.Type.GetType("System.Int16"))
            D.Columns.Add("Del5", System.Type.GetType("System.Int16"))
            D.Columns.Add("Dummy", System.Type.GetType("System.String"))
            D.Columns.Add("Plant", System.Type.GetType("System.String"))
            D.Columns.Add("OTD", System.Type.GetType("System.Int16"))
            D.Columns("OTD").Expression = "IIF(Del1 = 0 AND Del2 = 0 AND Del3 = 0 AND Del4 = 0 AND Del5 = 0, 0, ((Del1 + Del2)/(Del1 + Del2 + Del3 + Del4 + Del5) * 100))"
            D.PrimaryKey = New DataColumn() {D.Columns("Material"), D.Columns("Plant")}

            Dim FR As DataRow
            Dim M As String

            For Each R In T.Result.Rows
                If IsNumeric(R("MATNR")) Then
                    M = CStr(CDbl(R("MATNR")))
                Else
                    M = R("MATNR")
                End If
                FR = D.Rows.Find(New Object() {M, R("WERKS")})
                If Not FR Is Nothing Then
                    FR("Del1") = FR("Del1") + Val(R("TABW1"))
                    FR("Del2") = FR("Del2") + Val(R("TABW2"))
                    FR("Del3") = FR("Del3") + Val(R("TABW3"))
                    FR("Del4") = FR("Del4") + Val(R("TABW4"))
                    FR("Del5") = FR("Del5") + Val(R("TABW5"))
                    D.AcceptChanges()
                Else
                    FR = D.NewRow
                    FR("Material") = M
                    FR("Plant") = R("WERKS")
                    FR("Del1") = Val(R("TABW1"))
                    FR("Del2") = Val(R("TABW2"))
                    FR("Del3") = Val(R("TABW3"))
                    FR("Del4") = Val(R("TABW4"))
                    FR("Del5") = Val(R("TABW5"))
                    D.Rows.Add(FR)
                End If
            Next
        End If

        RF = False

    End Sub

End Class

Public NotInheritable Class Faxing_Report

    Private Con As R3Connection = Nothing
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private INFR As Boolean = False

    Private D As DataTable = Nothing

    Private IDN() As String = Nothing   'Document Number

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim SC As New SAPConnector
        Con = SC.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            SF = True
        Else
            EM = SC.Status
        End If

    End Sub

    Sub New(ByVal Connection As Object)

        If Connection.Ping Then
            SF = True
        Else
            EM = "Connection already closed"
        End If

    End Sub

    Public Property IncludeNoFaxingRecords() As Boolean

        Get
            IncludeNoFaxingRecords = INFR
        End Get
        Set(ByVal value As Boolean)
            INFR = value
        End Set

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If IDN Is Nothing Then
            ReDim IDN(0)
        Else
            ReDim Preserve IDN(UBound(IDN) + 1)
        End If

        IDN(UBound(IDN)) = Number

    End Sub

    Public Sub Execute()

        If (Con Is Nothing) Or (IDN Is Nothing) Then Exit Sub

        Dim S As String
        Dim I As Integer
        Dim DR As DataRow
        Dim TNAST As New ReadTable(Con)
        SF = False
        Try
            TNAST.TableName = "NAST"
            TNAST.AddField("CMFPNR")
            TNAST.AddField("DATVR")
            TNAST.AddField("OBJKY")
            TNAST.AddCriteria("(KAPPL = 'EA' OR KAPPL = 'EL' OR KAPPL = 'EF' OR KAPPL = 'EV') ")
            TNAST.AddCriteria("AND NACHA = '2' AND VSTAT <> '0' AND (")
            I = 0
            For Each S In IDN
                S = "OBJKY = '" & S & "'"
                If I < IDN.GetUpperBound(0) Then
                    S = S & " OR "
                Else
                    S = S & ")"
                End If
                TNAST.AddCriteria(S)
                I += 1
            Next

            TNAST.Run()

            If TNAST.Result.Rows.Count > 0 Then
                TNAST.Result.PrimaryKey = New DataColumn() {TNAST.Result.Columns("CMFPNR")}
                Dim TCMFP As New ReadTable(Con)
                TCMFP.TableName = "CMFP"
                TCMFP.AddField("NR")
                TCMFP.AddField("MSGV1")
                TCMFP.AddCriteria("APLID = 'WFMC' ")
                TCMFP.AddCriteria("AND ARBGB = 'VN' ")
                TCMFP.AddCriteria("AND MSGTY = 'I' AND MSGNR = '095' AND (")
                I = 1
                For Each DR In TNAST.Result.Rows
                    S = "NR = '" & DR("CMFPNR") & "'"
                    If I < TNAST.Result.Rows.Count Then
                        S = S & " OR "
                    Else
                        S = S & ")"
                    End If
                    TCMFP.AddCriteria(S)
                    I += 1
                Next
                TCMFP.Run()
                TCMFP.Result.PrimaryKey = New DataColumn() {TCMFP.Result.Columns("NR")}

                Dim TSOST_Recno As New ReadTable(Con)
                TSOST_Recno.TableName = "SOST"
                TSOST_Recno.AddField("OBJNO")
                TSOST_Recno.AddField("RECNO")
                TSOST_Recno.AddField("ENTRY_DATE")
                TSOST_Recno.AddCriteria("OBJTP = 'OTF' AND SNDART = 'FAX' ")
                TSOST_Recno.AddCriteria("AND COUNTER = '10000' AND (")
                I = 1
                For Each DR In TCMFP.Result.Rows
                    S = "(OBJNO = '" & DR("MSGV1") & "' AND ENTRY_DATE = '" & TNAST.Result.Rows.Find(DR("NR"))("DATVR") & "')"
                    If I < TCMFP.Result.Rows.Count Then
                        S = S & " OR "
                    Else
                        S = S & ")"
                    End If
                    TSOST_Recno.AddCriteria(S)
                    I += 1
                Next
                TSOST_Recno.Run()
                TSOST_Recno.Result.PrimaryKey = New DataColumn() {TSOST_Recno.Result.Columns("RECNO")}

                Dim TSOST As New ReadTable(Con)
                TSOST.TableName = "SOST"
                TSOST.AddField("OBJNO")
                TSOST.AddField("COUNTER")
                TSOST.AddField("ENTRY_DATE")
                TSOST.AddField("ENTRY_TIME")
                TSOST.AddField("STAT_DATE")
                TSOST.AddField("STAT_TIME")
                TSOST.AddField("RECNO")
                TSOST.AddField("MSGID")
                TSOST.AddField("MSGTY")
                TSOST.AddField("MSGV1")
                TSOST.AddField("MSGV4")
                TSOST.AddCriteria("OBJTP = 'OTF' AND SNDART = 'FAX' AND (")
                I = 1
                For Each DR In TSOST_Recno.Result.Rows
                    S = "(RECNO = '" & DR("RECNO") & "' AND OBJNO = '" & DR("OBJNO") & "')"
                    If I < TSOST_Recno.Result.Rows.Count Then
                        S = S & " OR "
                    Else
                        S = S & ")"
                    End If
                    TSOST.AddCriteria(S)
                    I += 1
                Next
                TSOST.Run()

                TSOST.Result.PrimaryKey = New DataColumn() {TSOST.Result.Columns("OBJNO"), TSOST.Result.Columns("COUNTER")}
                TCMFP.Result.PrimaryKey = New DataColumn() {TCMFP.Result.Columns("NR")}

                Dim C
                Dim M As String
                Dim F As String
                Dim TD As Date
                Dim TT As String
                Dim Scs As Boolean
                Dim A() As DataRow
                Dim R As DataRow
                Dim FR As DataRow

                D = New DataTable
                D.Columns.Add("Document", System.Type.GetType("System.String"))
                D.Columns.Add("Recno", System.Type.GetType("System.String"))
                D.Columns.Add("Transm Date", System.Type.GetType("System.DateTime"))
                D.Columns.Add("Transm Time", System.Type.GetType("System.String"))
                D.Columns.Add("Fax", System.Type.GetType("System.String"))
                D.Columns.Add("Success", System.Type.GetType("System.Boolean"))
                D.Columns.Add("Message", System.Type.GetType("System.String"))

                For Each S In IDN
                    A = TNAST.Result.Select("OBJKY = " & S)
                    If Not A.Length = 0 Then
                        For Each R In A
                            C = R("CMFPNR")
                            C = TCMFP.Result.Rows.Find(C)
                            If Not C Is Nothing Then
                                C = C("MSGV1")
                                FR = TSOST.Result.Rows.Find(New Object() {C, "10002"})
                                If Not FR Is Nothing Then
                                    If FR("MSGTY") = "E" Then
                                        Scs = False
                                    Else
                                        Scs = True
                                    End If
                                    F = FR("MSGV1")
                                    M = FR("MSGV4")
                                    TD = ConversionUtils.SAPDate2NetDate(FR("ENTRY_DATE"))
                                    TT = FR("ENTRY_TIME")
                                    TT = Left(TT, 2) & ":" & Mid(TT, 3, 2) & ":" & Right(TT, 2)
                                    D.Rows.Add(New Object() {S, Val(C), TD, TT, F, Scs, M})
                                End If
                            End If
                        Next
                    Else
                        If INFR Then
                            D.Rows.Add(New Object() {S, "0", Nothing, Nothing, Nothing, False, "No Faxing Records"})
                        End If
                    End If
                Next
            Else
                If INFR Then
                    D = New DataTable
                    D.Columns.Add("Document", System.Type.GetType("System.String"))
                    D.Columns.Add("Recno", System.Type.GetType("System.String"))
                    D.Columns.Add("Transm Date", System.Type.GetType("System.DateTime"))
                    D.Columns.Add("Transm Time", System.Type.GetType("System.String"))
                    D.Columns.Add("Fax", System.Type.GetType("System.String"))
                    D.Columns.Add("Success", System.Type.GetType("System.Boolean"))
                    D.Columns.Add("Message", System.Type.GetType("System.String"))
                    For Each S In IDN
                        D.Rows.Add(New Object() {S, "0", Nothing, Nothing, Nothing, False, "No Faxing Records"})
                    Next
                End If
            End If
            SF = True
        Catch ex As Exception
            EM = ex.Message
        End Try

    End Sub

    Public Sub EndSession()

        Con.Close()
        Con = Nothing

    End Sub

End Class

#End Region

#Region "PSS Custom Reports/Classes"

Public NotInheritable Class UMOF_Vendor_Status

    Private T0 As RTable = Nothing
    Private T1 As RTable = Nothing
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private Con As R3Connection = Nothing
    Private Vendor As String
    Private POrg As String
    Private CCode As String
    Private DT As DataTable = Nothing

    Sub New(ByVal Con)

        If Con.Ping Then
            T0 = New RTable(Con)
            T1 = New RTable(Con)
            SF = True
        Else
            EM = "Conection already closed"
        End If

    End Sub

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = DT
        End Get

    End Property

    Public Sub Include(ByVal AVendor As String, ByVal APOrg As String, ByVal ACCode As String)

        Vendor = AVendor
        POrg = APOrg
        CCode = ACCode

    End Sub

    Public Sub Execute()

        DT = New DataTable
        DT.Columns.Add("POrg_Link", System.Type.GetType("System.Boolean"))
        DT.Columns.Add("PO_Link_Blk", System.Type.GetType("System.Boolean"))
        DT.Columns.Add("CC_Link", System.Type.GetType("System.Boolean"))
        DT.Columns.Add("CC_Link_Blk", System.Type.GetType("System.Boolean"))
        DT.Columns.Add("Regi", System.Type.GetType("System.Boolean"))
        DT.Columns.Add("ERS", System.Type.GetType("System.Boolean"))
        Dim DR As DataRow = DT.NewRow

        T0.TableName = "LFM1"
        T0.AddField("SPERM")
        T0.AddField("LOEVM")
        T0.AddField("XERSY")
        T0.ParamInclude("LIFNR", Left("0000000000", 10 - Len(Vendor)) & Vendor)
        T0.ParamInclude("EKORG", Left("0000", 4 - Len(POrg)) & POrg)

        T1.TableName = "LFB1"
        T1.AddField("SPERR")
        T1.AddField("LOEVM")
        T1.AddField("BEGRU")
        T1.ParamInclude("LIFNR", Left("0000000000", 10 - Len(Vendor)) & Vendor)
        T1.ParamInclude("BUKRS", Left("000", 3 - Len(CCode)) & CCode)

        T0.Run()
        If T0.Success AndAlso T0.Result.Rows.Count > 0 Then
            DR("POrg_Link") = T0.Result.Rows(0)("SPERM") <> "X"
            DR("PO_Link_Blk") = T0.Result.Rows(0)("LOEVM") = "X"
            DR("ERS") = T0.Result.Rows(0)("XERSY") = "X"
            T1.Run()
            If T1.Success And T1.Result.Rows.Count > 0 Then
                DR("CC_Link") = T1.Result.Rows(0)("SPERR") <> "X"
                DR("CC_Link_Blk") = T1.Result.Rows(0)("LOEVM") = "X"
                DR("Regi") = T1.Result.Rows(0)("BEGRU").ToString.Contains("REGI")
            Else
                DR("CC_Link") = False
                DR("CC_Link_Blk") = False
                DR("Regi") = False
            End If
        Else
            DR("POrg_Link") = False
            DR("PO_Link_Blk") = False
            DR("CC_Link") = False
            DR("CC_Link_Blk") = False
            DR("Regi") = False
        End If
        DT.Rows.Add(DR)
        DT.AcceptChanges()

    End Sub

End Class

Public NotInheritable Class BI_ExtData_Report : Inherits RTable_Report

    Private D As DataTable = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Overrides ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public Sub IncludeDocument(ByVal Number As String)

        If Not T Is Nothing Then
            T.ParamInclude("BELNR", Number)
        End If

    End Sub

    Public Overrides Sub Execute()

        If Not RF Then Exit Sub

        T.TableName = "RBKP"

        T.AddField("BELNR")
        T.AddField("ZFBDT")
        T.AddField("GJAHR")
        T.AddField("LIFNR")
        T.AddField("XBLNR")

        MyBase.Execute()

        If SF AndAlso T.Result.Rows.Count > 0 Then
            Dim TBKPF As New RTable(Con)
            Dim S As String
            Dim I As Integer
            Dim OP As String = Nothing
            Dim R As DataRow

            TBKPF.TableName = "BKPF"
            TBKPF.AddField("BELNR")
            TBKPF.AddField("AWKEY")
            TBKPF.AddCriteria("AWTYP = 'RMRP' AND (")
            I = 1
            For Each R In T.Result.Rows
                S = "(AWKEY = '" & R("BELNR") & R("GJAHR") & "')"
                If I < T.Result.Rows.Count Then
                    S = S & " OR "
                Else
                    S = S & ")"
                End If
                TBKPF.AddCriteria(S)
                I += 1
            Next
            TBKPF.Run()

            Dim VR As New LFA1_Report(Con)
            For Each R In T.Result.Rows
                VR.IncludeVendor(R("LIFNR"))
            Next
            VR.Execute()

            D = New DataTable
            D.Columns.Add("InvoiceNumber", System.Type.GetType("System.String"))
            D.Columns.Add("Vendor", System.Type.GetType("System.String"))
            D.Columns.Add("VendorName", System.Type.GetType("System.String"))
            D.Columns.Add("BaselineDate", System.Type.GetType("System.String"))
            D.Columns.Add("ReferenceDoc", System.Type.GetType("System.String"))
            D.Columns.Add("FINumber", System.Type.GetType("System.String"))
            Dim NR As DataRow
            For Each R In T.Result.Rows
                NR = D.NewRow
                NR("InvoiceNumber") = R("BELNR")
                NR("Vendor") = CStr(CDbl(R("LIFNR")))
                NR("VendorName") = VR.Data.Select("Vendor = '" & NR("Vendor") & "'")(0)("Name")
                NR("BaselineDate") = ConversionUtils.SAPDate2NetDate(R("ZFBDT")).ToShortDateString
                NR("ReferenceDoc") = R("XBLNR")
                NR("FINumber") = TBKPF.Result.Select("AWKEY = '" & R("BELNR") & R("GJAHR") & "'")(0)("BELNR")
                D.Rows.Add(NR)
            Next
            SF = True
        End If

    End Sub

End Class

Public NotInheritable Class Material_In_TDB

    Private T As RTable = Nothing
    Private D As DataTable = Nothing
    Private Con As R3Connection = Nothing
    Private A(,) As String = Nothing
    Private SF As Boolean = False
    Private EM As String = Nothing

    Private Structure Params
        Public Material As String
        Public Plant As String
    End Structure

    Private IMP() As Params = Nothing

    Public ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public ReadOnly Property DataArray()

        Get
            If Not D Is Nothing Then
                If A Is Nothing Then
                    Dim R As DataRow
                    Dim X As Integer
                    Dim I As Integer
                    ReDim A(D.Rows.Count, D.Columns.Count)
                    For I = 1 To D.Columns.Count
                        A(0, I) = D.Columns(I - 1).ColumnName
                    Next
                    I = 1
                    For Each R In D.Rows
                        For X = 1 To D.Columns.Count
                            A(I, X) = R(X - 1).ToString
                        Next
                        I += 1
                    Next
                End If
            End If
            DataArray = A
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String
        Get
            ErrMessage = EM
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        OpenSession(Box, User, App)

    End Sub

    Public Sub Include(ByVal Material As String, ByVal Plant As String)

        If IMP Is Nothing Then
            ReDim IMP(0)
        Else
            ReDim Preserve IMP(IMP.GetUpperBound(0) + 1)
        End If

        IMP(IMP.GetUpperBound(0)).Material = Left("000000000000000000", 18 - Len(Material)) & Material
        IMP(IMP.GetUpperBound(0)).Plant = Plant

    End Sub

    Public Sub Execute()

        Dim S As String
        Dim I As Integer
        Dim OP As String

        If Not IMP Is Nothing Then
            For I = 0 To IMP.GetUpperBound(0)
                S = IMP(I).Material
                If InStr(S, "*") <> 0 Then
                    S = Replace(S, "*", "%")
                    OP = " LIKE "
                Else
                    OP = " = "
                End If
                S = "(MATNR" & OP & "'" & S & "' AND "
                If InStr(IMP(I).Plant, "*") <> 0 Then
                    IMP(I).Plant = Replace(IMP(I).Plant, "*", "%")
                    OP = " LIKE "
                Else
                    OP = " = "
                End If
                S = S & "WERKS" & OP & "'" & IMP(I).Plant & "')"
                If I < IMP.GetUpperBound(0) Then
                    S = S & " OR "
                End If
                T.AddCriteria(S)
            Next
            T.AddField("MATNR")
            T.AddField("WERKS")
            T.AddKeyColumn(0)
            T.AddKeyColumn(1)
            T.Run()
            If Not T.Success Then
                EM = T.ErrMessage
                SF = False
                Exit Sub
            Else
                SF = True
            End If
            If T.Result.Rows.Count > 0 Then
                T.Result.Columns("MATNR").ColumnName = "Material"
                T.Result.Columns("WERKS").ColumnName = "Plant"
                Dim R As DataRow
                For Each R In T.Result.Rows
                    If IsNumeric(R("Material")) Then
                        R("Material") = CStr(CDbl(R("Material")))
                    Else
                        R("Material") = Nothing
                    End If
                    R.AcceptChanges()
                Next
            End If
            D = New DataTable("MITDB")
            D.Columns.Add("Material", System.Type.GetType("System.String"))
            D.Columns.Add("Plant", System.Type.GetType("System.String"))
            D.Columns.Add("InTDB", System.Type.GetType("System.Boolean"))
            D.PrimaryKey = New DataColumn() {D.Columns("Material"), D.Columns("Plant")}
            For I = 0 To IMP.GetUpperBound(0)
                If T.Result.Rows.Find(New Object() {CStr(CDbl(IMP(I).Material)), IMP(I).Plant}) Is Nothing Then
                    D.LoadDataRow(New Object() {CStr(CDbl(IMP(I).Material)), IMP(I).Plant, False}, LoadOption.PreserveChanges)
                Else
                    D.LoadDataRow(New Object() {CStr(CDbl(IMP(I).Material)), IMP(I).Plant, True}, LoadOption.PreserveChanges)
                End If
            Next
        End If

    End Sub

    Public Sub OpenSession(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim SC As New SAPConnector
        Con = SC.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            T = New RTable(Con)
            T.TableName = "ZMXXMRC1"
            SF = True
        Else
            EM = SC.Status
        End If

    End Sub

End Class

Public NotInheritable Class GR_IR_Report : Inherits RTable_Report

    Private D As DataTable = Nothing
    Private Docs As DataTable = Nothing

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal Connection As Object)

        MyBase.New(Connection)

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        T.ParamInclude("EBELN", Number)

    End Sub

    Public Overrides ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public Overrides Sub Execute()

        T.TableName = "EKBE"

        T.AddCriteria("(VGABE = 1 OR VGABE = 2) AND ")

        T.AddField("EBELN") 'Doc Number
        T.AddField("EBELP") 'Item Number
        T.AddField("VGABE") 'Trans. Type
        T.AddField("MENGE") 'Amount
        T.AddField("SHKZG") 'Deb/Cred Ind
        T.AddField("BUDAT") 'Posting Date

        MyBase.Execute()

        D = New DataTable("GRIR")
        D.Columns.Add("PO", System.Type.GetType("System.String"))
        D.Columns.Add("Item", System.Type.GetType("System.String"))
        D.Columns.Add("GR", System.Type.GetType("System.Double"))
        D.Columns.Add("IR", System.Type.GetType("System.Double"))
        D.Columns.Add("Imbalance", Type.GetType("System.String"), "IIF (GR <> IR, 'Yes', 'No')")
        D.Columns.Add("Last GR Date", Type.GetType("System.String"))
        If T.Result.Rows.Count > 0 Then
            Docs = New DataTable
            Docs.Columns.Add("EBELN", System.Type.GetType("System.String"))
            Docs.Columns.Add("EBELP", System.Type.GetType("System.String"))
            Docs.PrimaryKey = New DataColumn() {Docs.Columns(0), Docs.Columns(1)}
            Dim R As DataRow
            For Each R In T.Result.Rows
                Docs.LoadDataRow(New Object() {R("EBELN"), R("EBELP")}, LoadOption.PreserveChanges)
            Next
            Dim FR() As DataRow
            Dim NR As DataRow
            Dim RR As DataRow
            Dim Val As Double
            For Each R In Docs.Rows
                FR = T.Result.Select("EBELN = '" & R("EBELN") & "' AND EBELP = '" & R("EBELP") & "'")
                NR = D.NewRow
                NR("PO") = R("EBELN")
                NR("Item") = CStr(CInt(R("EBELP")))
                NR("IR") = 0
                NR("GR") = 0
                For Each RR In FR
                    If IsNumeric(RR("MENGE")) Then
                        Val = CDbl(RR("MENGE"))
                    Else
                        Val = 0
                    End If
                    If RR("SHKZG") = "H" Then
                        Val = Val * -1
                    End If
                    If RR("VGABE") = 1 Then
                        NR("GR") = NR("GR") + Val
                        If RR("SHKZG") = "S" Then
                            If DBNull.Value.Equals(NR("Last GR Date")) OrElse CDate(NR("Last GR Date")) < SAPDate2NetDate(RR("BUDAT")) Then
                                NR("Last GR Date") = SAPDate2NetDate(RR("BUDAT")).ToShortDateString
                            End If
                        End If
                    Else
                        NR("IR") = NR("IR") + Val
                    End If
                Next
                D.Rows.Add(NR)
            Next
        End If

    End Sub

End Class

Public NotInheritable Class Agreement_Report : Inherits LINQ_Support

    Private Con As R3Connection = Nothing
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private RF As Boolean = False

    Private IOA As Boolean = False
    Private ISA As Boolean = False
    Private ARDT As DataTable = Nothing
    Private Buyers As DataTable = Nothing
    Private EKKO As DataTable = Nothing
    Private EKPO As DataTable = Nothing
    Private LFA1 As DataTable = Nothing
    Private T001W As DataTable = Nothing
    Private Total As Integer = 0
    Private Count As Integer = 0
    Private IVC() As String = Nothing   'Vendor Code
    Private IPG() As String = Nothing   'Purch Group
    Private IPC() As String = Nothing   'Plant Code

    Public Event Report_Progress(ByVal Percent As Integer, ByVal Msg As String)

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim SC As New SAPConnector
        Con = SC.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            RF = True
        Else
            EM = SC.Status
        End If

    End Sub

    Public Sub New(ByVal Connection)

        If Not Connection Is Nothing AndAlso Connection.Ping Then
            Con = Connection
            RF = True
        Else
            EM = "Connection already closed"
        End If

    End Sub

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = RF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = ARDT
        End Get

    End Property

    Public Property Include_OAs() As Boolean

        Get
            Include_OAs = IOA
        End Get
        Set(ByVal value As Boolean)
            IOA = value
        End Set

    End Property

    Public Property Include_SAs() As Boolean

        Get
            Include_SAs = ISA
        End Get
        Set(ByVal value As Boolean)
            ISA = value
        End Set

    End Property

    Public Sub IncludeVendor(ByVal Code As String)

        If IVC Is Nothing Then
            ReDim IVC(0)
        Else
            ReDim Preserve IVC(UBound(IVC) + 1)
        End If

        IVC(UBound(IVC)) = Code.PadLeft(10, "0")

    End Sub

    Public Sub IncludePurchGrp(ByVal Code As String)

        If IPG Is Nothing Then
            ReDim IPG(0)
        Else
            ReDim Preserve IPG(UBound(IPG) + 1)
        End If

        IPG(UBound(IPG)) = Code.Trim.ToUpper

    End Sub

    Public Sub IncludePlant(ByVal Code As String)

        If IPC Is Nothing Then
            ReDim IPC(0)
        Else
            ReDim Preserve IPC(UBound(IPC) + 1)
        End If

        IPC(UBound(IPC)) = Code

    End Sub

    ''' <summary>
    ''' Returns:
    '''    [PGrp] [Buyer_Name] [Agreement] [Vendor] [Vendor_Name] [Vendor_Country] [Plant] [Plant_Name] [Plant_Country] [Item] [Material] [Short_Text] [Net_Price] 
    '''    [Price_Unit] [Order_Unit] [Currency] [Incoterms] [Incoterms_Desc] [POrg] [Pmnt_Terms] [Comp_Code] [Doc_Type] [Validity_Start] [Validity_End] [Tax_Code]
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Execute()

        If Not RF Then Exit Sub
        If Not IOA And Not ISA Then
            SF = False
            EM = "Include_SAs and Include_OAs both are set to false."
            Exit Sub
        End If

        Dim T As DataTable

        Total = 3
        If Not IPG Is Nothing Then Total = Total + IPG.Count
        If Not IVC Is Nothing Then Total = Total + IVC.Count
        If Not IPC Is Nothing Then Total = Total + IPC.Count

        If Not IPG Is Nothing Then
            For Each PGrp In IPG
                T = EKKO_Table(PGrp, 0)
                If Not T Is Nothing Then
                    If EKKO Is Nothing Then
                        EKKO = T.Copy
                    Else
                        EKKO.Merge(T)
                    End If
                    T = EKPO_Table(NoDupList(T, "Doc Number"), PGrp)
                    If Not T Is Nothing Then
                        If EKPO Is Nothing Then
                            EKPO = T.Copy
                        Else
                            EKPO.Merge(T)
                        End If
                    End If
                End If
                ReportProgress()
            Next
        End If

        If Not IVC Is Nothing And IPG Is Nothing Then
            For Each Vendor As String In IVC
                T = EKKO_Table(Vendor, 1)
                If Not T Is Nothing Then
                    If EKKO Is Nothing Then
                        EKKO = T.Copy
                    Else
                        EKKO.Merge(T)
                    End If
                    T = EKPO_Table(NoDupList(T, "Doc Number"), Vendor)
                    If Not T Is Nothing Then
                        If EKPO Is Nothing Then
                            EKPO = T.Copy
                        Else
                            EKPO.Merge(T)
                        End If
                    End If
                End If
                ReportProgress()
            Next
        End If

        If Not IPC Is Nothing And IPG Is Nothing And IVC Is Nothing Then
            For Each Plant As String In IPC
                T = EKPO_Table(Plant)
                If Not T Is Nothing Then
                    If EKPO Is Nothing Then
                        EKPO = T.Copy
                    Else
                        EKPO.Merge(T)
                    End If
                    T = EKKO_Table(NoDupList(T, "Doc Number"), Plant)
                    If Not T Is Nothing Then
                        If EKKO Is Nothing Then
                            EKKO = T.Copy
                        Else
                            EKKO.Merge(T)
                        End If
                    End If
                End If
                ReportProgress()
            Next
        End If

        If EKKO Is Nothing Or EKPO Is Nothing Then Exit Sub

        SF = True
        RF = False
        Dim DL As String() = NoDupList(EKKO, "Doc Number")
        Dim PL As String() = NoDupList(EKPO, "Plant")

        LFA1 = LFA1_Table(NoDupList(EKKO, "Vendor"))
        ReportProgress()
        T001W = T001W_Table(PL)
        ReportProgress()
        Buyers = Buyers_Table()
        ReportProgress()

        Get_ARDT()

    End Sub

    Private Sub ReportProgress()

        Count += 1
        RaiseEvent Report_Progress(Int((Count / Total) * 100), Nothing)

    End Sub

    Private Function Buyers_Table() As DataTable

        Buyers_Table = Nothing
        If Con Is Nothing Then Exit Function

        Dim B As New Buyers_Report(Con)
        B.Execute()
        If B.Success AndAlso B.Data.Rows.Count > 0 Then
            Buyers_Table = B.Data
        End If

    End Function

    Private Function EKKO_Table(ByVal Param As String, ByVal PType As Byte) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading EKKO (" & Param & ")")

        Dim E As New EKKO_Report(Con)
        E.DeletionIndicator = False
        If IOA Then E.IncludeDocCategory("K")
        If ISA Then E.IncludeDocCategory("L")
        If PType = 0 Then
            E.IncludePurchGroup(Param)
            If Not IVC Is Nothing Then
                For Each Vendor As String In IVC
                    E.IncludeVendor(Vendor)
                Next
            End If
        Else
            E.IncludeVendor(Param)
        End If
        E.AddCustomField("INCO1", "Incoterms")
        E.AddCustomField("INCO2", "Incoterm Desc")
        E.AddCustomField("ZTERM", "Pmnt Terms")
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function EKKO_Table(ByVal Documents() As String, ByVal Param As String) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading EKKO (" & Param & ")")

        Dim E As New EKKO_Report(Con)
        E.DeletionIndicator = False
        If IOA Then E.IncludeDocCategory("K")
        If ISA Then E.IncludeDocCategory("L")
        For Each Doc As String In Documents
            E.IncludeDocument(Doc)
        Next
        E.AddCustomField("INCO1", "Incoterms")
        E.AddCustomField("INCO2", "Incoterm Desc")
        E.AddCustomField("ZTERM", "Pmnt Terms")
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function EKPO_Table(ByVal Documents() As String, ByVal Param As String) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading EKPO (" & Param & ")")

        Dim E As New EKPO_Report(Con)
        E.DeletionIndicator = False
        If IOA Then E.IncludeDocCategory("K")
        If ISA Then E.IncludeDocCategory("L")
        For Each Doc As String In Documents
            E.IncludeDocument(Doc)
        Next
        If Not IPC Is Nothing Then
            For Each Plant As String In IPC
                E.IncludePlant(Plant)
            Next
        End If
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function EKPO_Table(ByVal Plant As String) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading EKPO (" & Plant & ")")

        Dim E As New EKPO_Report(Con)
        E.DeletionIndicator = False
        If IOA Then E.IncludeDocCategory("K")
        If ISA Then E.IncludeDocCategory("L")
        E.IncludePlant(Plant)
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function LFA1_Table(ByVal Vendors() As String) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading Vendor Data")

        Dim E As New LFA1_Report(Con)
        For Each Vdr As String In Vendors
            E.IncludeVendor(Vdr)
        Next
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function T001W_Table(ByVal Plants() As String) As DataTable

        If Con Is Nothing Then Return Nothing

        RaiseEvent Report_Progress(0, "Downloading Plant Data")

        Dim E As New Plants_Report(Con)
        For Each Plant As String In Plants
            E.IncludePlant(Plant)
        Next
        E.Execute()
        If E.Success AndAlso E.Data.Rows.Count > 0 Then
            Return E.Data
        Else
            Return Nothing
        End If

    End Function

    Private Function NoDupList(ByVal DT As DataTable, ByVal ColName As String) As String()

        Dim T As New DataTable
        Dim R(0) As String
        Dim I As Integer = 0
        T.Columns.Add("C", DT.Columns(ColName).DataType)
        T.PrimaryKey = New DataColumn() {T.Columns("C")}
        For Each DR As DataRow In DT.Rows
            If Not DBNull.Value.Equals(DR(ColName)) Then T.LoadDataRow(New Object() {DR(ColName)}, LoadOption.PreserveChanges)
        Next
        For Each DR As DataRow In T.Rows
            R(R.GetUpperBound(0)) = DR("C")
            If I < T.Rows.Count - 1 Then
                I += 1
                ReDim Preserve R(I)
            End If
        Next
        NoDupList = R

    End Function

    Private Sub Get_ARDT()

        If Not SF Then Exit Sub
        If EKKO Is Nothing Or EKPO Is Nothing Or LFA1 Is Nothing Or T001W Is Nothing Or Buyers Is Nothing Then Exit Sub

        Dim Query = From REKKO In EKKO _
        Join REKPO In EKPO On REKKO("Client") Equals REKPO("Client") And REKKO("Doc Number") Equals REKPO("Doc Number") _
        Join RLFA1 In LFA1 On REKKO("Client") Equals RLFA1("Client") And REKKO("Vendor") Equals RLFA1("Vendor") _
        Join RPGRP In Buyers On REKKO("PGrp") Equals RPGRP("PGrp") _
        Join RT001W In T001W On REKPO("Plant") Equals RT001W("Plant") _
        Select New With { _
        .PGrp = REKKO("PGrp"), _
        .Buyer_Name = RPGRP("Description"), _
        .Agreement = REKKO("Doc Number"), _
        .Vendor = REKKO("Vendor"), _
        .Vendor_Name = RLFA1("Name"), _
        .Vendor_Country = RLFA1("Country"), _
        .Plant = REKPO("Plant"), _
        .Plant_Name = RT001W("Name"), _
        .Plant_Country = RT001W("Country"), _
        .Item = REKPO("Item Number"), _
        .Material = REKPO("Material"), _
        .Short_Text = REKPO("Short Text"), _
        .Net_Price = REKPO("Price"), _
        .Price_Unit = REKPO("Price Unit"), _
        .Order_Unit = REKPO("UOM"), _
        .Currency = REKKO("Currency"), _
        .Incoterms = REKKO("Incoterms"), _
        .Incoterms_Desc = REKKO("Incoterm Desc"), _
        .POrg = REKKO("POrg"), _
        .Pmnt_Terms = REKKO("Pmnt Terms"), _
        .Comp_Code = REKKO("Company Code"), _
        .Doc_Type = REKKO("Doc Type"), _
        .Validity_Start = REKKO("Validity Start"), _
        .Validity_End = REKKO("Validity End"), _
        .Tax_Code = REKPO("Tax code") _
        }

        Try
            ARDT = LinQToDataTable(Query)
        Catch ex As Exception
            SF = False
            EM = ex.Message
        End Try

    End Sub

End Class

#End Region

#Region "SAP Query Reports"

Public MustInherit Class OO_Query

    Private Con As R3Connection = Nothing
    Private Q As Query = Nothing
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private A(,) As String = Nothing
    Private D() As String = Nothing
    Private IDD As Boolean = False
    Private IGRIR As Boolean = False
    Private IYORF As Boolean = False
    Private IPTRM As Boolean = False
    Private IOA As Boolean = False
    Private RSAS As Boolean = False
    Private DNT As DataTable = Nothing
    Private RLV As Byte = 0

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim C As New SAPConnector
        Con = C.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            InitBAPI()
        Else
            EM = C.Status
        End If

    End Sub

    Sub New(ByVal Connection As Object)

        If Connection.Ping Then
            Con = Connection
            InitBAPI()
        Else
            EM = "Connection already closed"
        End If

    End Sub

    Public ReadOnly Property DataArray()

        Get
            If Not Q Is Nothing Then
                If A Is Nothing Then
                    Dim R As DataRow
                    Dim I As Integer
                    Dim X As Integer
                    ReDim A(Q.Result.Rows.Count, Q.Result.Columns.Count)
                    For I = 1 To Q.Result.Columns.Count
                        A(0, I) = Q.Result.Columns(I - 1).ColumnName
                    Next
                    I = 1
                    For Each R In Q.Result.Rows
                        For X = 1 To Q.Result.Columns.Count
                            A(I, X) = R(X - 1).ToString
                        Next
                        I += 1
                    Next
                End If
            End If
            DataArray = A
        End Get

    End Property

    Public ReadOnly Property DocsArray()

        Get
            If Not DNT Is Nothing Then
                If D Is Nothing Then
                    ReDim D(DNT.Rows.Count)
                    Dim I As Integer = 1
                    Dim R As DataRow
                    For Each R In DNT.Rows
                        D(I) = R(0)
                        I += 1
                    Next
                End If
            End If
            DocsArray = D
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            If Not Q Is Nothing Then
                Data = Q.Result
            Else
                Data = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Documents() As DataTable

        Get
            Documents = DNT
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public Property Report_SAs() As Boolean

        Get
            Report_SAs = RSAS
        End Get
        Set(ByVal value As Boolean)
            RSAS = value
        End Set

    End Property

    Public Property IncludeDelivDates() As Boolean

        Get
            IncludeDelivDates = IDD
        End Get
        Set(ByVal value As Boolean)
            IDD = value
        End Set

    End Property

    Public Property Include_GR_IR() As Boolean

        Get
            Include_GR_IR = IGRIR
        End Get
        Set(ByVal value As Boolean)
            IGRIR = value
        End Set

    End Property

    Public Property Include_YO_Ref() As Boolean

        Get
            Include_YO_Ref = IYORF
        End Get
        Set(ByVal value As Boolean)
            IYORF = value
        End Set

    End Property

    Public Property Include_PTerms() As Boolean

        Get
            Include_PTerms = IPTRM
        End Get
        Set(ByVal value As Boolean)
            IPTRM = value
        End Set

    End Property

    Public Property Include_OAs() As Boolean

        Get
            Include_OAs = IOA
        End Get
        Set(ByVal value As Boolean)
            IOA = value
        End Set

    End Property

    Public Property RepairsLevel() As Byte

        Get
            RepairsLevel = RLV
        End Get
        Set(ByVal value As Byte)
            RLV = value
        End Set

    End Property

    Private Sub IncludeParamFromTo(ByVal Param As String, ByVal LowV As String, ByVal HighV As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Include, RangeOption.Between, LowV, HighV)
        End If

    End Sub

    Private Sub IncludeParam(ByVal Param As String, ByVal Value As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Include, RangeOption.Equals, Value)
        End If

    End Sub

    Private Sub ExcludeParam(ByVal Param As String, ByVal Value As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Exclude, RangeOption.Equals, Value)
        End If

    End Sub

    Private Sub InitBAPI()

        Try
            Q = Con.CreateQuery(WorkSpace.GlobalArea, "/SAPQUERY/ME", "MEPO")
            IncludeParam("SP$00012", "NB")
            IncludeParam("SP$00012", "EC")
            DNT = New DataTable("Doc Numbers")
            DNT.Columns.Add("Doc Number", System.Type.GetType("System.String"))
            Dim Keys(0) As DataColumn
            Keys(0) = DNT.Columns("Doc Number")
            DNT.PrimaryKey = Keys
        Catch ex As Exception
            Q = Nothing
            EM = ex.Message
        End Try

    End Sub

    Friend Sub SelectCriteria(ByVal Code As String)

        IncludeParam("SP$00023", Code)

    End Sub

    Public Sub IncludePlant(ByVal Plant As String)

        IncludeParam("SP$00022", Plant)

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        ExcludeParam("SP$00022", Plant)

    End Sub

    Public Sub IncludeDocumentFromTo(ByVal DocFrom As String, ByVal DocTo As String)

        IncludeParamFromTo("SP$00014", DocFrom, DocTo)

    End Sub

    Public Sub IncludeDocsDatedFromTo(ByVal DateFrom As Date, ByVal DateTo As Date)

        IncludeParamFromTo("SP$00011", ERPConnect.ConversionUtils.NetDate2SAPDate(DateFrom), ERPConnect.ConversionUtils.NetDate2SAPDate(DateTo))

    End Sub

    Public Sub IncludeDocument(ByVal Number As String)

        IncludeParam("SP$00014", Number)

    End Sub

    Public Sub ExcludeDocument(ByVal Number As String)

        ExcludeParam("SP$00014", Number)

    End Sub

    Public Sub IncludePurchGroup(ByVal Code As String)

        IncludeParam("SP$00015", Code)

    End Sub

    Public Sub ExcludePurchGroup(ByVal Code As String)

        ExcludeParam("SP$00015", Code)

    End Sub

    Public Sub IncludePurchOrg(ByVal Code As String)

        IncludeParam("SP$00016", Code)

    End Sub

    Public Sub ExcludePurchOrg(ByVal Code As String)

        ExcludeParam("SP$00016", Code)

    End Sub

    Public Sub IncludeVendor(ByVal Code As String)

        IncludeParam("SP$00017", Left("0000000000", 10 - Len(Code)) & Code)

    End Sub

    Public Sub ExcludeVendor(ByVal Code As String)

        ExcludeParam("SP$00017", Left("0000000000", 10 - Len(Code)) & Code)

    End Sub

    Public Sub IncludeMatGroup(ByVal Code As String)

        IncludeParam("SP$00019", Code)

    End Sub

    Public Sub ExcludeMatGroup(ByVal Code As String)

        ExcludeParam("SP$00019", Code)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        IncludeParam("S_MATNR", Number)

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        ExcludeParam("S_MATNR", Number)

    End Sub

    Public Overridable Sub Execute()

        If Q Is Nothing Then Exit Sub

        Try
            IncludeParam("P_QCOUNT", "0")
            If RSAS Then
                IncludeParam("SP$00012", "LP")
                IncludeParam("SP$00012", "ZLP")
            End If
            Q.Execute()
            If Q.Result.Rows.Count > 0 Then
                Q.Result.Columns("EKKO-EBELN").ColumnName = "Doc Number"
                Q.Result.Columns("EKPO-EBELP").ColumnName = "Item Number"
                Q.Result.Columns("EKPO-MATKL").ColumnName = "Mat Group"
                Q.Result.Columns("EKPO-MATNR").ColumnName = "Material"
                Q.Result.Columns("EKPO-TXZ01").ColumnName = "Short Text"
                Q.Result.Columns("EKKO-LIFNR").ColumnName = "Vendor"
                Q.Result.Columns("LIEFNAM").ColumnName = "Vendor Name"
                Q.Result.Columns("EKKO-BUKRS").ColumnName = "Company Code"
                Q.Result.Columns("EKKO-EKORG").ColumnName = "Purch Org"
                Q.Result.Columns("EKKO-EKGRP").ColumnName = "Purch Grp"
                Q.Result.Columns("EKPO-WERKS").ColumnName = "Plant"
                Q.Result.Columns("EKKO-BEDAT").ColumnName = "Doc Date"
                Q.Result.Columns("EKKO-BSART").ColumnName = "Doc Type"
                Q.Result.Columns("EKKO-ERNAM").ColumnName = "Created By"
                Q.Result.Columns("EKPO-BEDNR").ColumnName = "Tracking Field"
                Q.Result.Columns("EKPO-MENGE").ColumnName = "Quantity"
                Q.Result.Columns("EKPO-MEINS").ColumnName = "UOM"
                Q.Result.Columns("EKKO-WAERS").ColumnName = "Currency"
                Q.Result.Columns("EKPO-LOEKZ").ColumnName = "Del Indicator"
                Q.Result.Columns("EKPO-ELIKZ").ColumnName = "Delivery Comp"
                Q.Result.Columns("EKPO-EREKZ").ColumnName = "Final Invoice"
                Q.Result.Columns("EKPO-AFNAM").ColumnName = "Requisitioner"
                Q.Result.Columns("EKPO-NETPR").ColumnName = "Price"
                Q.Result.Columns.Remove("EKKO-RESWK")
                Q.Result.Columns.Remove("EKPO-LGORT")
                Q.Result.Columns.Remove("EKKO-BSTYP")
                Q.Result.Columns.Remove("EKKO-STATU")
                Q.Result.Columns.Remove("EKPO-KTMNG")
                Q.Result.Columns.Remove("EKPO-MEINS-0120")
                Q.Result.Columns.Remove("EKPO-MEINS-0121")
                Q.Result.Columns.Remove("EKPO-NETWR")
                Q.Result.Columns.Remove("EKKO-WAERS-0202")
                Q.Result.Columns.Remove("EKPO-LEWED")
                Q.Result.Columns.Remove("EKKO-SUBMI")
                Q.Result.Columns.Remove("EKKO-KDATB")
                Q.Result.Columns.Remove("EKKO-KDATE")
                Q.Result.Columns.Remove("EKKO-FRGGR")
                Q.Result.Columns.Remove("EKKO-FRGKE")
                Q.Result.Columns.Remove("EKKO-FRGRL")
                Q.Result.Columns.Remove("EKKO-FRGSX")
                Q.Result.Columns.Remove("EKKO-FRGZU")
                Q.Result.Columns.Remove("EKKO-MEMORY")
                Q.Result.Columns.Remove("EKKO-WAERS-0219")
                If Q.Result.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
                    Q.Result.Columns.Remove("EKPO-ZWERT")
                End If
                If Q.Result.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
                    Q.Result.Columns.Remove("EKKO-WAERS-0218")
                End If
                If Q.Result.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
                    Q.Result.Columns.Remove("EKKO-WAERS-0219")
                End If
                If Q.Result.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
                    Q.Result.Columns.Remove("EKKO-WAERS-0220")
                End If
                If Q.Result.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
                    Q.Result.Columns.Remove("EKKO-MEMORYTYPE")
                End If
                Dim R As DataRow
                For Each R In Q.Result.Rows
                    If IsNumeric(R("Item Number")) Then
                        R("Item Number") = CStr(Val(R("Item Number")))
                    End If
                    If IsNumeric(R("Material")) Then
                        R("Material") = CStr(CDbl(R("Material")))
                    End If
                    If IsNumeric(R("Vendor")) Then
                        R("Vendor") = CStr(CDbl(R("Vendor")))
                    End If
                    If IsNumeric(R("Doc Date")) Then
                        If Val(R("Doc Date")) <> 0 Then
                            R("Doc Date") = CStr(ConversionUtils.SAPDate2NetDate(R("Doc Date")))
                        Else
                            R("Doc Date") = Nothing
                        End If
                    Else
                        R("Doc Date") = Nothing
                    End If
                    R.AcceptChanges()
                    DNT.LoadDataRow(New Object() {R("Doc Number")}, LoadOption.OverwriteChanges)
                Next
                DNT.AcceptChanges()

                Q.Result.PrimaryKey = New DataColumn() {Q.Result.Columns("Doc Number"), Q.Result.Columns("Item Number")}
                Dim FR As DataRow = Nothing

                If IGRIR Then
                    Dim C As New DataColumn
                    C.DataType = System.Type.GetType("System.Double")
                    C.DefaultValue = 0
                    C.ColumnName = "GR Qty"
                    Q.Result.Columns.Add(C)
                    C = New DataColumn
                    C.DataType = System.Type.GetType("System.Double")
                    C.DefaultValue = 0
                    C.ColumnName = "IR Qty"
                    Q.Result.Columns.Add(C)
                    Dim GI As New GR_IR_Report(Con)
                    For Each R In DNT.Rows
                        GI.IncludeDocument(R(0))
                    Next
                    GI.Execute()
                    If GI.Success Then
                        If GI.Data.Rows.Count > 0 Then
                            For Each R In GI.Data.Rows
                                FR = Q.Result.Rows.Find(New Object() {R("PO"), R("Item")})
                                If Not FR Is Nothing Then
                                    FR("GR Qty") = R("GR")
                                    FR("IR Qty") = R("IR")
                                End If
                            Next
                        End If
                    Else
                        EM = GI.ErrMessage
                    End If
                End If

                If IDD Then
                    Dim DD As New DD_Report(Con)
                    Dim MD As Boolean
                    For Each R In DNT.Rows
                        DD.IncludeDocument(R(0))
                    Next
                    DD.Execute()
                    If DD.Success Then
                        If DD.Data.Rows.Count > 0 Then
                            Q.Result.Columns.Add("Delivery Date", System.Type.GetType("System.String"))
                            Q.Result.Columns.Add("Multi Deliv", System.Type.GetType("System.Boolean"))
                            For Each R In DD.Data.Rows
                                If DD.Data.Select("[Doc Number] = '" & R("Doc Number") & "' AND [Item Number] = '" & R("Item Number") & "'").Count > 1 Then
                                    MD = True
                                Else
                                    MD = False
                                End If
                                FR = Q.Result.Rows.Find(New Object() {R("Doc Number"), R("Item Number")})
                                If Not FR Is Nothing Then
                                    FR("Delivery Date") = R("Delivery Date")
                                    FR("Multi Deliv") = MD
                                End If
                            Next
                        End If
                    Else
                        EM = DD.ErrMessage
                    End If
                    DD = Nothing
                    GC.Collect()
                End If

                If IOA Then
                    Dim OAR As New OA_Report(Con)
                    For Each R In DNT.Rows
                        OAR.IncludeDocument(R(0))
                    Next
                    OAR.Execute()
                    If OAR.Success Then
                        If OAR.Data.Rows.Count > 0 Then
                            Q.Result.Columns.Add("OA", System.Type.GetType("System.String"))
                            For Each R In OAR.Data.Rows
                                For Each FR In Q.Result.Select("[Doc Number] = '" & R("Doc Number") & "'")
                                    FR("OA") = R("OA")
                                Next
                                FR.AcceptChanges()
                            Next
                        End If
                    End If
                End If

                If IYORF Or IPTRM Then
                    Dim YOR As New YO_Ref_Report(Con)
                    For Each R In DNT.Rows
                        YOR.IncludeDocument(R(0))
                    Next
                    YOR.Execute()
                    If YOR.Success Then
                        If YOR.Data.Rows.Count > 0 Then
                            Dim C As New DataColumn
                            Dim DR() As DataRow
                            If IYORF Then
                                C.DataType = System.Type.GetType("System.String")
                                C.ColumnName = "YReference"
                                Q.Result.Columns.Add(C)
                                C = New DataColumn
                                C.DataType = System.Type.GetType("System.String")
                                C.ColumnName = "OReference"
                                Q.Result.Columns.Add(C)
                            End If
                            If IPTRM Then
                                C = New DataColumn
                                C.DataType = System.Type.GetType("System.String")
                                C.ColumnName = "PTerm"
                                Q.Result.Columns.Add(C)
                            End If
                            For Each R In YOR.Data.Rows
                                DR = Q.Result.Select("[Doc Number] = '" & R("Doc Number") & "'")
                                If Not DR Is Nothing Then
                                    For Each FR In DR
                                        If IYORF Then
                                            FR("YReference") = R("YReference")
                                            FR("OReference") = R("OReference")
                                        End If
                                        If IPTRM Then
                                            FR("PTerm") = R("PTerm")
                                        End If
                                    Next
                                    FR.AcceptChanges()
                                End If
                            Next
                        End If
                    Else
                        EM = YOR.ErrMessage
                    End If
                End If

                If RLV <> RepairsLevels.DoNotProcess Then
                    Dim C As New DataColumn("Repair", System.Type.GetType("System.Boolean"))
                    C.DefaultValue = False
                    Q.Result.Columns.Add(C)
                    Dim RF As New Repairs_Filter(Con)
                    For Each R In DNT.Rows
                        RF.IncludeDocument(R(0))
                    Next
                    RF.Execute()
                    If RF.Success Then
                        If RF.Data.Rows.Count > 0 Then
                            Dim SR() As DataRow
                            Dim RR As DataRow
                            For Each RR In RF.Data.Rows
                                SR = Q.Result.Select("[Doc Number] = '" & RR(0) & "'")
                                For Each R In SR
                                    If RLV = RepairsLevels.ExcludeRepairs Then
                                        R.Delete()
                                    Else
                                        R("Repair") = True
                                    End If
                                    R.AcceptChanges()
                                Next
                            Next
                        End If
                    End If
                    If RLV = RepairsLevels.OnlyRepairs Then
                        For Each R In Q.Result.Rows
                            If R("Repair") = False Then R.Delete()
                        Next
                        Q.Result.AcceptChanges()
                    End If
                    RF = Nothing
                    GC.Collect()
                End If
                If Q.Result.Columns.IndexOf("Multi Deliv") > 0 Then
                    Q.Result.Columns("Multi Deliv").SetOrdinal(Q.Result.Columns.Count - 1)
                End If
            End If
            SF = True
        Catch ex As Exception
            EM = ex.Message
        End Try

    End Sub

End Class

Public NotInheritable Class OpenOrders_Report : Inherits OO_Query

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal ACon As Object)

        MyBase.New(ACon)

    End Sub

    Public Overrides Sub Execute()

        MyBase.SelectCriteria("WE103")
        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class OpenGR105_Report : Inherits OO_Query


    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal ACon As Object)

        MyBase.New(ACon)

    End Sub

    Public Overrides Sub Execute()

        MyBase.SelectCriteria("WE105")
        MyBase.Execute()

    End Sub

End Class

Public NotInheritable Class POs_Report : Inherits OO_Query

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        MyBase.New(Box, User, App)

    End Sub

    Sub New(ByVal ACon As Object)

        MyBase.New(ACon)

    End Sub

End Class

Public NotInheritable Class OpenReqs_Report

    Private Con As R3Connection = Nothing
    Private Q As Query = Nothing
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private A(,) As String = Nothing
    Private D() As String = Nothing
    Private DNT As DataTable = Nothing
    Private RLV As Byte = 0
    Private RDUT As Date

    Public ReadOnly Property DataArray()

        Get
            If Not Q Is Nothing Then
                If A Is Nothing Then
                    Dim R As DataRow
                    Dim I As Integer
                    Dim X As Integer
                    ReDim A(Q.Result.Rows.Count, Q.Result.Columns.Count)
                    For I = 1 To Q.Result.Columns.Count
                        A(0, I) = Q.Result.Columns(I - 1).ColumnName
                    Next
                    I = 1
                    For Each R In Q.Result.Rows
                        For X = 1 To Q.Result.Columns.Count
                            A(I, X) = R(X - 1).ToString
                        Next
                        I += 1
                    Next
                End If
            End If
            DataArray = A
        End Get

    End Property

    Public ReadOnly Property DocsArray()

        Get
            If Not DNT Is Nothing Then
                If D Is Nothing Then
                    ReDim D(DNT.Rows.Count)
                    Dim I As Integer = 1
                    Dim R As DataRow
                    For Each R In DNT.Rows
                        D(I) = R(0)
                        I += 1
                    Next
                End If
            End If
            DocsArray = D
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            If Not Q Is Nothing Then
                Data = Q.Result
            Else
                Data = Nothing
            End If
        End Get

    End Property

    Public ReadOnly Property Documents() As DataTable

        Get
            Documents = DNT
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public Property RepairsLevel() As Byte

        Get
            RepairsLevel = RLV
        End Get
        Set(ByVal value As Byte)
            RLV = value
        End Set

    End Property

    Public Property ReleaseDateUpTo() As Date

        Get
            ReleaseDateUpTo = RDUT
        End Get
        Set(ByVal value As Date)
            RDUT = value
            IncludeParamUpTo("BA_FRGDT", ConversionUtils.NetDate2SAPDate(RDUT))
        End Set

    End Property

    Private Sub IncludeParam(ByVal Param As String, ByVal Value As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Include, RangeOption.Equals, Value)
        End If

    End Sub

    Private Sub IncludeParamFromTo(ByVal Param As String, ByVal LowValue As String, ByVal HighValue As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Include, RangeOption.Between, LowValue, HighValue)
        End If

    End Sub

    Private Sub IncludeParamUpTo(ByVal Param As String, ByVal HighValue As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Include, RangeOption.LessThanOrEqualTo, HighValue)
        End If

    End Sub

    Private Sub ExcludeParam(ByVal Param As String, ByVal Value As String)

        If Not Q Is Nothing Then
            Q.SelectionParameters(Param).Ranges.Add(Sign.Exclude, RangeOption.Equals, Value)
        End If

    End Sub

    Sub New(ByVal Box As String, ByVal User As String, ByVal App As String)

        Dim C As New SAPConnector
        Con = C.GetSAPConnection(Box, User, App)
        If Not Con Is Nothing Then
            InitBAPI()
        Else
            EM = C.Status
        End If

    End Sub

    Sub New(ByVal Connection As Object)

        If Connection.Ping Then
            Con = Connection
            InitBAPI()
        Else
            EM = "Connection already closed"
        End If

    End Sub

    Private Sub InitBAPI()

        Try
            Q = Con.CreateQuery(WorkSpace.GlobalArea, "/SAPQUERY/ME", "MEBANF")
            Q.Variant = "SAP&MEBANF"
            IncludeParam("P_QERLBA", "X")
            IncludeParam("P_QFREIG", "X")
            DNT = New DataTable("Req Numbers")
            DNT.Columns.Add("Req Number", System.Type.GetType("System.String"))
            Dim Keys(0) As DataColumn
            Keys(0) = DNT.Columns("Req Number")
            DNT.PrimaryKey = Keys
            SF = True
        Catch ex As Exception
            Q = Nothing
            EM = ex.Message
        End Try

    End Sub

    Public Sub EndSession()

        Con.Close()

    End Sub

    Public Sub IncludeReleaseDateFromTo(ByVal DateFrom As Date, ByVal DateTo As Date)

        IncludeParamFromTo("BA_FRGDT", ConversionUtils.NetDate2SAPDate(DateFrom), ConversionUtils.NetDate2SAPDate(DateTo))

    End Sub

    Public Sub IncludePlant(ByVal Plant As String)

        IncludeParam("SP$00036", Plant)

    End Sub

    Public Sub ExcludePlant(ByVal Plant As String)

        ExcludeParam("SP$00036", Plant)

    End Sub

    Public Sub IncludeReq(ByVal Number As String)

        IncludeParam("SP$00026", Number)

    End Sub

    Public Sub ExcludeReq(ByVal Number As String)

        ExcludeParam("SP$00026", Number)

    End Sub

    Public Sub IncludePurchGroup(ByVal Code As String)

        IncludeParam("SP$00031", Code)

    End Sub

    Public Sub ExcludePurchGroup(ByVal Code As String)

        ExcludeParam("SP$00031", Code)

    End Sub

    Public Sub IncludePurchOrg(ByVal Code As String)

        IncludeParam("SP$00040", Code)

    End Sub

    Public Sub ExcludePurchOrg(ByVal Code As String)

        ExcludeParam("SP$00040", Code)

    End Sub

    Public Sub IncludeMatGroup(ByVal Code As String)

        IncludeParam("SP$00033", Code)

    End Sub

    Public Sub ExcludeMatGroup(ByVal Code As String)

        ExcludeParam("SP$00033", Code)

    End Sub

    Public Sub IncludeMaterial(ByVal Number As String)

        IncludeParam("SP$00034", Number)

    End Sub

    Public Sub ExcludeMaterial(ByVal Number As String)

        ExcludeParam("SP$00034", Number)

    End Sub

    Public Sub IncludeMaterialFromTo(ByVal MatFrom As String, ByVal MatTo As String)

        IncludeParamFromTo("SP$00034", MatFrom, MatTo)

    End Sub

    Public Sub Execute()

        If Q Is Nothing Then Exit Sub

        Try
            IncludeParam("P_QCOUNT", "0")
            Q.Execute()
            If Q.Result.Rows.Count > 0 Then
                Q.Result.Columns("EBAN-BANFN").ColumnName = "Req Number"
                Q.Result.Columns("EBAN-BNFPO").ColumnName = "Item Number"
                Q.Result.Columns("EBAN-MATKL").ColumnName = "Mat Group"
                Q.Result.Columns("EBAN-MATNR").ColumnName = "Material"
                Q.Result.Columns("EBAN-TXZ01").ColumnName = "Short Text"
                Q.Result.Columns("EBAN-EKGRP").ColumnName = "Purch Grp"
                Q.Result.Columns("EBAN-WERKS").ColumnName = "Plant"
                Q.Result.Columns("EBAN-BEDNR").ColumnName = "Tracking Field"
                Q.Result.Columns("EBAN-BSART").ColumnName = "Doc Type"
                Q.Result.Columns("EBAN-PSTYP").ColumnName = "Item Category"
                Q.Result.Columns("EBAN-BSMNG").ColumnName = "PO Quantity"
                Q.Result.Columns("EBAN-FLIEF").ColumnName = "Fixed Vendor"
                Q.Result.Columns("FLIEFTXT").ColumnName = "Fix Vendor Name"
                Q.Result.Columns("EBAN-LIFNR").ColumnName = "Desired Vendor"
                Q.Result.Columns("WLIEFTXT").ColumnName = "Des Vendor Name"
                Q.Result.Columns("EBAN-EKORG").ColumnName = "Purch Org"
                Q.Result.Columns("EBAN-KONNR").ColumnName = "Outline Agreement"
                Q.Result.Columns("EBAN-KTPNR").ColumnName = "Agreement Item"
                Q.Result.Columns("EBAN-AFNAM").ColumnName = "Requisitioner"
                Q.Result.Columns("EBAN-BADAT").ColumnName = "Req Date"
                Q.Result.Columns("EBAN-BEDAT").ColumnName = "PO Date"
                Q.Result.Columns("EBAN-DISPO").ColumnName = "MRP Controller"
                Q.Result.Columns("EBAN-LGORT").ColumnName = "Storage"
                Q.Result.Columns("EBAN-LOEKZ").ColumnName = "Del Indicator"
                Q.Result.Columns("EBAN-EBELN").ColumnName = "PO Number"
                Q.Result.Columns("EBAN-EBELP").ColumnName = "PO Item"
                Q.Result.Columns("EBAN-ERNAM").ColumnName = "Created By"
                Q.Result.Columns("EBAN-LFDAT").ColumnName = "Delivery Date"
                Q.Result.Columns("EBAN-MEINS").ColumnName = "UOM"
                Q.Result.Columns("EBAN-MENGE").ColumnName = "Quantity"
                Q.Result.Columns("EBAN-FRGDT").ColumnName = "Release Date"
                Q.Result.Columns.Remove("EBAN-EMATN")
                Q.Result.Columns.Remove("EBAN-BSTYP")
                Q.Result.Columns.Remove("EBAN-KNTTP")
                Q.Result.Columns.Remove("EBAN-RESWK")
                Q.Result.Columns.Remove("EBAN-STATU")
                Q.Result.Columns.Remove("EBAN-MEINS-0307")
                Q.Result.Columns.Remove("EBAN-BSAKZ")
                Q.Result.Columns.Remove("EBAN-EBAKZ")
                Q.Result.Columns.Remove("EBAN-CHARG")
                Q.Result.Columns.Remove("EBAN-INFNR")
                Q.Result.Columns.Remove("EBAN-MEINS-0603")
                Q.Result.Columns.Remove("EBAN-SERNR")
                Q.Result.Columns.Remove("EBAN-SOBKZ")
                Q.Result.Columns.Remove("EBAN-FRGRL")
                Q.Result.Columns.Remove("EBAN-FRGGR")
                Q.Result.Columns.Remove("EBAN-FRGKZ")
                Q.Result.Columns.Remove("EBAN-FRGST")
                Q.Result.Columns.Remove("EBAN-FRGZU")
                Q.Result.Columns.Remove("EBAN-GSFRG")
                Q.Result.Columns.Remove("EBAN-ZUGBA")
                If Not Q.Result.Columns.IndexOf("EBAN-MEMORY") < 0 Then
                    Q.Result.Columns.Remove("EBAN-MEMORY")
                    Q.Result.Columns.Remove("EBAN-MEMORYTYPE")
                End If
                Dim R As DataRow
                For Each R In Q.Result.Rows
                    R("Req Number") = CStr(CDbl(R("Req Number")))
                    R("Item Number") = CStr(Val(R("Item Number")))
                    If IsNumeric(R("Agreement Item")) Then
                        If Val(R("Agreement Item")) > 0 Then
                            R("Agreement Item") = CStr(Val(R("Agreement Item")))
                        Else
                            R("Agreement Item") = ""
                        End If
                    End If
                    If IsNumeric(R("PO Item")) Then
                        If Val(R("PO Item")) > 0 Then
                            R("PO Item") = CStr(Val(R("PO Item")))
                        Else
                            R("PO Item") = ""
                        End If
                    End If
                    If IsNumeric(R("Material")) Then
                        R("Material") = CStr(CDbl(R("Material")))
                    End If
                    If IsNumeric(R("Fixed Vendor")) Then
                        R("Fixed Vendor") = CStr(CDbl(R("Fixed Vendor")))
                    End If
                    If IsNumeric(R("Desired Vendor")) Then
                        R("Desired Vendor") = CStr(CDbl(R("Desired Vendor")))
                    End If
                    If IsNumeric(R("Req Date")) Then
                        If Val(R("Req Date")) <> 0 Then
                            R("Req Date") = CStr(ConversionUtils.SAPDate2NetDate(R("Req Date")))
                        Else
                            R("Req Date") = Nothing
                        End If
                    Else
                        R("Req Date") = Nothing
                    End If
                    If IsNumeric(R("PO Date")) Then
                        If Val(R("PO Date")) <> 0 Then
                            R("PO Date") = CStr(ConversionUtils.SAPDate2NetDate(R("PO Date")))
                        Else
                            R("PO Date") = Nothing
                        End If
                    Else
                        R("PO Date") = Nothing
                    End If
                    If IsNumeric(R("Delivery Date")) Then
                        If Val(R("Delivery Date")) <> 0 Then
                            R("Delivery Date") = CStr(ConversionUtils.SAPDate2NetDate(R("Delivery Date")))
                        Else
                            R("Delivery Date") = Nothing
                        End If
                    Else
                        R("Delivery Date") = Nothing
                    End If
                    If IsNumeric(R("Release Date")) Then
                        If Val(R("Release Date")) <> 0 Then
                            R("Release Date") = CStr(ConversionUtils.SAPDate2NetDate(R("Release Date")))
                        Else
                            R("Release Date") = Nothing
                        End If
                    Else
                        R("Release Date") = Nothing
                    End If
                    R.AcceptChanges()
                    DNT.LoadDataRow(New Object() {R("Req Number")}, LoadOption.OverwriteChanges)
                Next
                DNT.AcceptChanges()

                If RLV <> RepairsLevels.DoNotProcess Then
                    Dim C As New DataColumn("Repair", System.Type.GetType("System.Boolean"))
                    C.DefaultValue = False
                    Q.Result.Columns.Add(C)
                    Dim RI() As DataRow = Q.Result.Select("[Item Category] = '3'")
                    Dim RF() As DataRow
                    Dim RR As DataRow
                    For Each R In RI
                        RF = Q.Result.Select("[Req Number] = '" & R("Req Number") & "'")
                        For Each RR In RF
                            RR("Repair") = True
                            RR.AcceptChanges()
                        Next
                    Next
                    If RLV = RepairsLevels.ExcludeRepairs Then
                        For Each R In Q.Result.Rows
                            If R("Repair") = True Then R.Delete()
                        Next
                        Q.Result.AcceptChanges()
                    End If
                    If RLV = RepairsLevels.OnlyRepairs Then
                        For Each R In Q.Result.Rows
                            If R("Repair") = False Then R.Delete()
                        Next
                        Q.Result.AcceptChanges()
                    End If
                End If
            End If

            SF = True

        Catch ex As Exception
            EM = ex.Message
        End Try

    End Sub

End Class

#End Region

#End Region

#Region "Private/Support Classes"

Friend Structure SAPText

    Dim Format As String
    Dim Text As String

End Structure

Friend Enum RTableOperator

    Include = 0
    Exclude = 1
    IncludeFrom = 2
    IncludeTo = 3
    ExcludeFrom = 4
    ExcludeTo = 5

End Enum

Friend Enum ContractType

    Unknown = 0
    OutlineAgreement = 1
    SchedulingAgreement = 2

End Enum

Public Class RTable 'Friend NotInheritable

    Private T As DataTable = Nothing
    Private WithEvents R As ReadTable
    Private EM As String = Nothing
    Private SF As Boolean = False
    Private KC() As Integer = Nothing
    Private Fields(,) As String = Nothing
    Private CIndex As New DataTable
    Private Criteria As New DataTable
    Private CF As Boolean = False

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String
        Get
            ErrMessage = EM
        End Get
    End Property

    Public Property MaxHits() As Integer

        Get
            MaxHits = R.RowCount
        End Get
        Set(ByVal value As Integer)
            R.RowCount = value
        End Set

    End Property

    Public Property TableName() As String

        Get
            TableName = R.TableName
        End Get
        Set(ByVal value As String)
            R.TableName = value
            R.RowCount = 10000000
        End Set

    End Property

    Public ReadOnly Property Result() As DataTable

        Get
            Result = T
        End Get

    End Property

    Public Sub New(ByVal Con As Object)

        R = New ReadTable(Con)
        R.PackageSize = 10000
        R.RaiseIncomingPackageEvent = True
        CIndex.Columns.Add("FName", System.Type.GetType("System.String"))
        CIndex.Columns.Add("Op", System.Type.GetType("System.Int32"))
        CIndex.PrimaryKey = New DataColumn() {CIndex.Columns("FName"), CIndex.Columns("Op")}
        Criteria.Columns.Add("FName", System.Type.GetType("System.String"))
        Criteria.Columns.Add("Op", System.Type.GetType("System.Int32"))
        Criteria.Columns.Add("Value", System.Type.GetType("System.String"))

    End Sub

    Private Sub OnIncomingPackage(ByVal Sender As ReadTable, ByVal PackageResult As DataTable) Handles R.IncomingPackage

        If T Is Nothing Then
            T = PackageResult.Clone
            If Not KC Is Nothing Then
                Dim PK(KC.GetUpperBound(0)) As DataColumn
                Dim I As Integer
                For I = 0 To KC.GetUpperBound(0)
                    PK(I) = T.Columns(KC(I))
                Next
                T.PrimaryKey = PK
            End If
            T.TableName = Sender.TableName
        End If
        Dim R As New DataTableReader(PackageResult)
        T.Load(R, LoadOption.PreserveChanges)

    End Sub

    Public Sub AddCriteria(ByVal Criteria As String)

        R.AddCriteria(Criteria)

    End Sub

    Public Sub Run()

        Try
            If CIndex.Rows.Count > 0 Then BuildCriteria()
            SF = True
            R.Run()
            If Not Fields Is Nothing Then
                For I As Integer = 0 To Fields.GetUpperBound(1)
                    T.Columns(Fields(0, I)).ColumnName = Fields(1, I)
                Next
            End If
        Catch ex As Exception
            SF = False
            EM = ex.Message
        End Try

    End Sub

    Public Sub AddField(ByVal FName As String)

        R.AddField(FName)

    End Sub

    Public Sub AddField(ByVal FName As String, ByVal FLabel As String)

        If Fields Is Nothing Then
            ReDim Fields(1, 0)
        Else
            ReDim Preserve Fields(1, Fields.GetUpperBound(1) + 1)
        End If
        Fields(0, Fields.GetUpperBound(1)) = FName
        Fields(1, Fields.GetUpperBound(1)) = FLabel
        R.AddField(FName)

    End Sub

    Public Sub IncludeAllFields()

        R.RetrieveAllFieldsOfTable()

    End Sub

    Public Sub AddKeyColumn(ByVal Index As Integer)

        If KC Is Nothing Then
            ReDim KC(0)
        Else
            ReDim Preserve KC(KC.GetUpperBound(0) + 1)
        End If
        KC(KC.GetUpperBound(0)) = Index

    End Sub

    Public Sub ParamInclude(ByVal FName As String, ByVal Value As String)

        CIndex.LoadDataRow(New Object() {FName, RTableOperator.Include}, LoadOption.OverwriteChanges)
        Criteria.LoadDataRow(New Object() {FName, RTableOperator.Include, Value}, True)

    End Sub

    Public Sub ParamExclude(ByVal FName As String, ByVal Value As String)

        CIndex.LoadDataRow(New Object() {FName, RTableOperator.Exclude}, LoadOption.OverwriteChanges)
        Criteria.LoadDataRow(New Object() {FName, RTableOperator.Exclude, Value}, True)

    End Sub

    Public Sub ParamIncludeFromTo(ByVal FName As String, ByVal VFrom As String, ByVal VTo As String)

        CIndex.LoadDataRow(New Object() {FName, RTableOperator.IncludeFrom}, LoadOption.OverwriteChanges)
        Criteria.LoadDataRow(New Object() {FName, RTableOperator.IncludeFrom, VFrom}, True)
        Criteria.LoadDataRow(New Object() {FName, RTableOperator.IncludeTo, VTo}, True)

    End Sub

    'Public Sub ParamExcludeFromTo(ByVal FName As String, ByVal VFrom As String, ByVal VTo As String)

    '    CIndex.LoadDataRow(New Object() {FName, RTableOperator.ExcludeFrom}, LoadOption.OverwriteChanges)
    '    Criteria.LoadDataRow(New Object() {FName, RTableOperator.ExcludeFrom, VFrom}, True)
    '    Criteria.LoadDataRow(New Object() {FName, RTableOperator.ExcludeTo, VTo}, True)

    'End Sub

    Public Sub ColumnToDateStr(ByVal ColumnName As String)

        Dim DR As DataRow
        For Each DR In T.Rows
            If IsNumeric(DR(ColumnName)) And DR(ColumnName) <> "00000000" Then
                DR(ColumnName) = CStr(ConversionUtils.SAPDate2NetDate(DR(ColumnName)))
            End If
        Next

    End Sub

    Public Sub ColumnToDoubleStr(ByVal ColumnName As String)

        Dim DR As DataRow
        For Each DR In T.Rows
            If IsNumeric(DR(ColumnName)) Then
                DR(ColumnName) = CStr(CDbl(DR(ColumnName)))
            End If
        Next

    End Sub

    Public Sub ColumnToIntStr(ByVal ColumnName As String)

        Dim DR As DataRow
        For Each DR In T.Rows
            If IsNumeric(DR(ColumnName)) Then
                DR(ColumnName) = CStr(CInt(DR(ColumnName)))
            End If
        Next

    End Sub

    Private Sub BuildCriteria()

        Dim TR() As DataRow
        Dim FR() As DataRow
        For Each DR As DataRow In CIndex.Rows
            Select Case DR("Op")
                Case RTableOperator.Include, RTableOperator.Exclude
                    BuildSingleParamCriteria(Criteria.Select("FName = '" & DR("FName") & "' AND Op = " & DR("Op")), DR("Op"))
                Case RTableOperator.IncludeFrom
                    FR = Criteria.Select("FName = '" & DR("FName") & "' AND Op = " & RTableOperator.IncludeFrom)
                    TR = Criteria.Select("FName = '" & DR("FName") & "' AND Op = " & RTableOperator.IncludeTo)
                    If TR.Length > 0 Then
                        BuildFromToParamCriteria(DR("FName"), DR("Op"), FR(0)("Value"), TR(0)("Value"))
                    End If
            End Select
        Next

    End Sub

    Private Sub BuildSingleParamCriteria(ByVal CDR() As DataRow, ByVal Op As Integer)

        Dim DR As DataRow
        Dim I As Integer = 0
        Dim S As String
        Dim OpS As String = Nothing

        Select Case Op
            Case RTableOperator.Include
                OpS = " = "
            Case RTableOperator.Exclude
                OpS = " <> "
        End Select

        For Each DR In CDR
            If I = 0 Then
                If CF Then
                    S = " AND (" & DR("FName") & OpS & "'" & DR("Value") & "'"
                Else
                    S = "(" & DR("FName") & OpS & "'" & DR("Value") & "'"
                End If
            Else
                S = DR("FName") & OpS & "'" & DR("Value") & "'"
            End If
            If I < CDR.GetUpperBound(0) Then
                If Op = RTableOperator.Include Then
                    S = S & " OR "
                Else
                    S = S & " AND "
                End If
            Else
                S = S & ")"
            End If
            If S.Contains("*") Then
                S = S.Replace("*", "%")
            End If
            If S.Contains("%") AndAlso S.Contains("=") Then
                S = S.Replace("=", "LIKE")
            End If
            If S.Contains("%") AndAlso S.Contains("<>") Then
                S = S.Replace("=", "NOT LIKE")
            End If
            AddCriteria(S)
            CF = True
            I += 1
        Next

    End Sub

    Private Sub BuildFromToParamCriteria(ByVal FName As String, ByVal Op As Integer, ByVal VFrom As String, ByVal VTo As String)

        Dim S As String
        If CF Then
            S = " AND (" & FName & " >= '" & VFrom & "' AND " & FName & " <= '" & VTo & "')"
        Else
            S = "(" & FName & " >= '" & VFrom & "' AND " & FName & " <= '" & VTo & "')"
        End If
        AddCriteria(S)
        CF = True

    End Sub

End Class

Public MustInherit Class LINQ_Support

    Public Function LinQToDataTable(Of T)(ByVal source As IEnumerable(Of T)) As DataTable

        Return New ObjectShredder(Of T)().Shred(source, Nothing, Nothing)

    End Function

End Class

Friend NotInheritable Class Repairs_Filter

    Private D As DataTable = Nothing
    Private T As RTable
    Private EM As String = Nothing
    Private SF As Boolean = False

    Private IDN() As String = Nothing

    Public Sub IncludeDocument(ByVal Number As String)

        If IDN Is Nothing Then
            ReDim IDN(0)
        Else
            ReDim Preserve IDN(UBound(IDN) + 1)
        End If

        IDN(UBound(IDN)) = Number

    End Sub

    Public Sub New(ByVal Con As R3Connection)

        If Not Con Is Nothing Then
            T = New RTable(Con)
            T.TableName = "EKPO"
            D = New DataTable("Repair Docs")
            D.Columns.Add("Doc Number", System.Type.GetType("System.String"))
            Dim Keys(0) As DataColumn
            Keys(0) = D.Columns("Doc Number")
        End If

    End Sub

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = D
        End Get

    End Property

    Public Sub Execute()

        Dim S As String
        Dim I As Integer
        Dim F As Boolean = False

        If Not IDN Is Nothing Then
            I = 0
            For Each S In IDN
                If I = 0 Then
                    S = "(EBELN = '" & S & "'"
                Else
                    S = "EBELN = '" & S & "'"
                End If
                If I < IDN.GetUpperBound(0) Then
                    S = S & " OR "
                Else
                    S = S & ")"
                End If
                T.AddCriteria(S)
                I += 1
            Next
            F = True
        End If

        If F Then
            S = " AND PSTYP = '3'"
        Else
            S = "PSTYP = '3'"
        End If
        T.AddCriteria(S)

        T.AddField("EBELN") 'Doc Number

        T.Run()
        If Not T.Success Then
            EM = T.ErrMessage
            Exit Sub
        Else
            SF = True
        End If
        If T.Result.Rows.Count > 0 Then
            Dim R As DataRow
            For Each R In T.Result.Rows
                D.LoadDataRow(New Object() {R("EBELN")}, LoadOption.OverwriteChanges)
            Next
            D.AcceptChanges()
            T = Nothing
            GC.Collect()
        End If

    End Sub

End Class

Friend NotInheritable Class DD_Report

    Private T As RTable
    Private EM As String = Nothing
    Private SF As Boolean = False

    Private IDN() As String = Nothing

    Public Sub IncludeDocument(ByVal Number As String)

        If IDN Is Nothing Then
            ReDim IDN(0)
        Else
            ReDim Preserve IDN(UBound(IDN) + 1)
        End If

        IDN(UBound(IDN)) = Number

    End Sub

    Public Sub New(ByVal Con As R3Connection)

        If Not Con Is Nothing Then
            T = New RTable(Con)
            T.TableName = "EKET"
        End If

    End Sub

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = T.Result
        End Get

    End Property

    Public Sub Execute()

        Dim S As String
        Dim I As Integer

        If Not IDN Is Nothing Then
            I = 0
            For Each S In IDN
                If I = 0 Then
                    S = "(EBELN = '" & S & "'"
                Else
                    S = "EBELN = '" & S & "'"
                End If
                If I < IDN.GetUpperBound(0) Then
                    S = S & " OR "
                Else
                    S = S & ")"
                End If
                T.AddCriteria(S)
                I += 1
            Next
        End If

        T.AddField("EBELN") 'Doc Number
        T.AddField("EBELP") 'Item Number
        T.AddField("EINDT") 'Delivery Date

        T.Run()
        If Not T.Success Then
            EM = T.ErrMessage
            Exit Sub
        Else
            SF = True
        End If
        If T.Result.Rows.Count > 0 Then
            T.Result.Columns("EBELN").ColumnName = "Doc Number"
            T.Result.Columns("EBELP").ColumnName = "Item Number"
            T.Result.Columns("EINDT").ColumnName = "Delivery Date"
            Dim R As DataRow
            For Each R In T.Result.Rows
                If IsNumeric(R("Item Number")) Then
                    R("Item Number") = CStr(Val(R("Item Number")))
                End If
                If IsNumeric(R("Delivery Date")) Then
                    If Val(R("Delivery Date")) > 0 Then
                        R("Delivery Date") = ConversionUtils.SAPDate2NetDate(R("Delivery Date")).ToShortDateString
                    Else
                        R("Delivery Date") = Nothing
                    End If
                Else
                    R("Delivery Date") = Nothing
                End If
                R.AcceptChanges()
            Next
        End If

    End Sub

End Class

Friend NotInheritable Class YO_Ref_Report

    Private T As RTable
    Private EM As String = Nothing
    Private SF As Boolean = False

    Private IDN() As String = Nothing

    Public Sub IncludeDocument(ByVal Number As String)

        If IDN Is Nothing Then
            ReDim IDN(0)
        Else
            ReDim Preserve IDN(UBound(IDN) + 1)
        End If

        IDN(UBound(IDN)) = Number

    End Sub

    Public Sub New(ByVal Con As R3Connection)

        If Not Con Is Nothing Then
            T = New RTable(Con)
            T.TableName = "EKKO"
        End If

    End Sub

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property ErrMessage() As String

        Get
            ErrMessage = EM
        End Get

    End Property

    Public ReadOnly Property Data() As DataTable

        Get
            Data = T.Result
        End Get

    End Property

    Public Sub Execute()

        Dim S As String
        Dim I As Integer

        If Not IDN Is Nothing Then
            I = 0
            For Each S In IDN
                If I = 0 Then
                    S = "(EBELN = '" & S & "'"
                Else
                    S = "EBELN = '" & S & "'"
                End If
                If I < IDN.GetUpperBound(0) Then
                    S = S & " OR "
                Else
                    S = S & ")"
                End If
                T.AddCriteria(S)
                I += 1
            Next
        End If

        T.AddField("EBELN") 'Doc Number
        T.AddField("IHREZ") 'YReference
        T.AddField("UNSEZ") 'OReference
        T.AddField("ZTERM") 'PTerms

        T.Run()
        If Not T.Success Then
            EM = T.ErrMessage
            Exit Sub
        Else
            SF = True
        End If
        If T.Result.Rows.Count > 0 Then
            T.Result.Columns("EBELN").ColumnName = "Doc Number"
            T.Result.Columns("IHREZ").ColumnName = "YReference"
            T.Result.Columns("UNSEZ").ColumnName = "OReference"
            T.Result.Columns("ZTERM").ColumnName = "PTerm"
        End If

    End Sub

End Class

Friend NotInheritable Class Simple3Des

    Private Key As String = ")@*$&^!\[]{/:';<~`+=zo2916qw-"
    Private TripleDes As New TripleDESCryptoServiceProvider

    Sub New()

        TripleDes.Key = TruncateHash(Key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)

    End Sub

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ReDim Preserve hash(length - 1)
        Return hash

    End Function

    Public Function EncryptData(ByVal plaintext As String) As String

        If plaintext Is Nothing OrElse plaintext = "" Then Return ""

        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
        Dim ms As New System.IO.MemoryStream
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        Return Convert.ToBase64String(ms.ToArray)

    End Function

    Public Function DecryptData(ByVal encryptedtext As String) As String

        If encryptedtext Is Nothing OrElse encryptedtext = "" Then Return ""

        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)
        Dim ms As New System.IO.MemoryStream
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)

    End Function

End Class

Friend Class ObjectShredder(Of T)
    ' Fields
    Private _fi As FieldInfo()
    Private _ordinalMap As Dictionary(Of String, Integer)
    Private _pi As PropertyInfo()
    Private _type As Type

    ' Constructor 
    Public Sub New()
        Me._type = GetType(T)
        Me._fi = Me._type.GetFields
        Me._pi = Me._type.GetProperties
        Me._ordinalMap = New Dictionary(Of String, Integer)
    End Sub

    Public Function ShredObject(ByVal table As DataTable, ByVal instance As T) As Object()
        Dim fi As FieldInfo() = Me._fi
        Dim pi As PropertyInfo() = Me._pi
        If (Not instance.GetType Is GetType(T)) Then
            ' If the instance is derived from T, extend the table schema
            ' and get the properties and fields.
            Me.ExtendTable(table, instance.GetType)
            fi = instance.GetType.GetFields
            pi = instance.GetType.GetProperties
        End If

        ' Add the property and field values of the instance to an array.
        Dim values As Object() = New Object(table.Columns.Count - 1) {}
        Dim f As FieldInfo
        For Each f In fi
            values(Me._ordinalMap.Item(f.Name)) = f.GetValue(instance)
        Next
        Dim p As PropertyInfo
        For Each p In pi
            values(Me._ordinalMap.Item(p.Name)) = p.GetValue(instance, Nothing)
        Next

        ' Return the property and field values of the instance.
        Return values
    End Function


    ' Summary:           Loads a DataTable from a sequence of objects.
    ' source parameter:  The sequence of objects to load into the DataTable.</param>
    ' table parameter:   The input table. The schema of the table must match that 
    '                    the type T.  If the table is null, a new table is created  
    '                    with a schema created from the public properties and fields 
    '                    of the type T.
    ' options parameter: Specifies how values from the source sequence will be applied to 
    '                    existing rows in the table.
    ' Returns:           A DataTable created from the source sequence.

    Public Function Shred(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable

        ' Load the table from the scalar sequence if T is a primitive type.
        If GetType(T).IsPrimitive Then
            Return Me.ShredPrimitive(source, table, options)
        End If

        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If

        ' Initialize the ordinal map and extend the table schema based on type T.
        table = Me.ExtendTable(table, GetType(T))

        ' Enumerate the source sequence and load the object values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Do While e.MoveNext
                If options.HasValue Then
                    table.LoadDataRow(Me.ShredObject(table, e.Current), options.Value)
                Else
                    table.LoadDataRow(Me.ShredObject(table, e.Current), True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function


    Public Function ShredPrimitive(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable
        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If
        If Not table.Columns.Contains("Value") Then
            table.Columns.Add("Value", GetType(T))
        End If

        ' Enumerate the source sequence and load the scalar values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Dim values As Object() = New Object(table.Columns.Count - 1) {}
            Do While e.MoveNext
                values(table.Columns.Item("Value").Ordinal) = e.Current
                If options.HasValue Then
                    table.LoadDataRow(values, options.Value)
                Else
                    table.LoadDataRow(values, True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function

    Public Function ExtendTable(ByVal table As DataTable, ByVal type As Type) As DataTable
        ' Extend the table schema if the input table was null or if the value 
        ' in the sequence is derived from type T.
        Dim f As FieldInfo
        Dim p As PropertyInfo

        For Each f In type.GetFields
            If Not Me._ordinalMap.ContainsKey(f.Name) Then
                Dim dc As DataColumn

                ' Add the field as a column in the table if it doesn't exist
                ' already.
                dc = IIf(table.Columns.Contains(f.Name), table.Columns.Item(f.Name), table.Columns.Add(f.Name, f.FieldType))

                ' Add the field to the ordinal map.
                Me._ordinalMap.Add(f.Name, dc.Ordinal)
            End If

        Next

        For Each p In type.GetProperties
            If Not Me._ordinalMap.ContainsKey(p.Name) Then
                ' Add the property as a column in the table if it doesn't exist
                ' already.
                Dim dc As DataColumn
                If table.Columns.Contains(p.Name) Then
                    dc = table.Columns.Item(p.Name)
                Else
                    dc = table.Columns.Add(p.Name, p.PropertyType)
                End If
                ' Add the property to the ordinal map.
                Me._ordinalMap.Add(p.Name, dc.Ordinal)
            End If
        Next

        ' Return the table.
        Return table
    End Function

End Class

#End Region