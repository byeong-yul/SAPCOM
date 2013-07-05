Imports System.Windows.Forms

Public Class Connection_Manager

    Private AC As String = Nothing
    Private UC As String = Nothing
    Private SC As New SAPConnector

    Sub New(ByVal UserCode As String, ByVal AppCode As String)

        Me.InitializeComponent()
        AC = AppCode
        UC = UserCode

    End Sub

    Private Sub SetUp()

        Dim D As ConnectionData = SC.GetConnectionData(cb_SAP.SelectedItem, UC, AC)
        txt_Login.Text = D.Login
        txt_Password.Text = D.Password
        CheckBox2.Checked = D.SSO

    End Sub

    Private Sub txt_Password_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Password.TextChanged

        lbl_Result.Text = ""

    End Sub

    Private Sub txt_Password_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Password.Leave

        If txt_Password.Modified Then
            Dim D As New ConnectionData
            D.Box = cb_SAP.SelectedItem
            D.Login = txt_Login.Text
            D.SSO = CheckBox2.Checked
            D.Password = txt_Password.Text
            SC.SaveConnectionData(AC, D)
        End If

    End Sub

    Private Sub txt_Login_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_Login.Leave

        If txt_Login.Modified Then
            Dim D As New ConnectionData
            D.Box = cb_SAP.SelectedItem
            D.Login = txt_Login.Text
            D.SSO = CheckBox2.Checked
            D.Password = txt_Password.Text
            SC.SaveConnectionData(AC, D)
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        System.Diagnostics.Process.Start("https://sua.internal.pg.com/shua/shua/sap_reset.showsapuser")

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        If Not cb_SAP.SelectedItem Is Nothing Then
            Dim D As New ConnectionData
            D.Box = cb_SAP.SelectedItem
            D.Login = txt_Login.Text
            D.SSO = CheckBox2.Checked
            D.Password = txt_Password.Text
            SC.SaveConnectionData(AC, D)
            If CheckBox2.Checked Then
                txt_Password.Enabled = False
                CheckBox1.Checked = True
                CheckBox1.Enabled = False
            Else
                txt_Password.Enabled = True
                CheckBox1.Enabled = True
            End If
        End If

    End Sub

    Private Sub cb_SAP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_SAP.SelectedIndexChanged

        lbl_Result.Text = ""
        SetUp()

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked Then
            txt_Password.PasswordChar = "*"
        Else
            txt_Password.PasswordChar = Nothing
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not cb_SAP.SelectedItem Is Nothing Then
            lbl_Result.Text = "Testing..."
            txt_Result.Text = Nothing
            Me.Refresh()
            Dim D As ConnectionData = SC.GetConnectionData(cb_SAP.SelectedItem, UC, AC)
            Dim B As Boolean = SC.TestConnection(D)
            If B Then
                lbl_Result.Text = "Connection to " & cb_SAP.SelectedItem & " Succeded!"
            Else
                lbl_Result.Text = "Connection to " & cb_SAP.SelectedItem & " Failed!"
                txt_Result.Text = SC.Status
            End If
        Else
            MsgBox("Please pick the SAP System to test from the list")
        End If
    End Sub

    Private Sub Connection_Manager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        For Each Box As String In SC.BoxList
            cb_SAP.Items.Add(Box)
        Next

    End Sub

End Class
