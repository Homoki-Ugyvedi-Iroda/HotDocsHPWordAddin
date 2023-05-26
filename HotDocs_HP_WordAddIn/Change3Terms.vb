Imports System.Windows.Forms
Imports HPW = HotDocs_HP_WordAddIn.HPWordHelper
#If NoHotDocs <> "Y" Then
Imports HD = HotDocs
#End If

Public Class frmChange3Terms
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        CheckSelectedTemplateTerms
    End Sub
    Private Sub CheckSelectedTemplateTerms()

#If NoHotDocs <> "Y" Then
        Dim MyAnswColl
        Try
            MyAnswColl = Globals.ThisAddIn.LastUsedAnswerCollection
            If MyAnswColl Is Nothing Then
                Globals.ThisAddIn.logger.Error("MyAnswColl Is Nothing")
                Exit Sub
            End If
        Catch ex As Exception
            Globals.ThisAddIn.HotDocsInstalled = False
            Exit Sub
        End Try
        Dim ClientRole, PartnerRole, FeeRole, SubjectRole As String
        Try
            ClientRole = MyAnswColl.Item("ClientContractualRoleDefaultName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            If Not String.IsNullOrWhiteSpace(ClientRole) Then tbMegrendelő.Text = ClientRole.Trim
        Catch ex As Exception
            Globals.ThisAddIn.logger.Error("CheckSelectedTemplateError: ClientRole" & ex.Message & "_" & ex.StackTrace)
        End Try
        Try
            PartnerRole = MyAnswColl.Item("PartnerContractualRoleDefaultName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            If Not String.IsNullOrWhiteSpace(PartnerRole) Then tbVállalkozó.Text = PartnerRole.Trim
        Catch ex As Exception
            Globals.ThisAddIn.logger.Error("CheckSelectedTemplateError: PartnerRole" & ex.Message & "_" & ex.StackTrace)
        End Try
        Try
            FeeRole = MyAnswColl.Item("FeeContractualRoleDefaultName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            If Not String.IsNullOrWhiteSpace(FeeRole) Then tbDij.Text = FeeRole.Trim
        Catch ex As Exception
            Globals.ThisAddIn.logger.Error("CheckSelectedTemplateError: FeeRole" & ex.Message & "_" & ex.StackTrace)
        End Try
        Try
            SubjectRole = MyAnswColl.Item("SubjectContractualRoleDefaultName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            If Not String.IsNullOrWhiteSpace(SubjectRole) Then tbTárgy.Text = SubjectRole.Trim
        Catch ex As Exception
            Globals.ThisAddIn.logger.Error("CheckSelectedTemplateError: SubjectRole" & ex.Message & "_" & ex.StackTrace)
        End Try
#End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.Hide()
        HPW.ChangeTerms("Megrendelő", tbMegrendelő.Text, True)
        HPW.ChangeTerms("Vállalkozó", tbVállalkozó.Text, True)
        HPW.ChangeTerms("Tárg", tbTárgy.Text, True)
        HPW.ChangeTerms("Díj", tbDij.Text, True)
        Me.Close()
    End Sub

    Private Sub tbChangeTerms_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles tbMegrendelő.Validating, tbDij.Validating, tbTárgy.Validating, tbVállalkozó.Validating
        Dim senderBox As TextBox = TryCast(sender, TextBox)
        senderBox.Text = senderBox.Text.Trim
        '        If senderBox.Text.Contains(" ") Then
        '       Dim senderWords As String() = senderBox.Text.Split(" ")
        '      If senderWords.Count > 0 Then senderBox.Text = senderWords.Last
        '     End If
    End Sub
End Class