Imports Microsoft.Office.Interop.Word

Public Class frmCreateAndApplyListStyle
    Private Sub frmCreateAndApplyListStyle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbNumberingStyle.DisplayMember = "Text"
        cbNumberingStyle.Items.Add(New With {.Text = "1-9 (pont nélkül)", .Value = WdListNumberStyle.wdListNumberStyleArabic})
        cbNumberingStyle.Items.Add(New With {.Text = "1.-9.", .Value = WdListNumberStyle.wdListNumberStyleOrdinal})
        cbNumberingStyle.Items.Add(New With {.Text = "I-X", .Value = WdListNumberStyle.wdListNumberStyleUppercaseRoman})
        cbNumberingStyle.Items.Add(New With {.Text = "i-x", .Value = WdListNumberStyle.wdListNumberStyleLowercaseRoman})
        cbNumberingStyle.Items.Add(New With {.Text = "A-Z", .Value = WdListNumberStyle.wdListNumberStyleUppercaseLetter})
        cbNumberingStyle.Items.Add(New With {.Text = "a-z", .Value = WdListNumberStyle.wdListNumberStyleLowercaseLetter})
        cbNumberingStyle.Items.Add(New With {.Text = "egy (nyelvfüggő)", .Value = WdListNumberStyle.wdListNumberStyleCardinalText})
        cbNumberingStyle.Items.Add(New With {.Text = "első (nyelvfüggő)", .Value = WdListNumberStyle.wdListNumberStyleOrdinalText})
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If Globals.ThisAddIn.Application.International(WdInternationalIndex.wdDecimalSeparator) = "." Then Me.tbIndent.Text = Replace(tbIndent.Text, ",", Globals.ThisAddIn.Application.International(WdInternationalIndex.wdDecimalSeparator))
        If Globals.ThisAddIn.Application.International(WdInternationalIndex.wdDecimalSeparator) = "," Then Me.tbIndent.Text = Replace(tbIndent.Text, ".", Globals.ThisAddIn.Application.International(WdInternationalIndex.wdDecimalSeparator))
        Dim blnStepped = cbxStepped.Checked
        HPWordStyleHelper.CreateMlevelListStyle(Me.tbMlvlName.Text, Me.tbIndent.Text, blnStepped, Me.cbNumberingStyle.SelectedValue, Me.tbSeparator.Text)
        HPWordHelper.CVarToFill("ListUsed", Me.tbMlvlName)
    End Sub
End Class