<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateAndApplyListStyle
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbMlvlName = New System.Windows.Forms.TextBox()
        Me.cbxStepped = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbNumberingStyle = New System.Windows.Forms.ComboBox()
        Me.tbIndent = New System.Windows.Forms.TextBox()
        Me.tbSeparator = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.46974!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.53026!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnOK, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCancel, 2, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.tbMlvlName, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cbxStepped, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cbNumberingStyle, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.tbIndent, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.tbSeparator, 1, 4)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40.35088!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 59.64912!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 67.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(449, 212)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(260, 184)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 19)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(360, 184)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 19)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Lista neve?"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(236, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Behúzás és tabulátor távolsága? (0-20 tört szám)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(201, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Számozási stílus (minden szinten azonos)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(232, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "További elválasztó a felsorolás után (opcionális)"
        '
        'tbMlvlName
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.tbMlvlName, 2)
        Me.tbMlvlName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbMlvlName.Location = New System.Drawing.Point(251, 3)
        Me.tbMlvlName.Name = "tbMlvlName"
        Me.tbMlvlName.Size = New System.Drawing.Size(195, 20)
        Me.tbMlvlName.TabIndex = 8
        '
        'cbxStepped
        '
        Me.cbxStepped.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.cbxStepped.AutoSize = True
        Me.cbxStepped.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbxStepped.Location = New System.Drawing.Point(251, 33)
        Me.cbxStepped.Name = "cbxStepped"
        Me.cbxStepped.Size = New System.Drawing.Size(15, 14)
        Me.cbxStepped.TabIndex = 2
        Me.cbxStepped.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(130, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Lépcsős tagolású legyen?"
        '
        'cbNumberingStyle
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.cbNumberingStyle, 2)
        Me.cbNumberingStyle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cbNumberingStyle.FormattingEnabled = True
        Me.cbNumberingStyle.Location = New System.Drawing.Point(251, 89)
        Me.cbNumberingStyle.Name = "cbNumberingStyle"
        Me.cbNumberingStyle.Size = New System.Drawing.Size(195, 21)
        Me.cbNumberingStyle.TabIndex = 10
        '
        'tbIndent
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.tbIndent, 2)
        Me.tbIndent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbIndent.Location = New System.Drawing.Point(251, 60)
        Me.tbIndent.Name = "tbIndent"
        Me.tbIndent.Size = New System.Drawing.Size(195, 20)
        Me.tbIndent.TabIndex = 11
        '
        'tbSeparator
        '
        Me.tbSeparator.Location = New System.Drawing.Point(251, 117)
        Me.tbSeparator.Name = "tbSeparator"
        Me.tbSeparator.Size = New System.Drawing.Size(93, 20)
        Me.tbSeparator.TabIndex = 12
        '
        'frmCreateAndApplyListStyle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 212)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "frmCreateAndApplyListStyle"
        Me.Text = "Multilevel lista stílus létrehozása és címsorra alkalmazása"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents cbxStepped As Windows.Forms.CheckBox
    Friend WithEvents tbMlvlName As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents cbNumberingStyle As Windows.Forms.ComboBox
    Friend WithEvents tbIndent As Windows.Forms.TextBox
    Friend WithEvents tbSeparator As Windows.Forms.TextBox
End Class
