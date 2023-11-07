<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        lbLog = New ListBox()
        dtpFrom = New DateTimePicker()
        dtpTo = New DateTimePicker()
        Label1 = New Label()
        Label2 = New Label()
        btnExtract = New Button()
        Label3 = New Label()
        SuspendLayout()
        ' 
        ' lbLog
        ' 
        lbLog.FormattingEnabled = True
        lbLog.ItemHeight = 15
        lbLog.Location = New Point(8, 59)
        lbLog.Margin = New Padding(2, 2, 2, 2)
        lbLog.Name = "lbLog"
        lbLog.Size = New Size(544, 199)
        lbLog.TabIndex = 0
        ' 
        ' dtpFrom
        ' 
        dtpFrom.Format = DateTimePickerFormat.Short
        dtpFrom.Location = New Point(50, 16)
        dtpFrom.Margin = New Padding(2, 2, 2, 2)
        dtpFrom.Name = "dtpFrom"
        dtpFrom.Size = New Size(97, 23)
        dtpFrom.TabIndex = 1
        ' 
        ' dtpTo
        ' 
        dtpTo.Format = DateTimePickerFormat.Short
        dtpTo.Location = New Point(176, 15)
        dtpTo.Margin = New Padding(2, 2, 2, 2)
        dtpTo.Name = "dtpTo"
        dtpTo.Size = New Size(120, 23)
        dtpTo.TabIndex = 2
        ' 
        ' Label1
        ' 
        Label1.Location = New Point(8, 16)
        Label1.Margin = New Padding(2, 0, 2, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(38, 19)
        Label1.TabIndex = 3
        Label1.Text = "From"
        Label1.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' Label2
        ' 
        Label2.Location = New Point(151, 19)
        Label2.Margin = New Padding(2, 0, 2, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(21, 19)
        Label2.TabIndex = 4
        Label2.Text = "To"
        Label2.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' btnExtract
        ' 
        btnExtract.Location = New Point(320, 14)
        btnExtract.Margin = New Padding(2, 2, 2, 2)
        btnExtract.Name = "btnExtract"
        btnExtract.Size = New Size(80, 25)
        btnExtract.TabIndex = 5
        btnExtract.Text = "Extract"
        btnExtract.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(11, 44)
        Label3.Margin = New Padding(2, 0, 2, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(27, 15)
        Label3.TabIndex = 6
        Label3.Text = "Log"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(560, 270)
        Controls.Add(Label3)
        Controls.Add(btnExtract)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(dtpTo)
        Controls.Add(dtpFrom)
        Controls.Add(lbLog)
        FormBorderStyle = FormBorderStyle.FixedDialog
        Margin = New Padding(2, 2, 2, 2)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Extract Outlook Schedule"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lbLog As ListBox
    Friend WithEvents dtpFrom As DateTimePicker
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnExtract As Button
    Friend WithEvents Label3 As Label
End Class
