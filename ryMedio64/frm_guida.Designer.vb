<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_guida
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_guida))
        Me.pctComandi = New System.Windows.Forms.PictureBox()
        Me.btChiudi = New System.Windows.Forms.Button()
        CType(Me.pctComandi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pctComandi
        '
        Me.pctComandi.BackColor = System.Drawing.Color.White
        Me.pctComandi.Location = New System.Drawing.Point(12, 12)
        Me.pctComandi.Name = "pctComandi"
        Me.pctComandi.Size = New System.Drawing.Size(440, 309)
        Me.pctComandi.TabIndex = 0
        Me.pctComandi.TabStop = False
        '
        'btChiudi
        '
        Me.btChiudi.Location = New System.Drawing.Point(12, 339)
        Me.btChiudi.Name = "btChiudi"
        Me.btChiudi.Size = New System.Drawing.Size(75, 23)
        Me.btChiudi.TabIndex = 1
        Me.btChiudi.Text = "&Chiudi"
        Me.btChiudi.UseVisualStyleBackColor = True
        '
        'frm_guida
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(465, 377)
        Me.Controls.Add(Me.btChiudi)
        Me.Controls.Add(Me.pctComandi)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frm_guida"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Guida"
        CType(Me.pctComandi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pctComandi As System.Windows.Forms.PictureBox
    Friend WithEvents btChiudi As System.Windows.Forms.Button
End Class
