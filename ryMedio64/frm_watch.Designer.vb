<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_watch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_watch))
        Me.lbDescr = New System.Windows.Forms.Label()
        Me.txRiga = New System.Windows.Forms.TextBox()
        Me.btAggiungi = New System.Windows.Forms.Button()
        Me.btChiudi = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lbDescr
        '
        Me.lbDescr.AutoSize = True
        Me.lbDescr.Location = New System.Drawing.Point(16, 27)
        Me.lbDescr.Name = "lbDescr"
        Me.lbDescr.Size = New System.Drawing.Size(187, 13)
        Me.lbDescr.TabIndex = 0
        Me.lbDescr.Text = "Selezione l'espressione da aggiungere"
        '
        'txRiga
        '
        Me.txRiga.Location = New System.Drawing.Point(19, 52)
        Me.txRiga.Name = "txRiga"
        Me.txRiga.Size = New System.Drawing.Size(602, 20)
        Me.txRiga.TabIndex = 1
        '
        'btAggiungi
        '
        Me.btAggiungi.Location = New System.Drawing.Point(19, 96)
        Me.btAggiungi.Name = "btAggiungi"
        Me.btAggiungi.Size = New System.Drawing.Size(75, 23)
        Me.btAggiungi.TabIndex = 2
        Me.btAggiungi.Text = "&Aggiungi"
        Me.btAggiungi.UseVisualStyleBackColor = True
        '
        'btChiudi
        '
        Me.btChiudi.Location = New System.Drawing.Point(100, 96)
        Me.btChiudi.Name = "btChiudi"
        Me.btChiudi.Size = New System.Drawing.Size(75, 23)
        Me.btChiudi.TabIndex = 3
        Me.btChiudi.Text = "&Chiudi"
        Me.btChiudi.UseVisualStyleBackColor = True
        '
        'frm_watch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(646, 140)
        Me.Controls.Add(Me.btChiudi)
        Me.Controls.Add(Me.btAggiungi)
        Me.Controls.Add(Me.txRiga)
        Me.Controls.Add(Me.lbDescr)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frm_watch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aggiunta watch"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbDescr As System.Windows.Forms.Label
    Friend WithEvents txRiga As System.Windows.Forms.TextBox
    Friend WithEvents btAggiungi As System.Windows.Forms.Button
    Friend WithEvents btChiudi As System.Windows.Forms.Button
End Class
