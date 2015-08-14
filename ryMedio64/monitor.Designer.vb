<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class monitor
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(monitor))
        Me.ScrollVerticale = New System.Windows.Forms.VScrollBar()
        Me.ScrollOrizzontale = New System.Windows.Forms.HScrollBar()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ShapeStatement = New System.Windows.Forms.PictureBox()
        Me.ShapeSelettore = New System.Windows.Forms.PictureBox()
        Me.pctSchermo = New System.Windows.Forms.PictureBox()
        CType(Me.ShapeStatement, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ShapeSelettore, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pctSchermo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ScrollVerticale
        '
        Me.ScrollVerticale.Location = New System.Drawing.Point(508, 0)
        Me.ScrollVerticale.Name = "ScrollVerticale"
        Me.ScrollVerticale.Size = New System.Drawing.Size(17, 292)
        Me.ScrollVerticale.TabIndex = 1
        '
        'ScrollOrizzontale
        '
        Me.ScrollOrizzontale.Location = New System.Drawing.Point(0, 292)
        Me.ScrollOrizzontale.Name = "ScrollOrizzontale"
        Me.ScrollOrizzontale.Size = New System.Drawing.Size(508, 17)
        Me.ScrollOrizzontale.TabIndex = 2
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "breakpoint")
        '
        'ShapeStatement
        '
        Me.ShapeStatement.BackColor = System.Drawing.Color.Transparent
        Me.ShapeStatement.BackgroundImage = Global.ryMedio64.My.Resources.Resources.semiyellow
        Me.ShapeStatement.Location = New System.Drawing.Point(62, 34)
        Me.ShapeStatement.Name = "ShapeStatement"
        Me.ShapeStatement.Size = New System.Drawing.Size(325, 18)
        Me.ShapeStatement.TabIndex = 5
        Me.ShapeStatement.TabStop = False
        '
        'ShapeSelettore
        '
        Me.ShapeSelettore.BackColor = System.Drawing.Color.Transparent
        Me.ShapeSelettore.BackgroundImage = Global.ryMedio64.My.Resources.Resources.dashed
        Me.ShapeSelettore.ErrorImage = Nothing
        Me.ShapeSelettore.InitialImage = Nothing
        Me.ShapeSelettore.Location = New System.Drawing.Point(62, 98)
        Me.ShapeSelettore.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeSelettore.Name = "ShapeSelettore"
        Me.ShapeSelettore.Size = New System.Drawing.Size(325, 18)
        Me.ShapeSelettore.TabIndex = 4
        Me.ShapeSelettore.TabStop = False
        '
        'pctSchermo
        '
        Me.pctSchermo.BackColor = System.Drawing.Color.White
        Me.pctSchermo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.pctSchermo.Location = New System.Drawing.Point(0, 0)
        Me.pctSchermo.Margin = New System.Windows.Forms.Padding(0)
        Me.pctSchermo.Name = "pctSchermo"
        Me.pctSchermo.Size = New System.Drawing.Size(508, 292)
        Me.pctSchermo.TabIndex = 0
        Me.pctSchermo.TabStop = False
        '
        'monitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ShapeStatement)
        Me.Controls.Add(Me.ShapeSelettore)
        Me.Controls.Add(Me.ScrollOrizzontale)
        Me.Controls.Add(Me.ScrollVerticale)
        Me.Controls.Add(Me.pctSchermo)
        Me.Name = "monitor"
        Me.Size = New System.Drawing.Size(525, 311)
        CType(Me.ShapeStatement, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ShapeSelettore, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pctSchermo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pctSchermo As System.Windows.Forms.PictureBox
    Friend WithEvents ScrollVerticale As System.Windows.Forms.VScrollBar
    Friend WithEvents ScrollOrizzontale As System.Windows.Forms.HScrollBar
    Friend WithEvents ShapeSelettore As System.Windows.Forms.PictureBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ShapeStatement As System.Windows.Forms.PictureBox
End Class
