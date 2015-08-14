<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_monitor
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_monitor))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MnFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnEsci = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnVisualizza = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnOrizzontale = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnVerticale = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnDebug = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnAggiungiWatch2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnDelBreakpoint2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnEsegui = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnIstruzione2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnNextSub2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnFree2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnInterrompi = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnGuida = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnInformazioni = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimerComando = New System.Windows.Forms.Timer(Me.components)
        Me.SplitMonitor = New System.Windows.Forms.SplitContainer()
        Me.pctAttesa = New System.Windows.Forms.Label()
        Me.TxMonitor = New ryMedio64.monitor()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MnBreakPoint = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnGoTo = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnMostraRiga = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.MnAggiungiWatch = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnInterruzione = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnDelBreakpoint = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.MnIstruzione = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnNextSub = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnFree = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnTermina = New System.Windows.Forms.ToolStripMenuItem()
        Me.TxValutazioni = New System.Windows.Forms.TextBox()
        Me.TimerStatus = New System.Windows.Forms.Timer(Me.components)
        Me.TimerMonitor = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        CType(Me.SplitMonitor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitMonitor.Panel1.SuspendLayout()
        Me.SplitMonitor.Panel2.SuspendLayout()
        Me.SplitMonitor.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnFile, Me.MnVisualizza, Me.MnDebug, Me.MnEsegui, Me.MnHelp})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(881, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MnFile
        '
        Me.MnFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnEsci})
        Me.MnFile.Name = "MnFile"
        Me.MnFile.Size = New System.Drawing.Size(37, 20)
        Me.MnFile.Text = "&File"
        '
        'MnEsci
        '
        Me.MnEsci.Name = "MnEsci"
        Me.MnEsci.Size = New System.Drawing.Size(94, 22)
        Me.MnEsci.Text = "&Esci"
        '
        'MnVisualizza
        '
        Me.MnVisualizza.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnOrizzontale, Me.MnVerticale})
        Me.MnVisualizza.Name = "MnVisualizza"
        Me.MnVisualizza.Size = New System.Drawing.Size(69, 20)
        Me.MnVisualizza.Text = "&Visualizza"
        '
        'MnOrizzontale
        '
        Me.MnOrizzontale.Name = "MnOrizzontale"
        Me.MnOrizzontale.Size = New System.Drawing.Size(133, 22)
        Me.MnOrizzontale.Text = "Ori&zzontale"
        '
        'MnVerticale
        '
        Me.MnVerticale.Name = "MnVerticale"
        Me.MnVerticale.Size = New System.Drawing.Size(133, 22)
        Me.MnVerticale.Text = "Ver&ticale"
        '
        'MnDebug
        '
        Me.MnDebug.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnAggiungiWatch2, Me.MnDelBreakpoint2})
        Me.MnDebug.Name = "MnDebug"
        Me.MnDebug.Size = New System.Drawing.Size(54, 20)
        Me.MnDebug.Text = "&Debug"
        '
        'MnAggiungiWatch2
        '
        Me.MnAggiungiWatch2.Name = "MnAggiungiWatch2"
        Me.MnAggiungiWatch2.Size = New System.Drawing.Size(216, 22)
        Me.MnAggiungiWatch2.Text = "Aggiungi watch"
        '
        'MnDelBreakpoint2
        '
        Me.MnDelBreakpoint2.Name = "MnDelBreakpoint2"
        Me.MnDelBreakpoint2.Size = New System.Drawing.Size(216, 22)
        Me.MnDelBreakpoint2.Text = "Elimina tutte le interruzioni"
        '
        'MnEsegui
        '
        Me.MnEsegui.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnIstruzione2, Me.MnNextSub2, Me.MnFree2, Me.MnInterrompi})
        Me.MnEsegui.Name = "MnEsegui"
        Me.MnEsegui.Size = New System.Drawing.Size(53, 20)
        Me.MnEsegui.Text = "&Esegui"
        '
        'MnIstruzione2
        '
        Me.MnIstruzione2.Name = "MnIstruzione2"
        Me.MnIstruzione2.Size = New System.Drawing.Size(233, 22)
        Me.MnIstruzione2.Text = "&Istruzione (SHIFT+F8)"
        '
        'MnNextSub2
        '
        Me.MnNextSub2.Name = "MnNextSub2"
        Me.MnNextSub2.Size = New System.Drawing.Size(233, 22)
        Me.MnNextSub2.Text = "Entrando nelle subroutine (F8)"
        '
        'MnFree2
        '
        Me.MnFree2.Name = "MnFree2"
        Me.MnFree2.Size = New System.Drawing.Size(233, 22)
        Me.MnFree2.Text = "Senza inte&ruzioni (F5)"
        '
        'MnInterrompi
        '
        Me.MnInterrompi.Name = "MnInterrompi"
        Me.MnInterrompi.Size = New System.Drawing.Size(233, 22)
        Me.MnInterrompi.Text = "Interrompi (BREAK)"
        '
        'MnHelp
        '
        Me.MnHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnGuida, Me.MnInformazioni})
        Me.MnHelp.Name = "MnHelp"
        Me.MnHelp.Size = New System.Drawing.Size(24, 20)
        Me.MnHelp.Text = "&?"
        '
        'MnGuida
        '
        Me.MnGuida.Name = "MnGuida"
        Me.MnGuida.Size = New System.Drawing.Size(203, 22)
        Me.MnGuida.Text = "&Guida"
        '
        'MnInformazioni
        '
        Me.MnInformazioni.Name = "MnInformazioni"
        Me.MnInformazioni.Size = New System.Drawing.Size(203, 22)
        Me.MnInformazioni.Text = "Informazioni su r&yMedio"
        '
        'TimerComando
        '
        '
        'SplitMonitor
        '
        Me.SplitMonitor.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitMonitor.Location = New System.Drawing.Point(0, 27)
        Me.SplitMonitor.Name = "SplitMonitor"
        '
        'SplitMonitor.Panel1
        '
        Me.SplitMonitor.Panel1.BackColor = System.Drawing.Color.White
        Me.SplitMonitor.Panel1.Controls.Add(Me.pctAttesa)
        Me.SplitMonitor.Panel1.Controls.Add(Me.TxMonitor)
        Me.SplitMonitor.Panel1MinSize = 200
        '
        'SplitMonitor.Panel2
        '
        Me.SplitMonitor.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitMonitor.Panel2.Controls.Add(Me.TxValutazioni)
        Me.SplitMonitor.Size = New System.Drawing.Size(881, 513)
        Me.SplitMonitor.SplitterDistance = 718
        Me.SplitMonitor.TabIndex = 4
        '
        'pctAttesa
        '
        Me.pctAttesa.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pctAttesa.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pctAttesa.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pctAttesa.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pctAttesa.Location = New System.Drawing.Point(144, 205)
        Me.pctAttesa.Name = "pctAttesa"
        Me.pctAttesa.Padding = New System.Windows.Forms.Padding(100, 0, 0, 0)
        Me.pctAttesa.Size = New System.Drawing.Size(359, 48)
        Me.pctAttesa.TabIndex = 5
        Me.pctAttesa.Text = "Attendere..."
        Me.pctAttesa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.pctAttesa.Visible = False
        '
        'TxMonitor
        '
        Me.TxMonitor.BackColor = System.Drawing.Color.White
        Me.TxMonitor.ContextMenuStrip = Me.ContextMenuStrip1
        Me.TxMonitor.Location = New System.Drawing.Point(0, 0)
        Me.TxMonitor.Name = "TxMonitor"
        Me.TxMonitor.Riga = CType(0, Long)
        Me.TxMonitor.Size = New System.Drawing.Size(715, 510)
        Me.TxMonitor.TabIndex = 4
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnBreakPoint, Me.MnGoTo, Me.MnMostraRiga, Me.ToolStripSeparator1, Me.MnAggiungiWatch, Me.MnInterruzione, Me.MnDelBreakpoint, Me.ToolStripSeparator2, Me.MnIstruzione, Me.MnNextSub, Me.MnFree, Me.MnTermina})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(248, 236)
        '
        'MnBreakPoint
        '
        Me.MnBreakPoint.Name = "MnBreakPoint"
        Me.MnBreakPoint.Size = New System.Drawing.Size(247, 22)
        Me.MnBreakPoint.Text = "Esegui fino alla riga"
        '
        'MnGoTo
        '
        Me.MnGoTo.Name = "MnGoTo"
        Me.MnGoTo.Size = New System.Drawing.Size(247, 22)
        Me.MnGoTo.Text = "Salta direttamente alla riga"
        '
        'MnMostraRiga
        '
        Me.MnMostraRiga.Name = "MnMostraRiga"
        Me.MnMostraRiga.Size = New System.Drawing.Size(247, 22)
        Me.MnMostraRiga.Text = "Mostra riga corrente"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(244, 6)
        '
        'MnAggiungiWatch
        '
        Me.MnAggiungiWatch.Name = "MnAggiungiWatch"
        Me.MnAggiungiWatch.Size = New System.Drawing.Size(247, 22)
        Me.MnAggiungiWatch.Text = "Aggiungi watch"
        '
        'MnInterruzione
        '
        Me.MnInterruzione.Name = "MnInterruzione"
        Me.MnInterruzione.Size = New System.Drawing.Size(247, 22)
        Me.MnInterruzione.Text = "Aggiungi interruzione"
        '
        'MnDelBreakpoint
        '
        Me.MnDelBreakpoint.Name = "MnDelBreakpoint"
        Me.MnDelBreakpoint.Size = New System.Drawing.Size(247, 22)
        Me.MnDelBreakpoint.Text = "Elimina tutte le interruzioni"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(244, 6)
        '
        'MnIstruzione
        '
        Me.MnIstruzione.Name = "MnIstruzione"
        Me.MnIstruzione.Size = New System.Drawing.Size(247, 22)
        Me.MnIstruzione.Text = "Esegui istruzione"
        '
        'MnNextSub
        '
        Me.MnNextSub.Name = "MnNextSub"
        Me.MnNextSub.Size = New System.Drawing.Size(247, 22)
        Me.MnNextSub.Text = "Esegui entrando nelle subroutine"
        '
        'MnFree
        '
        Me.MnFree.Name = "MnFree"
        Me.MnFree.Size = New System.Drawing.Size(247, 22)
        Me.MnFree.Text = "Esecuzione senza interruzioni"
        '
        'MnTermina
        '
        Me.MnTermina.Name = "MnTermina"
        Me.MnTermina.Size = New System.Drawing.Size(247, 22)
        Me.MnTermina.Text = "Termina esecuzione"
        '
        'TxValutazioni
        '
        Me.TxValutazioni.BackColor = System.Drawing.Color.White
        Me.TxValutazioni.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxValutazioni.Location = New System.Drawing.Point(3, 3)
        Me.TxValutazioni.Multiline = True
        Me.TxValutazioni.Name = "TxValutazioni"
        Me.TxValutazioni.ReadOnly = True
        Me.TxValutazioni.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TxValutazioni.Size = New System.Drawing.Size(153, 507)
        Me.TxValutazioni.TabIndex = 0
        Me.TxValutazioni.WordWrap = False
        '
        'TimerStatus
        '
        Me.TimerStatus.Interval = 200
        '
        'TimerMonitor
        '
        Me.TimerMonitor.Interval = 1000
        '
        'frm_monitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(881, 541)
        Me.Controls.Add(Me.SplitMonitor)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frm_monitor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Monitor"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.SplitMonitor.Panel1.ResumeLayout(False)
        Me.SplitMonitor.Panel2.ResumeLayout(False)
        Me.SplitMonitor.Panel2.PerformLayout()
        CType(Me.SplitMonitor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitMonitor.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents MnFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnEsci As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnVisualizza As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnOrizzontale As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnVerticale As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnDebug As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnAggiungiWatch2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnDelBreakpoint2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnEsegui As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnIstruzione2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnNextSub2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnFree2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnGuida As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnInformazioni As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnInterrompi As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TimerComando As System.Windows.Forms.Timer
    Friend WithEvents SplitMonitor As System.Windows.Forms.SplitContainer
    Friend WithEvents TxMonitor As monitor
    Friend WithEvents TxValutazioni As System.Windows.Forms.TextBox
    Friend WithEvents pctAttesa As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents MnBreakPoint As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnGoTo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnMostraRiga As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MnAggiungiWatch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnInterruzione As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnDelBreakpoint As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MnIstruzione As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnNextSub As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnFree As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnTermina As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TimerStatus As System.Windows.Forms.Timer
    Friend WithEvents TimerMonitor As System.Windows.Forms.Timer
End Class
