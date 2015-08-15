'************************************************************************
'* Name:            frm_monitor.vb                                      *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Imports ryMedio64.debugEnumerator.EnumQueryType

Public Class frm_monitor

    Public Directory As String
    Public QueryType As Long
    Public ProcID As String
    Public Polling As Long

    Public Event WriteCommand(ByRef strCommand As String)
    Public Event ReadMonitor(ByRef strCode As String)
    Public Event ReadStatus(ByRef strStatus As String)
    Public Event StatusExists(ByRef ThereIs As Boolean)
    Public Event MonitorExists(ByRef ThereIs As Boolean)

    Dim PuntiAttesa As Long
    Dim UltimaIstr As Long = 0

    Dim UltimoComando As Single

    Dim Visualizzazione As Long
    Dim ContaResize As Long = 0

    Dim StatoWin As StructStatoWin
    Dim FlagActivate As Boolean = False

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub ScriviComando(Comando As String)

        Dim PathFile As String

        Try

            If pctAttesa.Visible = False Or Comando = "break:0" Then

                If 1000 * Math.Abs(Timer - UltimoComando) > Me.Polling Or Comando = "break:0" Then

                    '------------------------------
                    ' Accendo il segnale di attesa
                    '------------------------------

                    PuntiAttesa = 3
                    AvanzaAttesa()
                    pctAttesa.Visible = True

                    If QueryType = qtFile Then

                        PathFile = Me.Directory + "\" + Me.ProcID + ".CMN"

                        My.Computer.FileSystem.WriteAllText(PathFile, Comando, False)

                    Else

                        RaiseEvent WriteCommand(Comando)

                    End If

                    UltimoComando = Timer

                    TimerStatus.Enabled = True

                End If

            End If

        Finally

        End Try

    End Sub

    Private Sub ComandoBreak(Riga As Long)

        On Error Resume Next

        ScriviComando("break:" & Riga)

    End Sub

    Private Sub ComandoWatch(Espressione As String)

        On Error Resume Next

        ScriviComando("watch:" & Espressione)

    End Sub

    Private Sub ComandoEnd()

        On Error Resume Next

        ScriviComando("end")

    End Sub

    Private Sub ComandoGoTo(Riga As Long)

        On Error Resume Next

        ScriviComando("goto:" & Riga)

    End Sub

    Private Sub ComandoLibera()

        On Error Resume Next

        Dim Buffer As String

        Buffer = TxMonitor.ListBreakPoint

        ScriviComando("free:" & Buffer)

    End Sub

    Private Sub ComandoNext()

        On Error Resume Next

        ScriviComando("next")

    End Sub

    Private Sub ComandoNextSub()

        On Error Resume Next

        ScriviComando("next:1")

    End Sub

    Private Sub AvanzaAttesa()

        On Error Resume Next

        pctAttesa.Text = "Attendere" & (New String(".", PuntiAttesa))
        PuntiAttesa = PuntiAttesa + 1

        If PuntiAttesa > 10 Then
            PuntiAttesa = 0
        End If

    End Sub

    Private Sub TxMonitor_DblClick(Riga As Long, Numerazione As Boolean) Handles TxMonitor.DblClick

        On Error Resume Next

        If Numerazione Then

            TxMonitor.BreakPoint(Riga) = Not TxMonitor.BreakPoint(Riga)
            TxMonitor.monRefresh()

        Else

            AggiungiWatch(Riga)

        End If

    End Sub

    Public Sub DisposizioneForm()

        On Error Resume Next

        Dim Proporzione As Single
        Dim EditW As Single
        Dim EditH As Single
        Dim ValuW As Single
        Dim ValuH As Single

        ContaResize = ContaResize + 1

        If ContaResize = 1 Then

            TxMonitor.Left = 0
            TxMonitor.Top = 0
            TxMonitor.Width = SplitMonitor.Panel1.Width + 3
            TxMonitor.Height = SplitMonitor.Panel1.Height + 3
            TxMonitor.monRefresh()

            TxValutazioni.Left = 0
            TxValutazioni.Top = 0
            TxValutazioni.Width = SplitMonitor.Panel2.Width
            TxValutazioni.Height = SplitMonitor.Panel2.Height

            DisponiAttendi()

        End If

        ContaResize = ContaResize - 1

    End Sub

    Private Sub DisponiAttendi()

        On Error Resume Next

        pctAttesa.Left = TxMonitor.Left + (TxMonitor.Width - pctAttesa.Width) / 2
        pctAttesa.Top = TxMonitor.Top + (TxMonitor.Height - pctAttesa.Height) / 2

    End Sub

    Private Sub AggiungiWatch(Riga As Long)

        On Error Resume Next

        Dim FrmWatch As frm_watch = New frm_watch

        FrmWatch.Codice = TxMonitor.Text

        If Riga >= 1 Then
            FrmWatch.Riga = TxMonitor.GetRow(Riga)
        Else
            FrmWatch.Riga = ""
        End If

        FrmWatch.ShowDialog()

        If FrmWatch.Cancel = False Then

            ComandoWatch(FrmWatch.Espressione)

        End If

    End Sub

    Private Sub ShowGuide()

        On Error Resume Next

        Dim f As frm_guida = New frm_guida

        f.ShowDialog()

    End Sub

    Private Sub ShowAbout()

        On Error Resume Next

        Dim f As frm_about = New frm_about

        f.ShowDialog()

    End Sub

    Private Sub frm_monitor_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        On Error Resume Next

        DisposizioneForm()

    End Sub

    Private Sub MnOrizzontale_Click(sender As Object, e As EventArgs) Handles MnOrizzontale.Click

        On Error Resume Next

        SplitMonitor.Orientation = System.Windows.Forms.Orientation.Horizontal

        Visualizzazione = 0
        MnOrizzontale.Checked = True
        MnVerticale.Checked = False
        SplitMonitor.SplitterDistance = 3 * SplitMonitor.Height / 4

    End Sub

    Private Sub MnVerticale_Click(sender As Object, e As EventArgs) Handles MnVerticale.Click

        On Error Resume Next

        SplitMonitor.Orientation = System.Windows.Forms.Orientation.Vertical

        Visualizzazione = 1
        MnOrizzontale.Checked = False
        MnVerticale.Checked = True
        SplitMonitor.SplitterDistance = 3 * SplitMonitor.Width / 4

    End Sub

    Private Sub MnInformazioni_Click(sender As Object, e As EventArgs) Handles MnInformazioni.Click

        ShowAbout()

    End Sub

    Private Sub MnGuida_Click(sender As Object, e As EventArgs) Handles MnGuida.Click

        ShowGuide()

    End Sub

    Private Sub MnEsci_Click(sender As Object, e As EventArgs) Handles MnEsci.Click

        On Error Resume Next

        Me.Close()

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

        On Error Resume Next

        Dim R As Long

        If pctAttesa.Visible Then

            e.Cancel = True

        Else

            If UltimaIstr = 0 Then

                R = TxMonitor.RigaCliccata

                If R >= 0 Then

                    MnBreakPoint.Enabled = True
                    MnBreakPoint.Text = "Esegui fino alla riga " & R

                    MnGoTo.Enabled = True
                    MnGoTo.Text = "Salta direttamente alla riga " & R

                    MnInterruzione.Enabled = True

                    If TxMonitor.BreakPoint(R) Then
                        MnInterruzione.Text = "Togli interruzione alla riga " & R
                    Else
                        MnInterruzione.Text = "Aggiungi interruzione alla riga " & R
                    End If

                Else

                    MnBreakPoint.Enabled = False
                    MnBreakPoint.Text = "Esegui fino alla riga"

                    MnGoTo.Enabled = False
                    MnGoTo.Text = "Salta direttamente alla riga"

                    MnInterruzione.Enabled = False
                    MnInterruzione.Text = "Aggiungi interruzione"

                End If

                MnAggiungiWatch.Enabled = True
                MnDelBreakpoint.Enabled = True
                MnFree.Enabled = True
                MnIstruzione.Enabled = True
                MnNextSub.Enabled = True

            Else

                MnBreakPoint.Enabled = False
                MnBreakPoint.Text = "Esegui fino alla riga"

                MnGoTo.Enabled = False
                MnGoTo.Text = "Salta direttamente alla riga"

                MnInterruzione.Enabled = False
                MnInterruzione.Text = "Aggiungi interruzione"

                MnAggiungiWatch.Enabled = False
                MnDelBreakpoint.Enabled = False
                MnFree.Enabled = False
                MnIstruzione.Enabled = False
                MnNextSub.Enabled = False

            End If

        End If

    End Sub

    Private Sub TimerStatus_Tick(sender As Object, e As EventArgs) Handles TimerStatus.Tick

        On Error Resume Next

        Dim Buffer As String = ""
        Dim PathFile As String = ""

        Dim Riga As Long
        Dim Commento As String
        Dim MessErr As String
        Dim MaxEval As Long
        Dim I As Long
        Dim Esiste As Boolean
        Dim V As Object

        If QueryType = qtFile Then

            PathFile = Me.Directory + "\" + Me.ProcID + ".STA"

            Esiste = (Dir(PathFile) <> "")

        Else

            Esiste = False
            RaiseEvent StatusExists(Esiste)

        End If

        If Esiste Then

            TimerStatus.Enabled = False

            If QueryType = qtFile Then

                Buffer = My.Computer.FileSystem.ReadAllText(PathFile)

                My.Computer.FileSystem.DeleteFile(PathFile)

            Else

                RaiseEvent ReadStatus(Buffer)

            End If

            V = Split(Buffer, vbCrLf)

            MessErr = V(0)

            If MessErr = "end" Then

                Me.Close()

            Else

                Commento = V(1)
                UltimaIstr = Val(V(2))
                Riga = Val(V(3))

                '-------------
                ' Valutazioni
                '-------------

                MaxEval = UBound(V) - 3

                If Commento <> "" Then
                    Buffer = "Commento: " + Commento + vbCrLf
                Else
                    Buffer = ""
                End If

                For I = 1 To MaxEval
                    Buffer = Buffer & V(I + 3) & vbCrLf
                Next I

                TxValutazioni.Text = Buffer

                '---------
                ' Monitor
                '---------

                If Riga <= 0 Then
                    Riga = 1
                End If

                TxMonitor.Riga = Riga - 1

                If Riga > TxMonitor.VisibleRows \ 2 Then
                    TxMonitor.SetTopRow(Riga - TxMonitor.VisibleRows \ 2)
                End If

                TxMonitor.monRefresh()

                If MessErr <> "" Then
                    MsgBox(MessErr, , "ryMedio - Riscontrato errore!")
                End If

            End If

            '-----------------------------
            ' Spengo il segnale di attesa
            '-----------------------------

            pctAttesa.Visible = False

        Else

            AvanzaAttesa()

        End If

    End Sub

    Private Sub frm_monitor_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        On Error Resume Next

        Select Case e.KeyCode

            Case System.Windows.Forms.Keys.F8

                If e.Shift Then
                    ComandoNext()
                Else
                    ComandoNextSub()
                End If

            Case System.Windows.Forms.Keys.F5

                ComandoLibera()

            Case System.Windows.Forms.Keys.Escape

                ComandoEnd()

            Case 19

                If pctAttesa.Visible Then
                    ComandoBreak(0)
                End If

        End Select

    End Sub

    Private Sub MnIstruzione_Click(sender As Object, e As EventArgs) Handles MnIstruzione.Click

        On Error Resume Next

        TimerComando.Tag = "next"
        TimerComando.Enabled = True

    End Sub

    Private Sub MnAggiungiWatch_Click(sender As Object, e As EventArgs) Handles MnAggiungiWatch.Click

        On Error Resume Next

        AggiungiWatch(-1)

    End Sub

    Private Sub MnAggiungiWatch2_Click(sender As Object, e As EventArgs) Handles MnAggiungiWatch2.Click

        On Error Resume Next

        AggiungiWatch(TxMonitor.RigaCliccata)

    End Sub

    Private Sub MnBreakPoint_Click(sender As Object, e As EventArgs) Handles MnBreakPoint.Click

        On Error Resume Next

        TimerComando.Tag = "break:" & TxMonitor.RigaCliccata
        TimerComando.Enabled = True

    End Sub

    Private Sub MnDelBreakpoint_Click(sender As Object, e As EventArgs) Handles MnDelBreakpoint.Click

        On Error Resume Next

        TxMonitor.ClearBreakPoint()

    End Sub

    Private Sub MnDelBreakpoint2_Click(sender As Object, e As EventArgs) Handles MnDelBreakpoint2.Click

        On Error Resume Next

        TxMonitor.ClearBreakPoint()

    End Sub

    Private Sub MnEsegui_Click(sender As Object, e As EventArgs) Handles MnEsegui.Click

        On Error Resume Next

        MnInterrompi.Enabled = pctAttesa.Visible

    End Sub

    Private Sub MnFree_Click(sender As Object, e As EventArgs) Handles MnFree.Click

        On Error Resume Next

        Dim Buffer As String

        Buffer = TxMonitor.ListBreakPoint

        TimerComando.Tag = "free:" & Buffer
        TimerComando.Enabled = True

    End Sub

    Private Sub MnFree2_Click(sender As Object, e As EventArgs) Handles MnFree2.Click

        On Error Resume Next

        Dim Buffer As String

        Buffer = TxMonitor.ListBreakPoint

        TimerComando.Tag = "free:" & Buffer
        TimerComando.Enabled = True

    End Sub

    Private Sub MnGoTo_Click(sender As Object, e As EventArgs) Handles MnGoTo.Click

        On Error Resume Next

        TimerComando.Tag = "goto:" & TxMonitor.RigaCliccata
        TimerComando.Enabled = True

    End Sub

    Private Sub MnInterrompi_Click(sender As Object, e As EventArgs) Handles MnInterrompi.Click

        On Error Resume Next

        TimerComando.Tag = "break:0"
        TimerComando.Enabled = True

    End Sub

    Private Sub MnInterruzione_Click(sender As Object, e As EventArgs) Handles MnInterruzione.Click

        On Error Resume Next

        Dim Rg As Long

        Rg = TxMonitor.RigaCliccata

        If Rg > 0 Then

            TxMonitor.BreakPoint(Rg) = Not TxMonitor.BreakPoint(Rg)
            TxMonitor.monRefresh()

        End If

    End Sub

    Private Sub MnIstruzione2_Click(sender As Object, e As EventArgs) Handles MnIstruzione2.Click

        On Error Resume Next

        TimerComando.Tag = "next"
        TimerComando.Enabled = True

    End Sub

    Private Sub MnMostraRiga_Click(sender As Object, e As EventArgs) Handles MnMostraRiga.Click

        On Error Resume Next

        If TxMonitor.Riga - TxMonitor.VisibleRows \ 2 > 0 Then
            TxMonitor.SetTopRow(TxMonitor.Riga - TxMonitor.VisibleRows \ 2)
        Else
            TxMonitor.SetTopRow(0)
        End If

        TxMonitor.monRefresh()

    End Sub

    Private Sub MnNextSub_Click(sender As Object, e As EventArgs) Handles MnNextSub.Click

        On Error Resume Next

        TimerComando.Tag = "next:1"
        TimerComando.Enabled = True

    End Sub

    Private Sub MnNextSub2_Click(sender As Object, e As EventArgs) Handles MnNextSub2.Click

        On Error Resume Next

        ScriviComando("next:1")

    End Sub

    Private Sub MnTermina_Click(sender As Object, e As EventArgs) Handles MnTermina.Click

        On Error Resume Next

        TimerComando.Tag = "end"
        TimerComando.Enabled = True

    End Sub

    Private Sub TimerMonitor_Tick(sender As Object, e As EventArgs) Handles TimerMonitor.Tick

        On Error Resume Next

        Dim Buffer As String = ""
        Dim PathFile As String = ""
        Dim Esiste As Boolean

        If QueryType = qtFile Then

            PathFile = Me.Directory + "\" + Me.ProcID + ".MON"

            Esiste = Dir(PathFile) <> ""

        Else

            Esiste = False
            RaiseEvent MonitorExists(Esiste)

        End If

        If Esiste Then

            TimerMonitor.Enabled = False

            If QueryType = qtFile Then

                Buffer = My.Computer.FileSystem.ReadAllText(PathFile)

            Else

                RaiseEvent ReadMonitor(Buffer)

            End If

            TxMonitor.Text = Buffer

            TimerStatus.Enabled = True

        End If

    End Sub

    Private Sub frm_monitor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        On Error Resume Next

        StatoWin.L = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Left", Str((My.Computer.Screen.Bounds.Width - Me.Width) / 2)))
        StatoWin.T = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Top", Str((My.Computer.Screen.Bounds.Height - Me.Height) / 2)))
        StatoWin.W = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Width", Str(Me.Width)))
        StatoWin.H = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Height", Str(Me.Height)))
        StatoWin.M = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "WindowState", "0"))
        StatoWin.B = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Offset", Str(SplitMonitor.SplitterDistance)))
        Visualizzazione = Val(RegReadSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "TileMode", Str(SplitMonitor.Orientation)))

    End Sub

    Private Sub frm_monitor_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        TimerMonitor.Enabled = False
        TimerStatus.Enabled = False

        If Me.WindowState = 0 Then

            RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Left", Str(Me.Left))
            RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Top", Str(Me.Top))
            RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Width", Str(Me.Width))
            RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Height", Str(Me.Height))

        End If

        RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "WindowState", Str(Me.WindowState))
        RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "TileMode", Str(Visualizzazione))
        RegWriteSetting(EnumRegHKey.RegCurrentUser, "ryMedio", "", "Offset", Str(SplitMonitor.SplitterDistance))

        ScriviComando("end")

    End Sub

    Private Sub frm_monitor_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp

        On Error Resume Next

        TxMonitor.Focus()

    End Sub

    Private Sub frm_monitor_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated

        On Error Resume Next

        If FlagActivate = False Then

            FlagActivate = True

            ContaResize = 1

            If Visualizzazione = 1 Then

                SplitMonitor.Orientation = System.Windows.Forms.Orientation.Vertical

                MnOrizzontale.Checked = False
                MnVerticale.Checked = True

            Else

                SplitMonitor.Orientation = System.Windows.Forms.Orientation.Horizontal

                MnOrizzontale.Checked = True
                MnVerticale.Checked = False

            End If

            If StatoWin.M = 1 Then
                StatoWin.M = 0
            End If

            Me.WindowState = StatoWin.M

            '------------------------------
            ' Accendo il segnale di attesa
            '------------------------------

            PuntiAttesa = 3
            AvanzaAttesa()
            pctAttesa.Visible = True

            TimerStatus.Interval = Me.Polling
            TimerMonitor.Enabled = True

            If StatoWin.M = 0 Then

                Me.Left = StatoWin.L
                Me.Top = StatoWin.T
                Me.Width = StatoWin.W
                Me.Height = StatoWin.H

            End If

            SplitMonitor.SplitterDistance = StatoWin.B

            ContaResize = 0
            DisposizioneForm()

        End If

    End Sub

    Private Sub SplitMonitor_SplitterMoved(sender As Object, e As System.Windows.Forms.SplitterEventArgs) Handles SplitMonitor.SplitterMoved

        On Error Resume Next

        DisposizioneForm()

    End Sub

    Private Sub TimerComando_Tick(sender As Object, e As EventArgs) Handles TimerComando.Tick

        On Error Resume Next

        TimerComando.Enabled = False
        ScriviComando(TimerComando.Tag)

    End Sub

End Class