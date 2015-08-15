'************************************************************************
'* Name:            frm_guida.vb                                        *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Imports System.Drawing

Public Class frm_guida

    Private Sub btChiudi_Click() Handles btChiudi.Click

        On Error Resume Next

        Me.Close()

    End Sub

    Private Sub Form_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        On Error Resume Next

        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.Close()
        End If

    End Sub

    Private Sub Form_Load() Handles Me.Load

        On Error Resume Next

        Const OffsetTit = 20
        Const OffsetVal = 170

        Dim Y As Long = 0

        Dim bm As New Bitmap(pctComandi.Width, pctComandi.Height, Drawing.Imaging.PixelFormat.Format24bppRgb)
        Dim g = Graphics.FromImage(bm)
        Dim fTit As Font = New Font("Arial", 11)
        Dim fItem As Font = New Font("Arial", 9, FontStyle.Bold)
        Dim fVal As Font = New Font("Arial", 9)

        g.Clear(pctComandi.BackColor)

        g.DrawString("Tastiera", fTit, Brushes.Black, 5, Y)
        Y = Y + 10

        Y = Y + 16
        g.DrawString("Singola istruzione:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("<SHIFT><F8>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Entra nelle subroutine:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("<F8>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Esecuzione normale:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("<F5>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Ferma l'esecuzione:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("<BREAK>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Termina:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("<ESC>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 30
        g.DrawString("Metacomandi", fTit, Brushes.Black, 5, Y)
        Y = Y + 10

        Y = Y + 16
        g.DrawString("Valutazione:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@Watch <espressione>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Valutazione membri:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@Inside <espressione>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Ferma alla riga:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@Stop", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Salta alla riga:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@Goto <riga>", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Termina:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@End", fVal, Brushes.Black, OffsetVal, Y)

        Y = Y + 16
        g.DrawString("Commento persistente:", fItem, Brushes.Black, OffsetTit, Y)
        g.DrawString("'@Rem (o anche solo Rem)", fVal, Brushes.Black, OffsetVal, Y)


        Y = Y + 30
        g.DrawString("Ogni istruzione VBScript può essere posta", fVal, Brushes.Black, OffsetTit, Y)
        Y = Y + 16
        g.DrawString("sotto metacomando '@", fVal, Brushes.Black, OffsetTit, Y)

        g.Dispose()

        pctComandi.Image = bm

    End Sub

End Class