Option Explicit On

Public Class frm_about

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

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        System.Diagnostics.Process.Start(LinkLabel1.Text)

    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked

        System.Diagnostics.Process.Start(LinkLabel2.Text)

    End Sub

End Class