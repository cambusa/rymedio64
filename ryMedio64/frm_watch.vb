Option Explicit On

Imports MSScriptControl

Public Class frm_watch

    Public Cancel As Boolean
    Public Espressione As String
    Public Riga As String
    Public Codice As String

    Dim ObjScript As ScriptControl

    Dim MessaggioNiente As String

    Private Sub btAggiungi_Click()

        On Error Resume Next

        Dim Valida As Boolean

        Valida = False
        Me.Espressione = ""

        If Me.Riga <> "" Then

            If txRiga.SelectionLength > 0 Then

                Me.Espressione = Trim(Mid(txRiga.Text, txRiga.SelectionStart + 1, txRiga.SelectionLength))

            End If

        Else

            Me.Espressione = Trim(txRiga.Text)

        End If

        If Me.Espressione <> "" Then

            Valida = True

        End If

        If Valida Then

            ObjScript.Reset
            ObjScript.AddCode(Me.Codice)

            ObjScript.Error.Clear
            ObjScript.AddCode("sub test___tmp():x=" & Me.Espressione & ":end sub")

            If ObjScript.Error.Number = 0 Then

                Me.Cancel = False
                Me.Hide()

            Else

                MsgBox(ObjScript.Error.Description, , "Aggiunta watch")

            End If

        Else

            MsgBox(MessaggioNiente, , "Aggiunta watch")

        End If

    End Sub

    Private Sub btChiudi_Click()

        On Error Resume Next

        Me.Hide()

    End Sub

    Private Sub frm_watch_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        On Error Resume Next

        ObjScript = Nothing

    End Sub

    Private Sub frm_watch_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        On Error Resume Next

        ObjScript = New ScriptControl
        ObjScript.Language = "vbscript"
        ObjScript.Timeout = -1

        txRiga.Text = Trim(Me.Riga)

        If Me.Riga <> "" Then
            MessaggioNiente = "Nessuna espressione selezionata"
            lbDescr.Text = "Seleziona l'espressione da aggiungere"
        Else
            MessaggioNiente = "Specificare l'espressione da aggiungere"
            lbDescr.Text = "Espressione da aggiungere"
        End If

        Me.Espressione = ""
        Me.Codice = ""

        Me.Cancel = True

    End Sub

    Private Sub frm_watch_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        On Error Resume Next

        If e.KeyCode = 27 Then

            btChiudi_Click()

        ElseIf e.KeyCode = 13 Then

            btAggiungi_Click()

        End If

    End Sub

    Private Sub btAggiungi_Click(sender As Object, e As EventArgs) Handles btAggiungi.Click

        btAggiungi_Click()

    End Sub

    Private Sub btChiudi_Click(sender As Object, e As EventArgs) Handles btChiudi.Click

        btChiudi_Click()

    End Sub

End Class