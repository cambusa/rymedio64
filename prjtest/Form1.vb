Option Explicit On

Imports ryMedio64
Imports System.Xml
Imports ryMedio64.debugEnumerator.EnumQueryType

Public Class Form1

    Dim WithEvents ObjDebug As debugParser

    Dim CL As Collection
    Dim ObjXml As XmlDocument

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ObjDebug = New debugParser
        ObjDebug.QueryType = qtEvent

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim r As debugMonitor = New debugMonitor

        Dim ObjMonitor As debugMonitor

        Dim C1 As Collection
        Dim C2 As Collection
        Dim C3 As Collection

        C1 = New Collection
        C2 = New Collection
        C3 = New Collection

        C1.Add("Pippo")

        C2.Add("Agenore")
        C2.Add("Flangio")

        CL = New Collection

        CL.Add(C1, "P")
        CL.Add(C2, "Q")
        CL.Add(C3, "R")

        ObjXml = New XmlDocument
        ObjXml.LoadXml("<xml tipo='spiaggia'><param/><dati nome='pippo' cognome='canino' /></xml>")


        ObjMonitor = New debugMonitor
        ObjMonitor.QueryType = qtEvent
        ObjMonitor.Polling = 200

        ObjDebug.ProcID = ObjMonitor.ProcID
        ObjDebug.Monitor = ObjMonitor

        ObjMonitor.DisplayMonitor()

        ObjDebug.SourceFile(My.Application.Info.DirectoryPath + "\prova.vbs")
        ObjDebug.Execute("Main")

        'ObjMonitor.HideMonitor()
        'ObjMonitor = Nothing

        CL = Nothing

        C1 = Nothing
        C2 = Nothing
        C3 = Nothing

    End Sub

    Private Sub ObjDebug_InitScript(ScriptObject As Object) Handles ObjDebug.InitScript

        'ScriptObject.addObject("Frm", Me.o)
        ScriptObject.addObject("CL", CL)
        'ScriptObject.addObject("DXML", ObjXml)

    End Sub

End Class
