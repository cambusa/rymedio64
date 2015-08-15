'************************************************************************
'* Name:            debugmonitor.vb                                     *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Imports ryMedio64.debugEnumerator
Imports ryMedio64.debugEnumerator.EnumQueryType

Public Class debugMonitor

    Dim WithEvents FrmDebug As frm_monitor
    Dim PropDirectory As String
    Dim PropQueryType As EnumQueryType
    Dim PropProcID As String
    Dim PropPolling As Long
    Dim PropModalMode As Boolean

    Public Sub New()

        On Error Resume Next

        PropQueryType = qtEvent
        PropPolling = 200
        PropModalMode = False

    End Sub

    Public Event WriteCommand(ByRef strCommand As String)
    Public Event ReadMonitor(ByRef strCode As String)
    Public Event ReadStatus(ByRef strStatus As String)
    Public Event StatusExists(ByRef ThereIs As Boolean)
    Public Event MonitorExists(ByRef ThereIs As Boolean)

    Public Sub DisplayMonitor()

        On Error Resume Next

        If Not FrmDebug Is Nothing Then

            FrmDebug.Close()
            FrmDebug = Nothing

        End If

        FrmDebug = New frm_monitor

        FrmDebug.Directory = Me.Directory
        FrmDebug.QueryType = Me.QueryType
        FrmDebug.ProcID = Me.ProcID
        FrmDebug.Polling = Me.Polling

        If PropModalMode Then
            FrmDebug.ShowDialog()
        Else
            FrmDebug.Show()
        End If

    End Sub

    Public Sub HideMonitor()

        On Error Resume Next

        If Not FrmDebug Is Nothing Then

            FrmDebug.Close()
            FrmDebug = Nothing

        End If

    End Sub

    Public Property Directory() As String

        Get

            On Error Resume Next

            If PropDirectory = "" Then

                PropDirectory = My.Application.Info.DirectoryPath + "\rymedio"

                If PropQueryType = qtFile Then

                    If Dir(PropDirectory, vbDirectory) = "" Then
                        MkDir(PropDirectory)
                    End If

                End If

            End If

            Return PropDirectory

        End Get

        Set(value As String)

            On Error Resume Next

            PropDirectory = value

            If Dir(PropDirectory, vbDirectory) = "" Then
                MkDir(PropDirectory)
            End If

        End Set

    End Property

    Public Property QueryType() As EnumQueryType

        Get

            Return PropQueryType

        End Get

        Set(value As EnumQueryType)

            PropQueryType = value

        End Set

    End Property

    Private Sub FrmDebug_ReadMonitor(ByRef strCode As String) Handles FrmDebug.ReadMonitor

        On Error Resume Next

        RaiseEvent ReadMonitor(strCode)

    End Sub

    Private Sub FrmDebug_ReadStatus(ByRef strStatus As String) Handles FrmDebug.ReadStatus

        On Error Resume Next

        RaiseEvent ReadStatus(strStatus)

    End Sub

    Private Sub FrmDebug_StatusExists(ByRef ThereIs As Boolean) Handles FrmDebug.StatusExists

        On Error Resume Next

        RaiseEvent StatusExists(ThereIs)

    End Sub

    Private Sub FrmDebug_MonitorExists(ByRef ThereIs As Boolean) Handles FrmDebug.MonitorExists

        On Error Resume Next

        RaiseEvent MonitorExists(ThereIs)

    End Sub

    Private Sub FrmDebug_WriteCommand(ByRef strCommand As String) Handles FrmDebug.WriteCommand

        On Error Resume Next

        RaiseEvent WriteCommand(strCommand)

    End Sub

    Public Property ProcID() As String

        Get

            On Error Resume Next

            If PropProcID = "" Then

                Randomize(Timer)
                PropProcID = Now.ToString("yyyymmddhhmmss") + Right("0000" & CLng(Rnd() * 1000), 4)

            End If

            Return PropProcID

        End Get

        Set(value As String)

            PropProcID = value

        End Set

    End Property

    Public Property Polling() As Long

        Get

            Return PropPolling

        End Get

        Set(value As Long)

            PropPolling = value

        End Set

    End Property

    Public Property ModalMode() As Boolean

        Get

            Return PropModalMode

        End Get

        Set(value As Boolean)

            PropModalMode = value

        End Set

    End Property

    Protected Overrides Sub Finalize()

        On Error Resume Next

        FrmDebug.Close()
        FrmDebug = Nothing

        MyBase.Finalize()

    End Sub

End Class
