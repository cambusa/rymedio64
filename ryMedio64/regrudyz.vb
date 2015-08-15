'************************************************************************
'* Name:            regrudyz.vb                                         *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Public Enum EnumRegHKey

    RegLocalMachine
    RegCurrentUser

End Enum

Module regrudyz

    Public Sub RegWriteSetting(TipoRegHKey As EnumRegHKey, Applicazione As String, Sezione As String, Variabile As String, Valore As String)

        Dim strPath As String
        Dim HKey As Microsoft.Win32.RegistryKey = Nothing

        Try

            strPath = "Software\RudyZ\"

            If Applicazione <> "" Then
                strPath = strPath + Applicazione + "\"
            End If

            If Sezione <> "" Then
                strPath = strPath + Sezione + "\"
            End If

            If TipoRegHKey = EnumRegHKey.RegCurrentUser Then
                HKey = My.Computer.Registry.CurrentUser.OpenSubKey(strPath, True)
                If HKey Is Nothing Then
                    HKey = My.Computer.Registry.CurrentUser.CreateSubKey(strPath, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                End If
            Else
                HKey = My.Computer.Registry.LocalMachine.OpenSubKey(strPath, True)
                If HKey Is Nothing Then
                    HKey = My.Computer.Registry.LocalMachine.CreateSubKey(strPath, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
                End If
            End If

            HKey.SetValue(Variabile, Valore)

        Finally

            If HKey IsNot Nothing Then
                HKey.Close()
            End If

        End Try

    End Sub

    Public Function RegReadSetting(TipoRegHKey As EnumRegHKey, Applicazione As String, Sezione As String, Variabile As String, DefaultValue As String) As String

        Dim strPath As String
        Dim HKey As Microsoft.Win32.RegistryKey = Nothing
        Dim value As String

        Try

            strPath = "Software\RudyZ\"

            If Applicazione <> "" Then
                strPath = strPath + Applicazione + "\"
            End If

            If Sezione <> "" Then
                strPath = strPath + Sezione + "\"
            End If

            If TipoRegHKey = EnumRegHKey.RegCurrentUser Then
                HKey = My.Computer.Registry.CurrentUser.OpenSubKey(strPath)
            Else
                HKey = My.Computer.Registry.LocalMachine.OpenSubKey(strPath)
            End If

            If HKey IsNot Nothing Then
                value = HKey.GetValue(Variabile).ToString()
                HKey.Close()
            Else
                value = DefaultValue
            End If

        Catch ex As Exception

            value = DefaultValue

        Finally

            If HKey IsNot Nothing Then
                HKey.Close()
            End If

        End Try

        Return value

    End Function

End Module
