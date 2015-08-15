'************************************************************************
'* Name:            debugenumerator.vb                                  *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Public Class debugEnumerator

    Public Enum EnumQueryType

        qtEvent
        qtFile

    End Enum

    Public Enum EnumCommand

        cmdNext
        cmdEnd
        cmdGoTo
        cmdBreak
        cmdBreakPoint
        cmdFree
        cmdNextSub

    End Enum

End Class
