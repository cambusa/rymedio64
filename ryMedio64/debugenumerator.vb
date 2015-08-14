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
