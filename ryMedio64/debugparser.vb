'************************************************************************
'* Name:            debugparser.vb                                     *
'* Project:         ryMedio                                             *
'* Version:         2.0                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit On

Imports System.Text.RegularExpressions
Imports MSScriptControl
Imports ryMedio64.debugEnumerator
Imports ryMedio64.debugEnumerator.EnumQueryType
Imports ryMedio64.debugEnumerator.EnumCommand
Imports ryMedio64.moduleDebug.EnumTipoRiga

Public Class debugParser

    Dim PropSource As String
    Dim PropDirectory As String
    Dim PropQueryType As EnumQueryType
    Dim PropTimeout As Long
    Dim PropPolling As Long
    Dim PropProcID As String

    Dim WithEvents PropMonitor As debugMonitor

    Dim VettRighe() As StructRiga
    Dim VettEval() As StructValutazione
    Dim VettFunct() As StructFunction

    Public Event ScriptRequest(ScriptObject As Object)
    Public Event InitScript(ScriptObject As Object)
    Public Event EndExecute(LastCommand As EnumCommand, ErrNumber As Long, ErrDescr As String)

    Dim PropScript As ScriptControl

    Dim FlagTermine As Boolean

    Dim QueryCommand As String
    Dim QueryMonitor As String
    Dim QueryStatus As String

    Dim PollingSecondi As Single

    Dim SerialeID As Long

    Dim ElencoBreakPoint As String

    Public Sub New()

        On Error Resume Next

        PropDirectory = ""
        PropQueryType = qtEvent
        PropTimeout = 600
        Me.Polling = 200

    End Sub

    Public Property Source() As String

        Get
            On Error Resume Next
            Source = PropSource
        End Get

        Set(value As String)

            On Error Resume Next

            PropSource = value

            AnalizzaSorgente(Directory, PropSource, VettRighe, VettEval, VettFunct)

        End Set

    End Property

    Public Function SourceFile(FileName As String) As Boolean

        On Error GoTo ErrorProcedure

        Me.Source = My.Computer.FileSystem.ReadAllText(FileName)

        Return True

        Exit Function

ErrorProcedure:

        GlobalErrDescription = Err.Description
        GlobalErrNumber = Err.Number

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        PropSource = ""
        SourceFile = False

    End Function

    Public Property LastErrDescription() As String

        Get

            Return GlobalErrDescription

        End Get

        Set(value As String)

            GlobalErrDescription = value

        End Set

    End Property

    Public Property LastErrNumber() As Long

        Get

            Return GlobalErrNumber

        End Get

        Set(value As Long)

            GlobalErrNumber = value

        End Set

    End Property

    Public Function Execute(ByVal SubName As String) As Object

        On Error GoTo ErrorProcedure

        Dim Ritorno As Object = String.Empty
        Dim Comando As EnumCommand
        Dim FlagNewScript As Boolean
        Dim OnErrorRaise As Boolean
        Dim ObjScriptAux As ScriptControl

        ScriviMonitor()

        RaiseEvent ScriptRequest(PropScript)

        If PropScript Is Nothing Then

            PropScript = New ScriptControl
            PropScript.Language = "vbscript"
            PropScript.Timeout = -1

            On Error GoTo GestioneScript
            PropScript.AddCode(PropSource)
            On Error GoTo ErrorProcedure

            FlagNewScript = True

        Else

            ObjScriptAux = New ScriptControl
            ObjScriptAux.Language = "vbscript"
            ObjScriptAux.Timeout = -1

            On Error GoTo GestioneScriptAux
            ObjScriptAux.AddCode(PropSource)
            On Error GoTo ErrorProcedure

            ObjScriptAux = Nothing

            FlagNewScript = False

        End If

        RaiseEvent InitScript(PropScript)

        Comando = cmdNext
        ElencoBreakPoint = ""
        OnErrorRaise = True

        Ritorno = InvocaFunzione(SubName, OnErrorRaise, Comando, True)

        EliminaMonitor()
        ScriviTermine()

        RaiseEvent EndExecute(Comando, 0, "")

        If FlagNewScript Then
            PropScript = Nothing
        End If

        Execute = Ritorno

        Exit Function

GestioneScript:

        GlobalErrDescription = PropScript.Error.Description
        GlobalErrNumber = PropScript.Error.Number

        Comando = cmdNext
        AttesaComando(True, (PropScript.Error.Line), "", Comando, -1)

        Resume AbortProcedure

GestioneScriptAux:

        GlobalErrDescription = PropScript.Error.Description
        GlobalErrNumber = PropScript.Error.Number

        Comando = cmdNext
        AttesaComando(True, (ObjScriptAux.Error.Line), "", Comando, -1)

        Resume AbortProcedure

ErrorProcedure:

        GlobalErrDescription = Err.Description
        GlobalErrNumber = Err.Number

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        ObjScriptAux = Nothing

        EliminaMonitor()
        ScriviTermine()

        RaiseEvent EndExecute(Comando, GlobalErrNumber, GlobalErrDescription)

        If FlagNewScript Then
            PropScript = Nothing
        End If

        Return String.Empty

    End Function

    Private Function InvocaFunzione(ByVal SubName As String, ByRef OnErrorRaise As Boolean, ByRef Comando As EnumCommand, FlagPrimary As Boolean) As Object

        On Error GoTo ErrorProcedure

        Dim Ritorno As Object = String.Empty
        Dim R As Long
        Dim I As Long
        Dim J As Long
        Dim IndCase As Long
        Dim RigaInizio As Long
        Dim RigaFine As Long
        Dim StrVal1 As Object
        Dim StrVal2 As Object
        Dim StrVal3 As Object
        Dim RigaSuccessiva As Long
        Dim FlagNewScript As Boolean
        Dim CurrCommento As String
        Dim ObjTmp As Object

        RigaInizio = EntryPoint(SubName, "")

        If RigaInizio > 0 Then

            R = RigaInizio
            RigaFine = VettRighe(R).Sorella
            FlagTermine = False
            QueryCommand = ""
            CurrCommento = ""

            Do

                Select Case Comando

                    Case cmdGoTo

                        R = RigaSuccessiva
                        Comando = cmdNext

                    Case cmdFree

                        If R = RigaFine Then
                            Comando = cmdNext
                        ElseIf InStr(ElencoBreakPoint, "|" & R & "|") > 0 Then
                            Comando = cmdNext
                        End If

                    Case cmdBreakPoint

                        If R = RigaSuccessiva Then
                            Comando = cmdNext
                        End If

                    Case cmdEnd

                        Exit Do

                End Select

                Valutazioni()

                Select Case VettRighe(R).Tipo

                    Case trStatement, trEndIf, trEndSelect, trEndWith, trIfThen, trIfThenElse

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).Espr1 <> "" Then

                                If FlagPrimary Then
                                    SerialeID = 0
                                End If

                                EseguiIstruzione(VettRighe(R).Espr1, R, Comando, OnErrorRaise, SubName, Ritorno)

                            End If

                            R = IstruzioneSuccessiva(R)

                        End If

                    Case trResumeNext, trGoToZero

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            OnErrorRaise = (VettRighe(R).Tipo = trGoToZero)

                            R = IstruzioneSuccessiva(R)

                        End If

                    Case trDim, trInside

                        R = IstruzioneSuccessiva(R)

                    Case trGoTo

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            StrVal1 = PropScript.Eval(VettRighe(R).Espr1)

                            If VarType(StrVal1) = vbString Then
                                R = Val(StrVal1)
                            Else
                                R = CLng(StrVal1)
                            End If

                        End If

                    Case trStop

                        Comando = cmdNext

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            R = IstruzioneSuccessiva(R)

                        End If

                    Case trRem

                        CurrCommento = VettRighe(R).Espr1

                        R = IstruzioneSuccessiva(R)

                    Case trEnd

                        Comando = cmdNext

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            R = 0

                        End If

                    Case trIf, trElseIf

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            StrVal1 = ValutaEspressione(VettRighe(R).Espr1, R, Comando, OnErrorRaise, False)

                            If StrVal1 Then

                                '----------------------
                                ' La condizione è vera
                                '----------------------

                                If VettRighe(R).Figlia > 0 Then

                                    '------------------------------------------
                                    ' L'istruzione figlia esiste: mi ci sposto
                                    '------------------------------------------

                                    R = VettRighe(R).Figlia

                                Else

                                    '-------------------------------------
                                    ' Il blocco è vuoto: vado alla End If
                                    '-------------------------------------

                                    Do
                                        R = VettRighe(R).Sorella
                                    Loop Until VettRighe(R).Tipo = trEndIf

                                End If

                            Else

                                '---------------------------------------------------
                                ' La condizione è falsa: mi sposto sulla successiva
                                '---------------------------------------------------

                                R = VettRighe(R).Sorella

                            End If

                        End If

                    Case trElse, trFunct, trSub

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).Figlia > 0 Then
                                R = VettRighe(R).Figlia
                            Else
                                R = VettRighe(R).Sorella
                            End If

                        End If

                    Case trForEach

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If FlagPrimary Then
                                SerialeID = 0
                            End If

                            StrVal3 = True

                            ObjTmp = PropScript.Eval(VettRighe(R).Espr2)

                            If VettRighe(R).FlagCiclo = False Then

                                VettRighe(R).Index = 1

                                If RisolviOggetto(ObjTmp, VettRighe(R).Index) Is Nothing Then

                                    StrVal3 = False

                                End If

                                VettRighe(R).FlagCiclo = True

                            Else

                                VettRighe(R).Index = VettRighe(R).Index + 1

                                If RisolviOggetto(ObjTmp, VettRighe(R).Index) Is Nothing Then

                                    StrVal3 = False

                                End If

                            End If

                            If StrVal3 Then

                                StrVal1 = "tmp___foreach" & R

                                EseguiIstruzione(StrVal1 & "=0:For Each " & VettRighe(R).Nome & " In " & VettRighe(R).Espr2 & ":" & StrVal1 & "=" & StrVal1 & "+1:If " & StrVal1 & "=" & VettRighe(R).Index & " Then:Exit For:End If:Next", R, Comando, OnErrorRaise)

                                VettRighe(VettRighe(R).Sorella).FlagCiclo = True

                                If VettRighe(R).Figlia > 0 Then
                                    R = VettRighe(R).Figlia
                                Else
                                    R = VettRighe(R).Sorella
                                End If

                            Else

                                VettRighe(R).FlagCiclo = False

                                R = VettRighe(R).Sorella
                                VettRighe(R).FlagCiclo = False

                            End If

                            ObjTmp = Nothing

                        End If

                    Case trFor, trForStep

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).FlagCiclo = False Then

                                If FlagPrimary Then
                                    SerialeID = 0
                                End If

                                EseguiIstruzione(VettRighe(R).Espr1, R, Comando, OnErrorRaise)
                                VettRighe(R).FlagCiclo = True
                                VettRighe(R).Confronto = PropScript.Eval(VettRighe(R).Espr2)

                            Else

                                If FlagPrimary Then
                                    SerialeID = 0
                                End If

                                EseguiIstruzione(VettRighe(R).Nome + "=" + VettRighe(R).Nome + "+(" + VettRighe(R).Espr3 + ")", R, Comando, OnErrorRaise)

                            End If

                            StrVal1 = PropScript.Eval(VettRighe(R).Nome)
                            StrVal2 = VettRighe(R).Confronto

                            If PropScript.Eval(VettRighe(R).Espr3) > 0 Then
                                StrVal3 = (StrVal1 <= StrVal2)
                            Else
                                StrVal3 = (StrVal2 <= StrVal1)
                            End If

                            If StrVal3 Then

                                VettRighe(VettRighe(R).Sorella).FlagCiclo = True

                                If VettRighe(R).Figlia > 0 Then
                                    R = VettRighe(R).Figlia
                                Else
                                    R = VettRighe(R).Sorella
                                End If

                            Else

                                VettRighe(R).FlagCiclo = False

                                R = VettRighe(R).Sorella
                                VettRighe(R).FlagCiclo = False

                            End If

                        End If

                    Case trExitFor

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            Do
                                R = VettRighe(R).Madre
                            Loop Until (VettRighe(R).Tipo = trFor Or VettRighe(R).Tipo = trForStep Or VettRighe(R).Tipo = trForEach)

                            VettRighe(R).FlagCiclo = False

                            R = VettRighe(R).Sorella
                            VettRighe(R).FlagCiclo = False

                        End If

                    Case trNext

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).FlagCiclo Then
                                R = VettRighe(R).Inizio
                            Else
                                R = IstruzioneSuccessiva(R)
                            End If

                        End If

                    Case trDo, trDoUntil, trDoWhile

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            StrVal1 = ValutaEspressione(VettRighe(R).Espr1, R, Comando, OnErrorRaise, False)

                            If StrVal1 Then

                                '------------------------------------------------------------------
                                ' Nel blocco DO LOOP, solo FlagCiclo sull'istruzione finale (LOOP)
                                ' ha senso. Si noti che LOOP e la sorella di DO.
                                '------------------------------------------------------------------

                                ' Informo LOOP che devo rimanere nel ciclo
                                VettRighe(VettRighe(R).Sorella).FlagCiclo = True

                                If VettRighe(R).Figlia > 0 Then
                                    R = VettRighe(R).Figlia
                                Else
                                    R = VettRighe(R).Sorella
                                End If

                            Else

                                R = VettRighe(R).Sorella
                                VettRighe(R).FlagCiclo = False

                            End If

                        End If

                    Case trExitDo

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            Do
                                R = VettRighe(R).Madre
                            Loop Until (VettRighe(R).Tipo = trDo Or VettRighe(R).Tipo = trDoUntil Or VettRighe(R).Tipo = trDoWhile)

                            R = VettRighe(R).Sorella
                            VettRighe(R).FlagCiclo = False

                        End If

                    Case trLoop, trLoopUntil, trLoopWhile

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).FlagCiclo Then

                                StrVal1 = ValutaEspressione(VettRighe(R).Espr1, R, Comando, OnErrorRaise, False)

                                If StrVal1 Then
                                    R = VettRighe(R).Inizio
                                Else
                                    R = IstruzioneSuccessiva(R)
                                End If

                            Else

                                R = IstruzioneSuccessiva(R)

                            End If

                        End If

                    Case trSelect

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            StrVal1 = ValutaEspressione(VettRighe(R).Espr1, R, Comando, OnErrorRaise, String.Empty)

                            IndCase = 0
                            I = VettRighe(R).Figlia

                            Do While I > 0

                                If VettRighe(I).Tipo = trCase Then

                                    For J = 1 To UBound(VettRighe(I).CaseEspr)

                                        StrVal2 = PropScript.Eval(VettRighe(I).CaseEspr(J))

                                        If StrVal1 = StrVal2 Then

                                            IndCase = I
                                            Exit Do

                                        End If

                                    Next J

                                ElseIf VettRighe(I).Tipo = trCaseElse Then

                                    IndCase = I
                                    Exit Do

                                Else

                                    Exit Do

                                End If

                                I = VettRighe(I).Sorella

                            Loop

                            If IndCase > 0 Then
                                R = IndCase
                            Else
                                R = VettRighe(R).Sorella
                            End If

                        End If

                    Case trWith, trCase, trCaseElse

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            If VettRighe(R).Figlia > 0 Then

                                R = VettRighe(R).Figlia

                            Else

                                '----------------------------------------------------------
                                ' Case vuoto: mi posizione sulla madre e poi sulla sorella
                                '----------------------------------------------------------

                                R = VettRighe(R).Madre
                                R = VettRighe(R).Sorella

                            End If

                        End If

                    Case trEndWith

                        If AttesaComando(OnErrorRaise, R, CurrCommento, Comando, RigaSuccessiva) Then

                            R = IstruzioneSuccessiva(R)

                        End If

                    Case trEndFunct, trEndSub, trExitFunction, trExitSub

                        If FlagPrimary = True Or Comando = cmdNext Or Comando = cmdNextSub Then

                            Comando = cmdNext
                            AttesaComando(OnErrorRaise, R, "", Comando, -1)

                        End If

                        R = 0

                    Case Else

                        R = 0

                End Select

                If R = 0 Then
                    Exit Do
                End If

            Loop

        Else

            GlobalErrDescription = "La subroutine non esiste"
            GlobalErrNumber = 5

            GoTo AbortProcedure

        End If

        Return Ritorno

GestioneScript:

        GlobalErrDescription = PropScript.Error.Description
        GlobalErrNumber = PropScript.Error.Number

        'ScriviHTML Me.Directory + "\~debug.htm", VettRighe(), PropScript.Error.Line, GlobalErrDescription

        Resume AbortProcedure

ErrorProcedure:

        GlobalErrDescription = Err.Description
        GlobalErrNumber = Err.Number

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        ObjTmp = Nothing

        Comando = cmdNext
        AttesaComando(False, R, GlobalErrDescription, Comando, -1)

        Return String.Empty

    End Function

    Private Function AttesaComando(OnErrorRaise As Boolean, Riga As Long, Commento As String, ByRef Comando As EnumCommand, byref RigaSuccessiva As Long) As Boolean

        On Error GoTo ErrorProcedure

        Dim FileStatoTmp As String
        Dim FileStato As String = ""
        Dim FileComando As String = ""
        Dim R As Long
        Dim Buffer As String = ""
        Dim NumFile As Integer
        Dim PosParola As Long
        Dim Param As String
        Dim Cmd As String
        Dim MessErr As String
        Dim Tempo As Single
        Dim Esiste As Boolean
        Dim MaxEval As Long

        If Comando = cmdNext Or Comando = cmdNextSub Then

RiprendiAttesa:

            MessErr = ""

            If PropScript.Error.Number <> 0 Then

                If OnErrorRaise Then
                    MessErr = PropScript.Error.Description
                End If

                'PropScript.Error.Clear

            End If

            If PropQueryType = qtFile Then

                FileStatoTmp = Me.Directory + "\" + Me.ProcID + ".TMP"
                FileStato = Me.Directory + "\" + Me.ProcID + ".STA"

                NumFile = FreeFile()
                FileOpen(NumFile, FileStatoTmp, OpenMode.Output)

                Print(NumFile, MessErr)
                Print(NumFile, Commento)

                If MessErr <> "" Then
                    If OnErrorRaise Then
                        Print(NumFile, 1)
                    Else
                        Print(NumFile, 0)
                    End If
                Else
                    If RigaSuccessiva = -1 Then
                        Print(NumFile, 1)
                    Else
                        Print(NumFile, 0)
                    End If
                End If

                Print(NumFile, Riga)

                For R = 1 To UBound(VettEval)

                    Buffer = VettEval(R).Valore
                    Print(NumFile, Buffer)

                Next R

                FileClose(NumFile)
                NumFile = 0

                If Dir(FileStato) <> "" Then
                    My.Computer.FileSystem.DeleteFile(FileStato)
                End If

                My.Computer.FileSystem.RenameFile(FileStatoTmp, FileStato)

            Else

                Buffer = ""

                Buffer = Buffer & MessErr & vbCrLf
                Buffer = Buffer & Commento & vbCrLf

                If MessErr <> "" Then
                    If OnErrorRaise Then
                        Buffer = Buffer & "1" & vbCrLf
                    Else
                        Buffer = Buffer & "0" & vbCrLf
                    End If
                Else
                    If RigaSuccessiva = -1 Then
                        Buffer = Buffer & "1" & vbCrLf
                    Else
                        Buffer = Buffer & "0" & vbCrLf
                    End If
                End If

                Buffer = Buffer & Riga & vbCrLf

                For R = 1 To UBound(VettEval)

                    Buffer = Buffer & VettEval(R).Valore & vbCrLf

                Next R

                QueryStatus = Buffer

            End If

            '------------------
            ' Gestione comando
            '------------------

            Tempo = Timer

            If PropQueryType = qtFile Then

                FileComando = Me.Directory + "\" + Me.ProcID + ".CMN"

                Do

                    PollingAttesa(PollingSecondi)

                    If Dir(FileComando) <> "" Then

                        Exit Do

                    ElseIf FlagTermine Then

                        Exit Do

                    ElseIf PropTimeout > 0 Then

                        If Math.Abs(Timer - Tempo) > PropTimeout Then

                            FlagTermine = True
                            Exit Do

                        End If

                    End If

                Loop

            Else

                Do

                    PollingAttesa(PollingSecondi)

                    If QueryCommand <> "" Then

                        Exit Do

                    ElseIf FlagTermine Then

                        Exit Do

                    ElseIf PropTimeout > 0 Then

                        If Math.Abs(Timer - Tempo) > PropTimeout Then

                            FlagTermine = True
                            Exit Do

                        End If

                    End If

                Loop

            End If

            If FlagTermine Then

                Comando = cmdEnd

            Else

                If PropQueryType = qtFile Then

                    On Error Resume Next
                    My.Computer.FileSystem.DeleteFile(FileStato)
                    On Error GoTo ErrorProcedure

                    Dim Lines() As String = IO.File.ReadAllLines(FileComando)
                    Buffer = Lines(0)

                    'NumFile = FreeFile()
                    'FileOpen(NumFile, FileComando, OpenMode.Input)
                    'Line Input #NumFile, Buffer
                    ' FileClose(NumFile)
                    'NumFile = 0

                    On Error Resume Next
                    My.Computer.FileSystem.DeleteFile(FileComando)
                    On Error GoTo ErrorProcedure

                Else

                    Buffer = QueryCommand
                    QueryCommand = ""

                End If

                PosParola = InStr(Buffer, ":")

                If PosParola > 0 Then
                    Param = Mid(Buffer, PosParola + 1)
                    Cmd = Left(Buffer, PosParola - 1)
                Else
                    Param = ""
                    Cmd = Buffer
                End If

                RigaSuccessiva = 0

                Select Case Cmd

                    Case "end"

                        Comando = cmdEnd

                    Case "goto"

                        Comando = cmdGoTo
                        RigaSuccessiva = Val(Param)

                    Case "break"

                        Comando = cmdBreakPoint
                        RigaSuccessiva = Val(Param)

                        If RigaSuccessiva = 0 Then
                            Comando = cmdNext
                        End If

                    Case "free"

                        Comando = cmdFree
                        ElencoBreakPoint = Param

                    Case "watch"

                        Esiste = False

                        For R = 1 To UBound(VettEval)

                            If LCase(VettEval(R).Espressione) = LCase(Param) Then

                                Esiste = True

                                Exit For

                            End If

                        Next R

                        If Esiste = False Then

                            MaxEval = UBound(VettEval) + 1

                            ReDim Preserve VettEval(MaxEval)

                            VettEval(MaxEval).Espressione = Param
                            VettEval(MaxEval).Valore = String.Empty

                            Valutazioni()

                        End If

                        If PropQueryType = qtFile Then

                            On Error Resume Next
                            My.Computer.FileSystem.DeleteFile(FileComando)
                            On Error GoTo ErrorProcedure

                        Else

                            QueryCommand = ""

                        End If

                        GoTo RiprendiAttesa

                    Case Else

                        If Val(Param) = 0 Then
                            Comando = cmdNext
                        Else
                            Comando = cmdNextSub
                        End If

                End Select

            End If

        ElseIf Comando = cmdFree Then

            System.Windows.Forms.Application.DoEvents()

            If PropQueryType = qtFile Then

                FileComando = Me.Directory + "\" + Me.ProcID + ".CMN"

                If Dir(FileComando) <> "" Then

                    Dim Lines() As String = IO.File.ReadAllLines(FileComando)
                    Buffer = Lines(0)

                    'NumFile = FreeFile()
                    'FileOpen(NumFile, FileComando, OpenMode.Input)
                    'Line Input #NumFile, Buffer
                    ' FileClose(NumFile)
                    'NumFile = 0

                    On Error Resume Next
                    My.Computer.FileSystem.DeleteFile(FileComando)
                    On Error GoTo ErrorProcedure

                End If

            Else

                If QueryCommand <> "" Then

                    Buffer = QueryCommand
                    QueryCommand = ""

                End If

            End If

            If Buffer = "break:0" Then

                Comando = cmdNext

            End If

        End If

        AttesaComando = (Comando <> cmdEnd) And (Comando <> cmdGoTo)

        Exit Function

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        If NumFile > 0 Then
            FileClose(NumFile)
            NumFile = 0
        End If

        Return False

    End Function

    Private Function IstruzioneSuccessiva(R As Long)

        On Error GoTo ErrorProcedure

        Dim NextR As Long

        NextR = R

        If VettRighe(NextR).Sorella > 0 Then

            '-----------------------------------------------
            ' La sorella esiste: mi posiziono sulla sorella
            '-----------------------------------------------

            NextR = VettRighe(NextR).Sorella

        Else

            '-----------------------
            ' La sorella non esiste
            '-----------------------

            If VettRighe(NextR).Madre > 0 Then

                '-------------------------------------------
                ' La madre esiste: mi posiziono sulla madre
                '-------------------------------------------

                NextR = VettRighe(NextR).Madre

                Select Case VettRighe(NextR).Tipo

                    Case trCase, trCaseElse

                        '---------------------------
                        ' Mi posiziono sulla SELECT
                        '---------------------------

                        NextR = VettRighe(NextR).Madre

                        '-------------------------------
                        ' Mi posiziono sulla END SELECT
                        '-------------------------------

                        NextR = VettRighe(NextR).Sorella

                    Case trIf, trElseIf, trElse

                        '------------------------
                        ' Avanzo fino all'END IF
                        '------------------------

                        Do
                            NextR = VettRighe(NextR).Sorella
                        Loop Until VettRighe(NextR).Tipo = trEndIf

                    Case trWith, trFor, trForStep, trForEach, trDo, trDoUntil, trDoWhile

                        '------------------------
                        ' Mi posiziono sulla END
                        '------------------------

                        NextR = VettRighe(NextR).Sorella

                    Case trFunct, trSub

                        '------------------------
                        ' Mi posiziono sulla END
                        '------------------------

                        NextR = VettRighe(NextR).Sorella

                    Case Else

                        NextR = 0

                End Select

            Else

                '--------------------------------------
                ' La madre non esiste: fine esecuzione
                '--------------------------------------

                NextR = 0

            End If

        End If

        Return NextR

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        Return 0

    End Function

    Private Sub Valutazioni()

        On Error GoTo ErrorProcedure

        Dim I As Long

        For I = 1 To UBound(VettEval)

            VettEval(I).Valore = EseguiValutazione(VettEval(I).Espressione, VettEval(I).FlagInside)

        Next I

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

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

    Private Sub ScriviMonitor()

        On Error GoTo ErrorProcedure

        Dim PathFile As String
        Dim R As Long
        Dim Buffer As String
        Dim NumFile As Integer
        Dim LenPref As Long

        LenPref = Len("" & UBound(VettRighe))

        If PropQueryType = qtFile Then

            PathFile = Me.Directory + "\" + Me.ProcID + ".MON"
            NumFile = FreeFile()
            FileOpen(NumFile, PathFile, OpenMode.Output)

            For R = 1 To UBound(VettRighe)

                Buffer = VettRighe(R).Orig
                Print(NumFile, Buffer)

            Next R

            FileClose(NumFile)
            NumFile = 0

        Else

            Buffer = ""

            For R = 1 To UBound(VettRighe)

                Buffer = Buffer + VettRighe(R).Orig + vbCrLf

            Next R

            QueryMonitor = Buffer

        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        If NumFile > 0 Then
            FileClose(NumFile)
            NumFile = 0
        End If

    End Sub

    Private Sub EseguiIstruzione(ByVal Istruzione As String, R As Long, ByRef Comando As EnumCommand, ByRef OnErrorRaise As Boolean, Optional ByRef SubName As String = "", Optional ByRef Ritorno As Object = Nothing)

        On Error GoTo ErrorProcedure

        Dim Commento As String
        Dim PosUguale As Long
        Dim SalvaOnError As Boolean

        If StrComp(Left(Replace(Istruzione, " ", ""), Len(SubName) + 1), SubName + "=", vbTextCompare) = 0 Then

            PosUguale = InStr(Istruzione, "=")

            SerialeID = SerialeID + 1
            Ritorno = "tmp___" & SerialeID

            SalvaOnError = OnErrorRaise
            AnalizzaEspressione(Ritorno + " = " + Mid(Istruzione, PosUguale + 1), OnErrorRaise, Comando, Commento)
            OnErrorRaise = SalvaOnError

            If Commento <> "" Then

                Comando = cmdNext

                GoTo AbortProcedure

            ElseIf Comando = cmdNextSub Then

                Comando = cmdNext

            End If

        ElseIf StrComp(Left(Replace(Istruzione, " ", ""), Len(SubName) + 1), "Set" + SubName + "=", vbTextCompare) = 0 Then

            PosUguale = InStr(Istruzione, "=")

            SerialeID = SerialeID + 1
            Ritorno = "tmp___" & SerialeID

            SalvaOnError = OnErrorRaise
            AnalizzaEspressione("Set " + Ritorno + " = " + Mid(Istruzione, PosUguale + 1), OnErrorRaise, Comando, Commento)
            OnErrorRaise = SalvaOnError

            If Commento <> "" Then

                Comando = cmdNext

                GoTo AbortProcedure

            ElseIf Comando = cmdNextSub Then

                Comando = cmdNext

            End If

        Else

            SalvaOnError = OnErrorRaise
            AnalizzaEspressione(Istruzione, OnErrorRaise, Comando, Commento)
            OnErrorRaise = SalvaOnError

            If Commento <> "" Then

                Comando = cmdNext

                GoTo AbortProcedure

            ElseIf Comando = cmdNextSub Then

                Comando = cmdNext

            End If

        End If

        Exit Sub

ErrorProcedure:

        Commento = Err.Description

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        If OnErrorRaise Then
            Comando = cmdNext
        End If

        AttesaComando(OnErrorRaise, R, Commento, Comando, -1)

        If OnErrorRaise Then

            R = 0

        End If

    End Sub

    Public Property ObjScript() As Object

        Get

            On Error Resume Next

            If PropScript Is Nothing Then

                PropScript = New ScriptControl
                PropScript.Language = "vbscript"
                PropScript.Timeout = -1

            End If

            Return PropScript

        End Get

        Set(value As Object)

            PropScript = Nothing
            PropScript = value

        End Set

    End Property

    Public Property Monitor() As debugMonitor

        Get

            Return PropMonitor

        End Get

        Set(value As debugMonitor)

            PropMonitor = value

        End Set

    End Property

    Private Sub PropMonitor_ReadMonitor(ByRef strCode As String) Handles PropMonitor.ReadMonitor

        On Error Resume Next

        strCode = QueryMonitor
        QueryMonitor = ""

    End Sub

    Private Sub PropMonitor_ReadStatus(ByRef strStatus As String) Handles PropMonitor.ReadStatus

        On Error Resume Next

        strStatus = QueryStatus
        QueryStatus = ""

    End Sub

    Private Sub PropMonitor_StatusExists(ByRef ThereIs As Boolean) Handles PropMonitor.StatusExists

        On Error Resume Next

        ThereIs = (QueryStatus <> "")

    End Sub

    Private Sub PropMonitor_MonitorExists(ByRef ThereIs As Boolean) Handles PropMonitor.MonitorExists

        On Error Resume Next

        ThereIs = (QueryMonitor <> "")

    End Sub

    Private Sub PropMonitor_WriteCommand(ByRef strCommand As String) Handles PropMonitor.WriteCommand

        On Error Resume Next

        QueryCommand = strCommand

    End Sub

    Public Property QueryType() As EnumQueryType

        Get

            Return PropQueryType

        End Get

        Set(NewValue As EnumQueryType)

            PropQueryType = NewValue

        End Set

    End Property

    Private Function EseguiValutazione(Istruzione As String, FlagInside As Boolean) As Object

        On Error GoTo ErrorProcedure

        Dim Valore As Object = ""
        Dim I As Long
        Dim V As Object
        Dim Obj As Object
        Dim Prop As Object

        Obj = PropScript.Eval(Istruzione)

        Select Case VarType(Obj)

            Case VariantType.String

                Valore = CStr(Obj).Replace(vbCr, " ").Replace(vbLf, " ")

                If Len(Valore) > 1000 Then
                    Valore = Left(Valore, 1000) + "..."
                End If

                Valore = Istruzione + " = " + Chr(34) + Valore + Chr(34)

            Case VariantType.Date

                Valore = Istruzione + " = " + CDate(Obj).ToString("dd/mm/yyyy hh:mm:ss")

            Case VariantType.Empty

                Valore = Istruzione + " = [Empty]"

            Case VariantType.Null

                Valore = Istruzione + " = [Null]"

            Case VariantType.Array, 8204, 8201

                If FlagInside Then

                    V = Obj

                    For I = LBound(V) To UBound(V)
                        V(I) = EseguiValutazione(Istruzione & "(" & I & ")", False)
                    Next

                    Valore = Join(V, vbCrLf)

                Else

                    Valore = Istruzione + " = [Array]"

                End If

            Case VariantType.Object

                If Obj Is Nothing Then

                    Valore = Istruzione + " = [Nothing]"

                Else

                    If FlagInside Then

                        ElencoProp(Obj, Istruzione, Valore)

                    Else

                        Valore = Istruzione + " = [" + TypeName(Obj) + "]"

                    End If

                End If

            Case Else

                Valore = Istruzione + " = " + Str(Obj)

        End Select

        Return Valore

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        Return String.Empty

    End Function

    Public Property Timeout() As Long

        Get

            Return PropTimeout

        End Get

        Set(value As Long)

            PropTimeout = value

        End Set

    End Property

    Public Property Polling() As Long

        Get

            Return PropPolling

        End Get

        Set(value As Long)

            PropPolling = value
            PollingSecondi = value / 1000

        End Set

    End Property

    Private Sub ScriviTermine()

        On Error GoTo ErrorProcedure

        Dim FileStatoTmp As String
        Dim FileStato As String
        Dim NumFile As Integer
        Dim MessEnd As String

        MessEnd = "end"

        If PropQueryType = qtFile Then

            FileStatoTmp = Me.Directory + "\" + Me.ProcID + ".TMP"
            FileStato = Me.Directory + "\" + Me.ProcID + ".STA"

            NumFile = FreeFile()
            FileOpen(NumFile, FileStatoTmp, OpenMode.Output)
            Print(NumFile, MessEnd)
            FileClose(NumFile)
            NumFile = 0

            If Dir(FileStato) <> "" Then
                My.Computer.FileSystem.DeleteFile(FileStato)
            End If

            My.Computer.FileSystem.RenameFile(FileStatoTmp, FileStato)

        Else

            QueryStatus = MessEnd & vbCrLf

        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        If NumFile > 0 Then
            FileClose(NumFile)
            NumFile = 0
        End If

    End Sub

    Private Sub EliminaMonitor()

        On Error Resume Next

        Dim PathFile As String

        If PropQueryType = qtFile Then

            PathFile = Me.Directory + "\" + Me.ProcID + ".MON"

            My.Computer.FileSystem.DeleteFile(PathFile)

        End If

    End Sub

    Public Property ProcID() As String

        Get

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

    Private Sub ElencoProp(ByRef Obj As Object, Nome As String, ByRef Valori As Object)

        On Error GoTo ErrorProcedure

        Dim I As Long
        Dim strVal As String
        Dim Elenco As String
        Dim PropName As String
        Dim Cnt As Long
        Dim Ind As Object
        Dim VarAux As Object

        Valori = ""
        Elenco = "|"

        Select Case LCase(TypeName(Obj))

            Case "recordset"

                '------------------------------
                ' Espongo il nome dell'oggetto
                '------------------------------

                Valori = Nome + " = [" + TypeName(Obj) + "]" + vbCrLf

                '--------------------------------------
                ' Espongo i valori delle sue proprietà
                '--------------------------------------

                Cnt = -1

                On Error Resume Next
                Cnt = Obj.Fields.Count
                On Error GoTo ErrorProcedure

                For I = 0 To Cnt

                    Ind = I

                    On Error Resume Next
                    Ind = Chr(34) + Obj.Fields(I).Name + Chr(34)
                    On Error GoTo ErrorProcedure

                    strVal = EseguiValutazione(Nome + "(" & Ind & ").Value", False)

                    If strVal <> "" Then

                        strVal = Replace(strVal, Nome + "(" & Ind & ").Value", Nome + "(" & Ind & ")")

                        Valori = Valori + strVal + vbCrLf

                    End If

                Next I

            Case "ixmldomnodelist"

                '------------------------------
                ' Espongo il nome dell'oggetto
                '------------------------------

                Valori = Nome + " = [" + TypeName(Obj) + "]" + vbCrLf

                '--------------------------------------
                ' Espongo i valori delle sue proprietà
                '--------------------------------------

                strVal = EseguiValutazione(Nome + ".length", False)

                If strVal <> "" Then

                    Valori = Valori + strVal + vbCrLf

                End If

            Case "ixmldomelement"

                '------------------------------
                ' Espongo il nome dell'oggetto
                '------------------------------

                Valori = Nome + " = [" + TypeName(Obj) + "]" + vbCrLf

                '--------------------------------------
                ' Espongo i valori delle sue proprietà
                '--------------------------------------

                Cnt = -1

                On Error Resume Next
                Cnt = Obj.Attributes.Length
                On Error GoTo ErrorProcedure

                For I = 0 To Cnt - 1

                    Ind = ""

                    On Error Resume Next
                    Ind = Chr(34) + Obj.Attributes(I).Basename + Chr(34)
                    On Error GoTo ErrorProcedure

                    strVal = EseguiValutazione(Nome + ".Attributes(" & I & ").text", False)

                    If strVal <> "" Then

                        If Ind <> "" Then
                            strVal = Replace(strVal, Nome + ".Attributes(" & I & ").text", Nome + "(" & Ind & ")")
                        End If

                        Valori = Valori + strVal + vbCrLf

                    End If

                Next I

            Case Else

                '------------------------------
                ' Espongo il nome dell'oggetto
                '------------------------------

                Valori = Nome + " = [" + TypeName(Obj) + "]" + vbCrLf

                '------------------------------
                ' Tento di vedere se è un form
                '------------------------------

                Cnt = -1

                On Error Resume Next
                Cnt = Obj.Controls.Count
                On Error GoTo ErrorProcedure

                For I = 0 To Cnt - 1

                    strVal = ""
                    Ind = ""
                    VarAux = String.Empty

                    On Error Resume Next

                    Ind = "(" & Obj.Controls(I).Index & ")"

                    VarAux = Obj.Controls(I)

                    If VarType(VarAux) = vbString Then
                        VarAux = Chr(34) + VarAux + Chr(34)
                    Else
                        VarAux = Str(VarAux)
                    End If

                    strVal = Nome + "." & Obj.Controls(I).Name + Ind + "=" & VarAux

                    On Error GoTo ErrorProcedure

                    If strVal <> "" Then

                        Valori = Valori + strVal + vbCrLf

                    End If

                Next I

        End Select

        If Right(Valori, 2) = vbCrLf Then
            Valori = Left(Valori, Len(Valori) - 2)
        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        Valori = Nome + " = [" + TypeName(Obj) + "]"

    End Sub

    Private Sub AnalizzaEspressione(ByVal Espressione As String, ByRef OnErrorRaise As Boolean, ByRef Comando As EnumCommand, ByRef ErrDescr As String)

        On Error GoTo ErrorProcedure

        Dim K As String
        Dim I As Long
        Dim CopiaEspress As String
        Dim FlagStr As Boolean
        Dim FlagId As Boolean
        Dim CurrId As String
        Dim CurrIn As String
        Dim LivParentesi As Long
        Dim MaxParentesi As Long
        Dim CurrPar As Long
        Dim TempId As String
        Dim Lunghezza As Long
        Dim Inizio As Long
        Dim FlagStop As Boolean
        Dim Argomenti As String = ""
        Dim G As Long
        Dim SalvaOnError As Boolean
        Dim Ripristini() As StructRipristino
        Dim VettLivEspr() As StructParentesi
        Dim VettParentesi() As StructParentesi

        ReDim Ripristini(0)
        ReDim VettLivEspr(100)
        ReDim VettParentesi(100)

        CopiaEspress = Trim(Espressione)
        FlagStr = False
        FlagId = False
        Lunghezza = Len(CopiaEspress)
        LivParentesi = 0
        MaxParentesi = 0

        With VettLivEspr(0)

            .Id = ""
            .Inizio = 1
            .Arg = 0
            .Fine = Lunghezza
            .Liv = 0
            .MaxComma = 0
            ReDim .Comma(20)

        End With

        For I = Lunghezza To 1 Step -1

            K = Mid(CopiaEspress, I, 1)

            If FlagStr Then

                If K = Chr(34) Then
                    FlagStr = False
                End If

            ElseIf K = Chr(34) Then

                FlagStr = True

            Else

                Select Case K

                    Case ")"

                        LivParentesi = LivParentesi + 1

                        If LivParentesi > UBound(VettLivEspr) Then
                            ReDim Preserve VettLivEspr(LivParentesi + 100)
                        End If

                        With VettLivEspr(LivParentesi)

                            .Id = ""
                            .Inizio = 0
                            .Arg = 0
                            .Fine = I
                            .Liv = LivParentesi
                            .MaxComma = 0
                            ReDim .Comma(20)

                        End With

                    Case "("

                        VettLivEspr(LivParentesi).Arg = I
                        CurrIn = I

                        FlagId = False
                        CurrId = ""

                        Do While I > 1

                            I = I - 1
                            K = Mid(CopiaEspress, I, 1)

                            Select Case K

                                Case "A" To "Z", "a" To "z", "0" To "9", "_", "."

                                    FlagId = True
                                    CurrId = K + CurrId
                                    CurrIn = I

                                Case " "

                                    If FlagId Then

                                        I = I + 1
                                        Exit Do

                                    End If

                                Case Else

                                    I = I + 1
                                    Exit Do

                            End Select

                        Loop

                        VettLivEspr(LivParentesi).Id = CurrId
                        VettLivEspr(LivParentesi).Inizio = CurrIn

                        MaxParentesi = MaxParentesi + 1

                        If MaxParentesi > UBound(VettParentesi) Then
                            ReDim Preserve VettParentesi(MaxParentesi + 100)
                        End If

                        VettParentesi(MaxParentesi) = VettLivEspr(LivParentesi)

                        LivParentesi = LivParentesi - 1

                    Case ","

                        With VettLivEspr(LivParentesi)

                            .MaxComma = .MaxComma + 1

                            If .MaxComma > UBound(.Comma) Then
                                ReDim Preserve .Comma(.MaxComma + 20)
                            End If

                            .Comma(.MaxComma) = I

                        End With

                End Select

            End If

        Next I

        '---------------------------
        ' Potrebbe esserci una call
        '---------------------------

        I = 0
        CurrId = ""

        Do While I <= Lunghezza

            I = I + 1
            K = Mid(CopiaEspress, I, 1)

            Select Case K

                Case "A" To "Z", "a" To "z", "0" To "9", "_", "."

                    CurrId = CurrId + K

                Case Else

                    Exit Do

            End Select

        Loop

        If CurrId <> "" Then

            MaxParentesi = MaxParentesi + 1

            If MaxParentesi > UBound(VettParentesi) Then
                ReDim Preserve VettParentesi(MaxParentesi + 100)
            End If

            With VettParentesi(MaxParentesi)

                .Id = CurrId
                .Inizio = 1
                .Arg = Len(CurrId) + 1
                .Fine = Lunghezza + 1
                .Liv = 0
                .MaxComma = VettLivEspr(0).MaxComma
                .Comma = VettLivEspr(0).Comma

            End With

        End If

        ReDim Preserve VettParentesi(MaxParentesi)

        '---------------------------------------------------------------------------
        ' Analizzo le parentesi per vedere se sono chiamate a funzioni dello script
        '---------------------------------------------------------------------------

        CurrPar = 1

        Dim Delta As Long

        Do Until CurrPar > MaxParentesi

            If VettParentesi(CurrPar).Id <> "" Then

                Inizio = EntryPoint(VettParentesi(CurrPar).Id, Argomenti)

                If Inizio > 0 Then

                    FlagStop = False

                    If Comando <> cmdNextSub Then

                        For I = Inizio To UBound(VettRighe)

                            Select Case VettRighe(I).Tipo

                                Case trExitSub, trExitFunction, trEndSub, trEndSub

                                    Exit For

                                Case trEnd, trStop

                                    FlagStop = True

                                    Exit For

                            End Select

                        Next I

                    End If

                    If Comando = cmdNextSub Or FlagStop = True Then

                        PassaggioArgomenti(CopiaEspress, VettParentesi, CurrPar, Argomenti, Ripristini)

                        SalvaOnError = OnErrorRaise
                        TempId = InvocaFunzione(VettParentesi(CurrPar).Id, OnErrorRaise, Comando, False)
                        OnErrorRaise = SalvaOnError

                        RipristinaValori(Ripristini)

                        Delta = Len(TempId) - (VettParentesi(CurrPar).Fine - VettParentesi(CurrPar).Inizio) - 1

                        CopiaEspress = Left(CopiaEspress, VettParentesi(CurrPar).Inizio - 1) + TempId + Mid(CopiaEspress, VettParentesi(CurrPar).Fine + 1)

                        For I = 1 To MaxParentesi

                            If VettParentesi(I).Inizio > VettParentesi(CurrPar).Inizio Then
                                VettParentesi(I).Inizio = VettParentesi(I).Inizio + Delta
                            End If

                            If VettParentesi(I).Fine > VettParentesi(CurrPar).Inizio Then
                                VettParentesi(I).Fine = VettParentesi(I).Fine + Delta
                            End If

                            If VettParentesi(I).Arg > VettParentesi(CurrPar).Inizio Then
                                VettParentesi(I).Arg = VettParentesi(I).Arg + Delta
                            End If

                            For G = 1 To VettParentesi(I).MaxComma

                                If VettParentesi(I).Comma(G) > VettParentesi(CurrPar).Inizio Then
                                    VettParentesi(I).Comma(G) = VettParentesi(I).Comma(G) + Delta
                                End If

                            Next G

                        Next I

                    End If

                End If

            End If

            CurrPar = CurrPar + 1

        Loop

        PropScript.ExecuteStatement(CopiaEspress)

        Exit Sub

ErrorProcedure:

        ErrDescr = Err.Description

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Private Function EntryPoint(SubName As String, ByRef Argomenti As String) As Long

        Try

            Dim RigaInizio As Long
            Dim F As Long

            SubName = LCase(SubName)
            RigaInizio = 0

            For F = 1 To UBound(VettFunct)

                If VettFunct(F).Normale = SubName Then

                    RigaInizio = VettFunct(F).EntryPoint
                    Argomenti = VettFunct(F).Argomenti

                    Exit For

                End If

            Next F

            Return RigaInizio

        Catch ex As Exception

            Return 0

        End Try

        Exit Function

    End Function

    Private Function ValutaEspressione(ByRef Espressione As String, ByRef R As Long, ByRef Comando As EnumCommand, OnErrorRaise As Boolean, DefaultValue As Object) As Object

        On Error GoTo ErrorProcedure

        Dim Ritorno As Object
        Dim TempId As String
        Dim Commento As String
        Dim SalvaOnError As Boolean

        SerialeID = SerialeID + 1
        TempId = "tmp___" & SerialeID

        SalvaOnError = OnErrorRaise
        AnalizzaEspressione(TempId + " = " + Espressione, OnErrorRaise, Comando, Commento)
        OnErrorRaise = SalvaOnError

        If Commento <> "" Then

            Comando = cmdNext

            GoTo AbortProcedure

        End If

        Ritorno = PropScript.Eval(TempId)
        Return Ritorno

ErrorProcedure:

        Commento = PropScript.Error.Description

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        If OnErrorRaise Then
            Comando = cmdNext
        End If

        AttesaComando(OnErrorRaise, R, Commento, Comando, -1)

        If OnErrorRaise Then

            R = 0

        End If

        Return DefaultValue

    End Function

    Protected Overrides Sub Finalize()

        FlagTermine = True

        MyBase.Finalize()

    End Sub

    Private Sub PassaggioArgomenti(Espressione As String, VettParentesi() As StructParentesi, Index As Long, Argomenti As String, ByRef Ripristini() As StructRipristino)

        On Error GoTo ErrorProcedure

        Dim ProtArgs As Object
        Dim I As Long
        Dim MaxArg As Long
        Dim PosPrev As Long
        Dim PosSucc As Long
        Dim Id As String
        Dim MaxSalva As Long
        Dim VettArgs() As String

        If Argomenti <> "" Then

            ProtArgs = Split(Argomenti, ",")

            MaxArg = VettParentesi(Index).MaxComma + 1

            ReDim VettArgs(MaxArg)
            ReDim Ripristini(MaxArg)

            PosPrev = VettParentesi(Index).Arg

            For I = 1 To MaxArg

                If I < MaxArg Then
                    PosSucc = VettParentesi(Index).Comma(I)
                Else
                    PosSucc = VettParentesi(Index).Fine
                End If

                VettArgs(I) = Trim(Mid(Espressione, PosPrev + 1, PosSucc - PosPrev - 1))

                PosPrev = PosSucc

            Next I

            MaxSalva = 0

            For I = 0 To UBound(ProtArgs)

                If I + 1 <= MaxArg Then

                    MaxSalva = MaxSalva + 1

                    SerialeID = SerialeID + 1
                    Id = "tmp___" & SerialeID

                    '---------------------------------------------
                    ' Salvo i valori in variabili temporanee
                    ' e memorizzo delle espressioni di ripristino
                    '---------------------------------------------

                    If VarType(PropScript.Eval((ProtArgs(I)))) = VariantType.Object Then

                        PropScript.ExecuteStatement("Set " & Id & "=" & ProtArgs(I))
                        Ripristini(MaxSalva).Ripristino = "Set " & ProtArgs(I) & "=" & Id

                    Else

                        PropScript.ExecuteStatement(Id & "=" & ProtArgs(I))
                        Ripristini(MaxSalva).Ripristino = ProtArgs(I) & "=" & Id

                    End If

                    Ripristini(MaxSalva).Argomento = ProtArgs(I)
                    Ripristini(MaxSalva).Espressione = VettArgs(I + 1)

                    '-----------------------------------------------------
                    ' Assegno i valori passati alle variabili prototipali
                    '-----------------------------------------------------

                    If VarType(PropScript.Eval(VettArgs(I + 1))) = VariantType.Object Then
                        PropScript.ExecuteStatement("Set " & ProtArgs(I) & "=" & VettArgs(I + 1))
                    Else
                        PropScript.ExecuteStatement(ProtArgs(I) & "=" & VettArgs(I + 1))
                    End If

                End If

            Next I

            ReDim Preserve Ripristini(MaxSalva)

        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Private Sub RipristinaValori(Ripristini() As StructRipristino)

        On Error Resume Next

        Dim I As Long

        '----------------------
        ' Porto fuori i valori
        '----------------------

        For I = 1 To UBound(Ripristini)

            If VarType(PropScript.Eval(Ripristini(I).Argomento)) = VariantType.Object Then

                PropScript.ExecuteStatement("Set " & Ripristini(I).Espressione & "=" & Ripristini(I).Argomento)

            Else

                PropScript.ExecuteStatement(Ripristini(I).Espressione & "=" & Ripristini(I).Argomento)

            End If

        Next I

        For I = 1 To UBound(Ripristini)

            If StrComp(Ripristini(I).Argomento, Ripristini(I).Espressione, vbTextCompare) <> 0 Then
                PropScript.ExecuteStatement(Ripristini(I).Ripristino)
            End If

        Next I

    End Sub

End Class

