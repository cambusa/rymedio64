Option Explicit On

Imports System.Text.RegularExpressions
Imports ryMedio64.moduleDebug.EnumTipoRiga

Module moduleDebug

    Public GlobalErrNumber As Long
    Public GlobalErrDescription As String

    Public Enum EnumTipoRiga

        trNull
        trDim
        trStatement
        trIf
        trIfThen
        trIfThenElse
        trElseIf
        trElse
        trEndIf
        trFor
        trForStep
        trForEach
        trExitFor
        trNext
        trDo
        trDoUntil
        trDoWhile
        trExitDo
        trLoop
        trLoopUntil
        trLoopWhile
        trSelect
        trCase
        trCaseElse
        trEndSelect
        trFunct
        trEndFunct
        trSub
        trEndSub
        trWith
        trEndWith
        trExitFunction
        trExitSub
        trResumeNext
        trGoToZero
        trGoTo
        trStop
        trEnd
        trRem
        trInside

    End Enum

    Public Structure StructRiga

        Public Orig As String
        Public Pura As String
        Public Tipo As EnumTipoRiga

        Public Inizio As Long
        Public Madre As Long
        Public Figlia As Long
        Public Sorella As Long

        Public Livello As Long

        Public Nome As String
        Public Espr1 As String
        Public Espr2 As String
        Public Espr3 As String

        Public CaseEspr As Object

        Public DefaultName As String

        Public FlagCiclo As Boolean
        Public Confronto As Object
        Public Index As Long

    End Structure

    Public Structure StructLivello

        Public Inizio As Long
        Public Sorella As Long
        Public DefaultName As String

    End Structure

    Structure StructValutazione

        Public Espressione As String
        Public Valore As Object
        Public FlagInside As Boolean

    End Structure

    Structure StructFunction

        Public Nome As String
        Public Normale As String
        Public EntryPoint As Long
        Public Argomenti As String

    End Structure

    Structure StructStatoWin

        Public L As Single      ' left
        Public T As Single      ' top
        Public W As Single      ' width
        Public H As Single      ' height
        Public B As Single      ' posizione barra
        Public M As Long        ' statewindows

    End Structure

    Structure StructParentesi

        Public Id As String
        Public Inizio As Long
        Public Arg As Long
        Public Fine As Long
        Public Liv As Long
        Public MaxComma As Long
        Public Comma As Object

    End Structure

    Structure StructRipristino

        Public Argomento As String
        Public Espressione As String
        Public Ripristino As String

    End Structure

    Public Function AnalizzaSorgente(ObjRegEx As Regex, Directory As String, ByVal Text As String, ByRef VettRighe() As StructRiga, ByRef VettEval() As StructValutazione, ByRef VettFunct() As StructFunction) As Boolean

        On Error GoTo ErrorProcedure

        Dim Esito As Boolean
        Dim MaxRiga As Long
        Dim M As MatchCollection
        Dim S As MatchCollection
        Dim R As Long
        Dim I As Long
        Dim Buff As String
        Dim PosOffset As Long
        Dim StrLen As Long
        Dim V As Object
        Dim ElencoVal As String

        Dim UltimaRiga As Long
        Dim UltimoTipo As EnumTipoRiga

        Dim FlagVuoto As Boolean

        Dim PrevLivello As Long

        Dim CurrDefaultName As String

        Dim CurrLivello As Long
        Dim VettLivelli() As StructLivello
        ReDim VettLivelli(100)

        Dim MaxEval As Long
        ReDim VettEval(100)

        Dim MaxFunct As Long
        ReDim VettFunct(100)

        ElencoVal = "|"

        '-------------------------
        ' Sostituzione tabulatori
        '-------------------------

        Text = Replace(Text, vbTab, Space(5))

        '----------------------------------
        ' Giunzione delle righe continuate
        '----------------------------------

        Text = CreaRegExp(" _\r\n *").Replace(Text, " ")
        Text = CreaRegExp(" _\r *").Replace(Text, " ")
        Text = CreaRegExp(" _\n *").Replace(Text, " ")

        '-----------------------------------
        ' Eliminazione prefisso metacomando
        '-----------------------------------

        ObjRegEx = CreaRegExp("((^|\r|\n) *'[@])")
        M = ObjRegEx.Matches(Text)
        ObjRegEx = Nothing

        For R = M.Count - 1 To 0 Step -1

            Buff = M(R).Groups(1).Value
            I = M(R).Index
            Text = Left(Text, I) + Left(Buff, Len(Buff) - 2) + Mid(Text, I + Len(Buff) + 1)

        Next R

        M = Nothing

        '---------------------------------
        ' Disabilitazione Option Explicit
        '---------------------------------

        Text = CreaRegExp("^ *option +explicit").Replace(Text, "'option explicit")

        '-------------------------------------
        ' Sostituzione caratteri di controllo
        '-------------------------------------

        Text = CreaRegExp("[\x00-\x09]").Replace(Text, "?")
        Text = CreaRegExp("[\x0B\x0C]").Replace(Text, "?")
        Text = CreaRegExp("[\x0E\x1F]").Replace(Text, "?")

        '---------------
        ' Analisi righe
        '---------------

        ObjRegEx = CreaRegExp("(.*)(\r|\n|$)")
        M = ObjRegEx.Matches(Text)
        ObjRegEx = Nothing

        MaxRiga = M.Count

        ReDim VettRighe(MaxRiga)

        For R = 1 To MaxRiga

            Buff = M(R - 1).Groups(1).Value.Replace(vbLf, "").Replace(vbCr, "")

            VettRighe(R).Orig = Buff

            ObjRegEx = CreaRegExp(Chr(34) + "[^" + Chr(34) + "]*" + Chr(34))
            S = ObjRegEx.Matches(Buff)
            ObjRegEx = Nothing

            For I = 0 To S.Count - 1
                Mid(Buff, S(I).Index + 1) = New String("§", S(I).Length)
            Next I
            S = Nothing

            PosOffset = InStr(Buff, "'")

            If PosOffset > 0 Then
                Buff = Left(Buff, PosOffset - 1)
            End If

            VettRighe(R).Pura = RTrim(Buff)

        Next R

        M = Nothing

        CurrLivello = 1
        UltimaRiga = 0
        UltimoTipo = trNull
        MaxEval = 0
        CurrDefaultName = ""

        For R = 1 To MaxRiga

            VettRighe(R).Tipo = trNull

            If Trim(VettRighe(R).Pura) <> "" Then

                '-----
                ' DIM
                '-----

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *dim )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trDim

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        AggiungiValutazione(ElencoVal, VettEval, MaxEval, VettRighe(R).Nome)

                    End If

                    M = Nothing

                End If

                '-------
                ' WATCH
                '-------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *watch )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trDim

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        AggiungiValutazione(ElencoVal, VettEval, MaxEval, VettRighe(R).Nome)

                    End If

                    M = Nothing

                End If

                '--------
                ' INSIDE
                '--------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *inside )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trInside

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        MaxEval = MaxEval + 1

                        If MaxEval > UBound(VettEval) Then
                            ReDim Preserve VettEval(MaxEval + 100)
                        End If

                        VettEval(MaxEval).Espressione = VettRighe(R).Nome
                        VettEval(MaxEval).Valore = String.Empty
                        VettEval(MaxEval).FlagInside = True

                    End If

                    M = Nothing

                End If

                '------
                ' GOTO
                '------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *goto )(\d+) *$")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trGoTo

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '------
                ' REM
                '------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *rem )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trRem

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '------
                ' STOP
                '------

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "stop" Then

                        VettRighe(R).Tipo = trStop

                    End If

                End If

                '-----
                ' END
                '-----

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "end" Then

                        VettRighe(R).Tipo = trEnd

                    End If

                End If

                '----------
                ' FUNCTION
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^((private|public|) *function )(.+)[(](.*)[)]")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trFunct

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(3).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 1
                        StrLen = Len(M(0).Groups(4).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        MaxFunct = MaxFunct + 1

                        If MaxFunct > UBound(VettFunct) Then
                            ReDim Preserve VettFunct(MaxFunct + 100)
                        End If

                        VettFunct(MaxFunct).Nome = VettRighe(R).Nome
                        VettFunct(MaxFunct).Normale = LCase(VettRighe(R).Nome)
                        VettFunct(MaxFunct).Argomenti = VettRighe(R).Espr1.Replace(" ", "")
                        VettFunct(MaxFunct).EntryPoint = R

                        If VettFunct(MaxFunct).Argomenti <> "" Then

                            V = Split(VettFunct(MaxFunct).Argomenti, ",")

                            For I = 0 To UBound(V)

                                AggiungiValutazione(ElencoVal, VettEval, MaxEval, (V(I)))

                            Next I

                        End If

                    End If

                    M = Nothing

                End If

                '--------------
                ' END FUNCTION
                '--------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *end +function")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then
                        VettRighe(R).Tipo = trEndFunct
                    End If

                    M = Nothing

                End If

                '-----
                ' SUB
                '-----

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^((private|public|) *sub )(.+)[(](.*)[)]")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trSub

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(3).Value)

                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 1
                        StrLen = Len(M(0).Groups(4).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        MaxFunct = MaxFunct + 1

                        If MaxFunct > UBound(VettFunct) Then
                            ReDim Preserve VettFunct(MaxFunct + 100)
                        End If

                        VettFunct(MaxFunct).Nome = VettRighe(R).Nome
                        VettFunct(MaxFunct).Normale = LCase(VettRighe(R).Nome)
                        VettFunct(MaxFunct).Argomenti = VettRighe(R).Espr1.Replace(" ", "")
                        VettFunct(MaxFunct).EntryPoint = R

                        If VettFunct(MaxFunct).Argomenti <> "" Then

                            V = Split(VettFunct(MaxFunct).Argomenti, ",")

                            For I = 0 To UBound(V)

                                AggiungiValutazione(ElencoVal, VettEval, MaxEval, (V(I)))

                            Next I

                        End If

                    End If

                    M = Nothing

                End If

                '---------
                ' END SUB
                '---------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *end +sub")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)
                    ObjRegEx = Nothing

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trEndSub

                    End If

                    M = Nothing

                End If

                '----------
                ' EXIT SUB
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *exit +sub")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trExitSub

                    End If

                    M = Nothing

                End If

                '---------------
                ' EXIT FUNCTION
                '---------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *exit +function")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trExitFunction

                    End If

                    M = Nothing

                End If

                '----------------------
                ' ON ERROR RESUME NEXT
                '----------------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *on +error +resume +next")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trResumeNext

                        VettRighe(R).Espr1 = Trim(VettRighe(R).Pura)

                    End If

                    M = Nothing

                End If

                '-----------------
                ' ON ERROR GOTO 0
                '-----------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *on +error +goto +0")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trGoToZero

                        VettRighe(R).Espr1 = Trim(VettRighe(R).Pura)

                    End If

                    M = Nothing

                End If

                '--------
                ' ELSEIF
                '--------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *elseif )(.+) then")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trElseIf

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '--------------
                ' IF THEN ELSE
                '--------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *if (.+) then (.+) else (.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trIfThenElse
                        VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)

                    End If

                    M = Nothing

                End If

                '---------
                ' IF THEN
                '---------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *if (.+) then (.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trIfThen
                        VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)

                    End If

                    M = Nothing

                End If

                '----
                ' IF
                '----

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *)if (.+) then")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trIf

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 4
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '-------------
                ' SELECT CASE
                '-------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *select +case )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trSelect

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '-----------
                ' CASE ELSE
                '-----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *case +else")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trCaseElse

                    End If

                    M = Nothing

                End If

                '------
                ' CASE
                '------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *case )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trCase

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        VettRighe(R).CaseEspr = AnalizzaCase(VettRighe(R).Espr1)

                    End If

                    M = Nothing

                End If

                '------------
                ' END SELECT
                '------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *end +select")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trEndSelect

                    End If

                    M = Nothing

                End If

                '------
                ' ELSE
                '------

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "else" Then

                        VettRighe(R).Tipo = trElse

                    End If

                End If

                '--------
                ' END IF
                '--------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *end +if")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trEndIf

                    End If

                    M = Nothing

                End If

                '----------
                ' END WITH
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *end +with")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trEndWith

                    End If

                    M = Nothing

                End If

                '------
                ' WITH
                '------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *with )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trWith

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        CurrDefaultName = VettRighe(R).Espr1

                    End If

                    M = Nothing

                End If

                '-------------
                ' FOR EACH IN
                '-------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *for +each )(.+) in (.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trForEach

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)
                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 4
                        StrLen = Len(M(0).Groups(3).Value)
                        VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        VettRighe(R).Nome = VettRighe(R).Espr1

                    End If

                    M = Nothing

                End If

                '-------------
                ' FOR TO STEP
                '-------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *for )(.+) to (.+) step (.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trForStep

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)
                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 4
                        StrLen = Len(M(0).Groups(3).Value)
                        VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 6
                        StrLen = Len(M(0).Groups(4).Value)
                        VettRighe(R).Espr3 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = InStr(VettRighe(R).Espr1, "=")
                        If PosOffset > 0 Then
                            VettRighe(R).Nome = Trim(Mid(VettRighe(R).Espr1, 1, PosOffset - 1))
                        End If

                    End If

                    M = Nothing

                End If

                '--------
                ' FOR TO
                '--------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *for )(.+) to (.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trFor

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)
                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        PosOffset = PosOffset + StrLen + 4
                        StrLen = Len(M(0).Groups(3).Value)
                        VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                        VettRighe(R).Espr3 = "1"

                        PosOffset = InStr(VettRighe(R).Espr1, "=")
                        If PosOffset > 0 Then
                            VettRighe(R).Nome = Trim(Mid(VettRighe(R).Espr1, 1, PosOffset - 1))
                        End If

                    End If

                    M = Nothing

                End If

                '----------
                ' EXIT FOR
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *exit +for")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trExitFor

                    End If

                    M = Nothing

                End If

                '------
                ' NEXT
                '------

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "next" Then

                        VettRighe(R).Tipo = trNext

                    End If

                End If

                '----------
                ' DO UNTIL
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *do +until )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trDoUntil

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = "Not (" + Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen)) + ")"

                    End If

                    M = Nothing

                End If

                '----------
                ' DO WHILE
                '----------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *do +while )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trDoWhile

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '---------
                ' EXIT DO
                '---------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^ *exit +do")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trExitDo

                    End If

                    M = Nothing

                End If

                '------------
                ' LOOP UNTIL
                '------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *loop +until )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trLoopUntil

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = "Not (" + Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen)) + ")"

                    End If

                    M = Nothing

                End If

                '------------
                ' LOOP WHILE
                '------------

                If VettRighe(R).Tipo = trNull Then

                    ObjRegEx = CreaRegExp("^( *loop +while )(.+)")
                    M = ObjRegEx.Matches(VettRighe(R).Pura)

                    If M.Count = 1 Then

                        VettRighe(R).Tipo = trLoopWhile

                        PosOffset = M(0).Index + Len(M(0).Groups(1).Value) + 1
                        StrLen = Len(M(0).Groups(2).Value)

                        VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))

                    End If

                    M = Nothing

                End If

                '----
                ' DO
                '----

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "do" Then

                        VettRighe(R).Tipo = trDo

                        VettRighe(R).Espr1 = "-1"

                    End If

                End If

                '------
                ' LOOP
                '------

                If VettRighe(R).Tipo = trNull Then

                    If LCase(Trim(VettRighe(R).Pura)) = "loop" Then

                        VettRighe(R).Tipo = trLoop

                        VettRighe(R).Espr1 = "-1"

                    End If

                End If

                '----------
                ' CHIAMATA
                '----------

                If VettRighe(R).Tipo = trNull Then

                    VettRighe(R).Tipo = trStatement
                    VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)

                End If

                '----------------------------
                ' Gestione livello
                '----------------------------

                VettRighe(R).Livello = 0
                VettRighe(R).Inizio = 0
                VettRighe(R).Madre = 0
                VettRighe(R).Figlia = 0

                FlagVuoto = BloccoVuoto(UltimoTipo, VettRighe(R).Tipo)

                Select Case UltimoTipo

                    Case trCase, trCaseElse, trDo, trDoUntil, trDoWhile, trElse, trElseIf, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith

                        VettLivelli(CurrLivello).Inizio = UltimaRiga

                        If Not FlagVuoto Then

                            CurrLivello = CurrLivello + 1

                            VettLivelli(CurrLivello).Inizio = R

                            VettRighe(UltimaRiga).Figlia = R

                        End If

                End Select

                Select Case VettRighe(R).Tipo

                    Case trCase, trCaseElse, trDo, trDoUntil, trDoWhile, trElse, trElseIf, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith

                        If CurrDefaultName <> "" Then

                            VettLivelli(CurrLivello + 1).DefaultName = VettLivelli(CurrLivello).DefaultName + CurrDefaultName

                            CurrDefaultName = ""

                        Else

                            VettLivelli(CurrLivello + 1).DefaultName = VettLivelli(CurrLivello).DefaultName

                        End If

                End Select

                Select Case VettRighe(R).Tipo

                    Case trEndFunct, trEndIf, trEndSelect, trEndSub, trEndWith, trLoop, trLoopUntil, trLoopWhile, trNext

                        If Not FlagVuoto Then

                            CurrLivello = CurrLivello - 1

                        End If

                        If VettRighe(R).Tipo = trEndSelect Then
                            CurrLivello = CurrLivello - 1
                        End If

                        VettRighe(R).Livello = CurrLivello
                        VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio

                    Case trElse, trElseIf

                        If Not FlagVuoto Then

                            CurrLivello = CurrLivello - 1

                        End If

                        VettRighe(R).Livello = CurrLivello
                        VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio

                    Case trCase, trCaseElse

                        If Not FlagVuoto Then

                            If VettRighe(UltimaRiga).Tipo <> trSelect Then
                                CurrLivello = CurrLivello - 1
                            Else
                                VettRighe(UltimaRiga).Figlia = R
                            End If

                        End If

                        VettRighe(R).Livello = CurrLivello
                        VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio

                    Case trDo, trDoUntil, trDoWhile, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith

                        VettRighe(R).Livello = CurrLivello
                        VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio
                        VettLivelli(CurrLivello).Inizio = R

                    Case Else

                        If UltimaRiga > 0 Then

                            If CurrLivello = VettRighe(UltimaRiga).Livello Then
                                VettRighe(R).Madre = VettRighe(UltimaRiga).Madre
                            Else
                                VettRighe(R).Madre = UltimaRiga
                            End If

                            VettRighe(R).Livello = CurrLivello

                        Else

                            VettRighe(R).Livello = CurrLivello
                            VettLivelli(CurrLivello).Inizio = R

                        End If

                End Select

                VettRighe(R).DefaultName = VettLivelli(CurrLivello).DefaultName

                If VettRighe(R).DefaultName <> "" Then

                    NormalizzaPredefinito(VettRighe(R).Espr1, VettRighe(R).DefaultName)
                    NormalizzaPredefinito(VettRighe(R).Espr2, VettRighe(R).DefaultName)
                    NormalizzaPredefinito(VettRighe(R).Espr3, VettRighe(R).DefaultName)

                End If

                UltimaRiga = R
                UltimoTipo = VettRighe(R).Tipo

            End If

        Next R

        UltimaRiga = 0      ' Precedente riga non vuota
        PrevLivello = 0     ' Livello di UltimaRiga
        CurrLivello = 0     ' Livello di R

        For R = MaxRiga To 1 Step -1

            If Trim(VettRighe(R).Pura) <> "" Then

                CurrLivello = VettRighe(R).Livello

                If CurrLivello > PrevLivello Then

                    '-------------------------------
                    ' Indentazione: resetto Sorella
                    '-------------------------------

                    VettLivelli(CurrLivello).Sorella = 0

                End If

                '--------------------------------------------------------
                ' Se non siamo sull'ultima riga dello script,
                ' memorizzo in Sorella del precedente livello UltimaRiga
                '--------------------------------------------------------

                If PrevLivello > 0 Then
                    VettLivelli(PrevLivello).Sorella = UltimaRiga
                End If

                '--------------------------------------------------------------------------
                ' Assegno Sorella corrente che ho bufferizziato nella struttura di livello
                '--------------------------------------------------------------------------

                VettRighe(R).Sorella = VettLivelli(CurrLivello).Sorella

                '-------------------------------------------------------------------------------
                ' Se sono all'inizio di un ciclo FOR o DO, assegno inizio alla NEXT o alla LOOP
                '-------------------------------------------------------------------------------

                I = VettRighe(R).Sorella

                If I > 0 Then

                    Select Case VettRighe(I).Tipo

                        Case trNext, trLoop, trLoopUntil, trLoopWhile
                            VettRighe(I).Inizio = R

                    End Select

                End If

                '-------------------------------
                ' Memorizzo lo stato precedente
                '-------------------------------

                PrevLivello = CurrLivello
                UltimaRiga = R

            End If

        Next R

        M = Nothing
        ObjRegEx = Nothing

        'ScriviHTML Directory + "\~debug.htm", VettRighe(), 0, ""

        ReDim Preserve VettEval(MaxEval)
        ReDim Preserve VettFunct(MaxFunct)

        Return Esito

ErrorProcedure:

        GlobalErrDescription = Err.Description
        GlobalErrNumber = Err.Number

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        M = Nothing
        S = Nothing
        ObjRegEx = Nothing

        ReDim Preserve VettEval(0)
        ReDim Preserve VettFunct(0)

        Return False

    End Function

    Private Function BloccoVuoto(PrecTipo As EnumTipoRiga, CurrTipo As EnumTipoRiga) As Boolean

        Dim Esito As Boolean

        Esito = False

        Select Case PrecTipo

            Case trCase, trCaseElse

                Select Case CurrTipo

                    Case trCase, trCaseElse, trEndSelect

                        Esito = True

                End Select

            Case trDo, trDoUntil, trDoWhile

                Select Case CurrTipo

                    Case trLoop, trLoopUntil, trLoopWhile

                        Esito = True

                End Select

            Case trElse, trElseIf, trIf

                Select Case CurrTipo

                    Case trElse, trElseIf, trEndIf

                        Esito = True

                End Select

            Case trFor, trForStep, trForEach

                Select Case CurrTipo

                    Case trNext

                        Esito = True

                End Select

            Case trFunct, trSub

                Select Case CurrTipo

                    Case trEndFunct, trEndSub

                        Esito = True

                End Select

            Case trSelect

                Select Case CurrTipo

                    Case trEndSelect

                        Esito = True

                End Select

            Case trWith

                Select Case CurrTipo

                    Case trEndWith

                        Esito = True

                End Select

        End Select

        BloccoVuoto = Esito

    End Function

    Public Sub ScriviHTML(PathFile As String, VettRighe() As StructRiga, CurrentRow As Long, Messaggio As String)

        On Error GoTo ErrorProcedure

        Dim R As Long
        Dim Buffer As String
        Dim NumFile As Integer
        Dim TestataHTML As String
        Dim CodaHTML As String
        Dim StrClass As String
        Dim VettBuff() As String

        ReDim VettBuff(UBound(VettRighe))

        For R = 1 To UBound(VettRighe)

            If R = CurrentRow Then
                StrClass = " class='selected'"
            Else
                StrClass = ""
            End If

            If Trim(VettRighe(R).Pura) <> "" Then
                VettBuff(R - 1) = "<tr" + StrClass + "><td>" & R & "</td><td style='padding-left:20px;'>" + Replace(VettRighe(R).Orig, " ", "&nbsp;") + "</td><td style='padding-left:20px;'>" & VettRighe(R).Livello & "</td><td style='padding-left:20px;'>" & VettRighe(R).Madre & "</td><td style='padding-left:20px;'>" & VettRighe(R).Figlia & "</td><td style='padding-left:20px;'>" & VettRighe(R).Sorella & "</td><td style='padding-left:20px;'>" & VettRighe(R).Inizio & "</td><tr>"
            Else
                VettBuff(R - 1) = "<tr><td>" & R & "</td><td colspan='5' style='font-size:4px;'>&nbsp;</td><tr>"
            End If

        Next R

        TestataHTML = ""
        TestataHTML = TestataHTML + "<html>" + vbCrLf
        TestataHTML = TestataHTML + "<style>" + vbCrLf
        TestataHTML = TestataHTML + ".selected{background-color:yellow;}" + vbCrLf
        TestataHTML = TestataHTML + ".cell{padding-left:20px;}" + vbCrLf
        TestataHTML = TestataHTML + "</style>" + vbCrLf
        TestataHTML = TestataHTML + "<body style='font-family:verdana;font-size:12px;'>" + vbCrLf

        If Messaggio <> "" Then
            TestataHTML = TestataHTML + "<div style='background-color:red;color:white;'>" + Messaggio + "</div>" + vbCrLf
        End If

        TestataHTML = TestataHTML + "<table style='font-family:courier new;font-size:12px;border-collapse:collapse;'>" + vbCrLf
        TestataHTML = TestataHTML + "<tr><th></th><th></th><th style='padding-left:20px;'>Livello</th><th style='padding-left:20px;'>Madre</th><th style='padding-left:20px;'>Figlia</th><th style='padding-left:20px;'>Sorella</th><th style='padding-left:20px;'>Inizio</th></tr>" + vbCrLf

        Buffer = Join(VettBuff, vbCrLf)

        CodaHTML = ""
        CodaHTML = CodaHTML + "</table>" + vbCrLf
        CodaHTML = CodaHTML + "</body>" + vbCrLf
        CodaHTML = CodaHTML + "</html>" + vbCrLf

        Buffer = TestataHTML + Buffer + CodaHTML

        My.Computer.FileSystem.WriteAllText(PathFile, Buffer, False)

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

    End Sub

    Public Sub NormalizzaPredefinito(Espressione As String, DefaultName As String)

        Dim I As Long
        Dim Buffer As String = ""
        Dim FlagStr As Boolean

        Dim C As String
        Dim P As String
        Dim S As String

        Dim L As Long

        If Espressione <> "" Then

            If InStr(Espressione, ".") > 0 Then

                P = ""
                FlagStr = False
                L = Len(Espressione)

                For I = 1 To L

                    C = Mid(Espressione, I, 1)

                    If C = Chr(34) Then
                        FlagStr = Not FlagStr
                    End If

                    If FlagStr = False Then

                        If C = "." Then

                            If I < L Then
                                S = Mid(Espressione, I + 1, 1)
                            Else
                                S = ""
                            End If

                            If S < "0" Or "9" < S Then

                                Select Case P

                                    Case "", " ", ":", "=", "<", ">", "+", "-", "*", "/", "\", "^", "("
                                        C = DefaultName + "."

                                End Select

                            End If

                        End If

                    End If

                    Buffer = Buffer + C

                    P = C

                Next I

                Espressione = Buffer

            End If

        End If

    End Sub

    Sub PollingAttesa(Ritardo As Single)

        On Error Resume Next

        Dim Tempo As Single



        Tempo = Timer

        Do

            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(1)

        Loop Until Math.Abs(Timer - Tempo) > Ritardo

    End Sub

    Public Sub AggiungiValutazione(ByRef ElencoVal As String, ByRef VettEval() As StructValutazione, ByRef MaxEval As Long, Id As String)

        On Error Resume Next

        If InStr(1, ElencoVal, "|" & Id & "|", vbTextCompare) = 0 Then

            MaxEval = MaxEval + 1

            If MaxEval > UBound(VettEval) Then
                ReDim Preserve VettEval(MaxEval + 100)
            End If

            VettEval(MaxEval).Espressione = Id
            VettEval(MaxEval).Valore = String.Empty

            ElencoVal = ElencoVal + Id + "|"

        End If

    End Sub

    Private Function AnalizzaCase(CaseArgs As String) As String()

        On Error GoTo ErrorProcedure

        Dim V() As String
        Dim I As Long
        Dim K As String
        Dim FlagStr As Boolean
        Dim MaxArg As Long
        Dim PosPrev As Long

        MaxArg = 0
        ReDim V(20)
        PosPrev = 0
        FlagStr = False

        For I = 1 To Len(CaseArgs)

            K = Mid(CaseArgs, I, 1)

            If FlagStr Then

                If K = Chr(34) Then
                    FlagStr = False
                End If

            ElseIf K = Chr(34) Then

                FlagStr = True

            Else

                Select Case UCase(K)

                    Case ","

                        MaxArg = MaxArg + 1

                        If MaxArg > UBound(V) Then
                            ReDim Preserve V(MaxArg + 20)
                        End If

                        V(MaxArg) = Trim(Mid(CaseArgs, PosPrev + 1, I - PosPrev - 1))

                        PosPrev = I

                End Select

            End If

        Next I

        MaxArg = MaxArg + 1

        If MaxArg > UBound(V) Then
            ReDim Preserve V(MaxArg + 20)
        End If

        V(MaxArg) = Trim(Mid(CaseArgs, PosPrev + 1))

        ReDim Preserve V(MaxArg)

        Return V

        Exit Function

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        ReDim Preserve V(0)

        Return V

    End Function

    Public Function RisolviOggetto(ObjColl As Object, Indice As Long) As Object

        On Error GoTo ErrorProcedure

        Dim CurrInd As Long
        Dim Oggetto As Object = Nothing

        CurrInd = 0

        For Each Oggetto In ObjColl

            CurrInd = CurrInd + 1

            If CurrInd >= Indice Then

                Exit For

            End If

        Next

        If CurrInd <> Indice Then
            Oggetto = Nothing
        End If

        Return Oggetto

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        On Error Resume Next

        Return Nothing

    End Function

    Public Function CreaRegExp(pattern As String) As Regex

        Return New Regex(pattern, RegexOptions.Multiline Or RegexOptions.IgnoreCase)

    End Function

End Module
