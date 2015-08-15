'************************************************************************
'* Name:            monitor.vb                                          *
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
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Structure StructMonRow
    Public Rg As String
    Public Bk As Boolean
End Structure

Public Class monitor

    Public RigaCliccata As Long
    Dim ColNumerazione As Boolean
    Dim Dizio As String
    Dim CurrentRow As Long
    Dim CurrentCol As Long
    Dim TopRow As Long
    Dim LeftCol As Long
    Dim SchermoWidth As Long
    Dim SchermoHeight As Long
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim PrevCol As Long
    Dim CurrentFont As Font
    Dim WidthCharacter As Single
    Dim HeightRow As Long
    Dim ShiftState As Boolean = False
    Dim CurrentGraphics As Graphics
    Dim CurrentX As Long
    Dim CurrentY As Long
    Dim CurrentC As Brush

    Dim CodeRows() As StructMonRow

    Dim BorderX As Single
    Dim BorderY As Single

    Dim ColRem As Brush = Brushes.Green
    Dim ColId As Brush = Brushes.Maroon
    Dim ColConst As Brush = Brushes.Black
    Dim ColFunct As Brush = Brushes.Navy

    Dim SospendiRefresh As Boolean

    Public Shadows Event DblClick(Riga As Long, Numerazione As Boolean)
    Public Shadows Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim PrevSelettoreRiga As Long

    Dim CarNumeri As Long
    Dim LargBordo As Long

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim bm As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height, Drawing.Imaging.PixelFormat.Format24bppRgb)
        CurrentGraphics = Graphics.FromImage(bm)
        CurrentFont = New Font("Courier New", 10)

        WidthCharacter = pctSchermo.CreateGraphics().MeasureString("XXXXXXXXXX", CurrentFont).Width / 10
        HeightRow = 18

        pctSchermo.Image = bm

        CurrentGraphics.Clear(pctSchermo.BackColor)

        On Error Resume Next

        Dizio = ""
        Dizio = Dizio + ",And,Call,Case,Const,Dim,Watch,Inside,Stop,Do,Each"
        Dizio = Dizio + ",Else,ElseIf,Empty,End,Eqv,Err,Error,Exit,Explicit"
        Dizio = Dizio + ",False,For,Function,Goto,If,Imp,In,Is,Let,Like,Loop"
        Dizio = Dizio + ",New,Next,Not,Nothing,Null,On,Option,Or,Private"
        Dizio = Dizio + ",Public,Rem,Resume,Select,Set,Step,Sub,Then,To,True"
        Dizio = Dizio + ",Until,Wend,While,With,Abs,Asc,Array,Atn,Chr,Cos,Date"
        Dizio = Dizio + ",Erase,Exp,Filter,Fix,Hex,Hour,Int,Join,Left,Len,Log"
        Dizio = Dizio + ",Mid,Minute,Now,Oct,Randomize,Replace,Right,Rnd,Round"
        Dizio = Dizio + ",Sgn,Second,Sin,Space,Split,Sqr,String,Tan,Time,Trim"
        Dizio = Dizio + ",Weekday,Year,Month,Day,CBool,CByte,CCur,CDate,CDbl,CInt"
        Dizio = Dizio + ",CLng,CSng,CStr,LBound,LCase,LTrim,RTrim,UBound,UCase"
        Dizio = Dizio + ",AscB,AscW,ChrB,ChrW,CreateObject,DateAdd,DateDiff,DatePart"
        Dizio = Dizio + ",DateSerial,DateValue,FormatCurrency,FormatDateTime,FormatNumber"
        Dizio = Dizio + ",FormatPercent,GetObject,InputBox,InStr,InStrB,InStrRev"
        Dizio = Dizio + ",IsArray,IsDate,IsEmpty,IsNull,IsNumeric,IsObject,LeftB"
        Dizio = Dizio + ",LenB,LoadPicture,MidB,MsgBox,ReDim,RGB,RightB,StrComp"
        Dizio = Dizio + ",StrReverse,TimeSerial,TimeValue,TypeName,VarType,WeekdayName"
        Dizio = Dizio + ","

        InizializzaContenuto()

        ShapeStatement.Parent = pctSchermo
        ShapeStatement.BackColor = Color.Transparent
        ShapeStatement.Left = -1
        ShapeStatement.Top = BorderY
        ShapeStatement.Width = My.Computer.Screen.Bounds.Width

        ShapeSelettore.Parent = pctSchermo
        ShapeSelettore.BackColor = Color.Transparent
        ShapeSelettore.Left = -1
        ShapeSelettore.Top = BorderY
        ShapeSelettore.Width = My.Computer.Screen.Bounds.Width

    End Sub

    Private Sub InizializzaContenuto()

        MaxRow = 0
        MaxCol = 0
        CurrentRow = 0
        CurrentCol = 0
        PrevCol = 0
        LeftCol = 0
        TopRow = 0
        CurrentC = ColConst

        BorderX = 2
        BorderY = 2
        ReDim CodeRows(1023)
        ShapeStatement.Width = pctSchermo.Width
        ShapeStatement.Height = HeightRow
        PosizionaSelettore()

    End Sub

    Private Sub pctSchermo_DblClick() Handles pctSchermo.DoubleClick

        On Error Resume Next

        If RigaCliccata >= 0 Then
            RaiseEvent DblClick(RigaCliccata, ColNumerazione)
        End If

    End Sub

    Private Sub pctSchermo_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pctSchermo.MouseDown

        Try

            ClickedRow(e.Y)

            ColNumerazione = CLng(e.X \ WidthCharacter) <= LargBordo

            DisponiSelettore(e.Y)

            RaiseEvent MouseDown(e.Button, ShiftState, e.X, e.Y)

        Catch ex As Exception

            RigaCliccata = -1

        End Try

    End Sub

    Private Function ClickedRow(Y As Single)

        Dim Riga As String

        RigaCliccata = TopRow + Y \ HeightRow

        If RigaCliccata > MaxRow Then
            RigaCliccata = MaxRow
        End If

        If Trim(CodeRows(RigaCliccata).Rg) = "" Then

            RigaCliccata = -1

        Else

            Riga = Trim(CodeRows(RigaCliccata).Rg)

            If Riga.ToLower.StartsWith("'") Then
                RigaCliccata = -1
            ElseIf Riga.ToLower.StartsWith("rem ") Then
                RigaCliccata = -1
            ElseIf Riga.ToLower.StartsWith("dim ") Then
                RigaCliccata = -1
            ElseIf Riga.ToLower.StartsWith("watch ") Then
                RigaCliccata = -1
            ElseIf Riga.ToLower.StartsWith("inside ") Then
                RigaCliccata = -1
            Else
                RigaCliccata = RigaCliccata + 1
            End If

        End If

        Return RigaCliccata

    End Function

    Private Sub pctSchermo_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pctSchermo.MouseMove

        DisponiSelettore(e.Y)

    End Sub

    Private Sub pctSchermo_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles pctSchermo.MouseUp

        DisponiSelettore(e.Y)

    End Sub

    Public Sub PrintRow(Row As Long)

        Dim Parola As String = ""
        Dim Car As String
        Dim InizioParola As Long
        Dim I As Long
        Dim OnRem As Boolean = False
        Dim OnString As Boolean = False
        Dim ObjId As String = ""
        Dim Separatore As String = ""

        CurrentX = BorderX + CLng((CarNumeri + 2) * WidthCharacter)
        CurrentY = BorderY + (Row - TopRow) * HeightRow
        CurrentC = ColConst

        InizioParola = 0


        If MaxRow > 0 And Row < MaxRow Then
            CurrentGraphics.DrawString((Row + 1).ToString().PadLeft(CarNumeri + 1), CurrentFont, CurrentC, BorderX, CurrentY)
        End If

        For I = 1 To Len(CodeRows(Row).Rg)

            Car = Mid(CodeRows(Row).Rg, I, 1)

            If OnRem Then

                Parola = Parola + Car

            ElseIf OnString Then

                If Car = Chr(34) Then

                    Parola = Parola + Car
                    PrintWord(Parola, ObjId, OnRem, OnString, Separatore, InizioParola)
                    OnString = False

                Else

                    Parola = Parola + Car

                End If

            Else

                If Car = "'" Then

                    PrintWord(Parola, ObjId, OnRem, OnString, Separatore, InizioParola)
                    OnRem = True
                    InizioParola = I
                    Parola = "'"

                ElseIf Car = Chr(34) Then

                    PrintWord(Parola, ObjId, OnRem, OnString, Separatore, InizioParola)
                    OnString = True
                    InizioParola = I
                    Parola = Chr(34)

                Else

                    Select Case Car

                        Case "0" To "9", "a" To "z", "A" To "Z", "_"
                            Parola = Parola + Car
                            If InizioParola = 0 Then
                                InizioParola = I
                            End If

                        Case Else
                            Separatore = Car
                            PrintWord(Parola, ObjId, OnRem, OnString, Separatore, InizioParola)

                            If I - LeftCol > SchermoWidth + 2 Then
                                Exit For
                            ElseIf I > LeftCol Then
                                CurrentGraphics.DrawString(Car, CurrentFont, CurrentC, CurrentX, CurrentY)
                                CurrentX = CLng(CurrentX + WidthCharacter)
                            End If

                    End Select

                End If

            End If

        Next I

        PrintWord(Parola, ObjId, OnRem, OnString, Separatore, InizioParola)

        CurrentGraphics.DrawString(Space(SchermoWidth), CurrentFont, ColConst, BorderX, CurrentY)

        If MaxRow > 0 Then
            If CodeRows(Row).Bk Then
                CurrentGraphics.DrawImage(ImageList1.Images(0), BorderX, 2 * BorderY + (Row - TopRow) * HeightRow)
            End If
        End If

    End Sub

    Public Sub PrintWord(
        ByRef Parola As String,
        ByRef ObjId As String,
        ByRef OnRem As Boolean,
        ByRef OnString As Boolean,
        ByRef Separatore As String,
        ByRef InizioParola As Long
    )

        Dim Colore As Brush = ColConst
        Dim Normale As String = ""
        Dim Buff As String = ""

        If InizioParola > 0 Then

            If OnRem Then
                Colore = ColRem
            ElseIf OnString Then
                Colore = ColConst
            ElseIf Riconoscimento(Parola, Colore, Normale) Then
                Colore = Colore
                Parola = Normale
            ElseIf IsNumeric(Parola) Then
                Colore = ColConst
            Else
                Colore = ColId
            End If

            If InizioParola + Len(Parola) > LeftCol Then
                If InizioParola > LeftCol Then
                    Buff = Parola
                Else
                    Buff = Mid(Parola, LeftCol - InizioParola + 2)
                End If
                CurrentGraphics.DrawString(Buff, CurrentFont, Colore, CurrentX, CurrentY)
                CurrentX = CLng(CurrentX + Len(Buff) * WidthCharacter)
            End If

            Colore = ColConst

            If Separatore = "." Then
                ObjId = Parola
            Else
                ObjId = ""
            End If

            Parola = ""
            InizioParola = 0

        Else

            ObjId = ""

            Colore = ColConst

        End If

    End Sub

    Private Sub PosizionaSelettore()

        If 0 <= (CurrentRow - TopRow) And (CurrentRow - TopRow) <= SchermoHeight + 2 Then

            'If LeftCol = 0 Then
            ShapeStatement.Left = 0
            ShapeStatement.Top = (CurrentRow - TopRow) * HeightRow + BorderY
            ShapeStatement.Width = pctSchermo.Width
            ShapeStatement.Height = HeightRow
            'Else
            'ShapeStatement.Left = -BorderX
            'ShapeStatement.Top = (CurrentRow - TopRow) * HeightRow + BorderY
            'ShapeStatement.Width = pctSchermo.Width + 2 * BorderY
            'ShapeStatement.Height = HeightRow
            'End If

            ShapeStatement.Visible = True

        Else

            ShapeStatement.Visible = False

        End If

        If ScrollOrizzontale.Value <> LeftCol Then
            If ScrollOrizzontale.Maximum < LeftCol Then
                MaxCol = CurrentCol
                ScrollOrizzontale.Maximum = LeftCol
            End If
            ScrollOrizzontale.Value = LeftCol
        End If

        If ScrollVerticale.Value <> TopRow Then
            If TopRow <= ScrollVerticale.Maximum Then
                ScrollVerticale.Value = TopRow
            Else
                ScrollVerticale.Value = ScrollVerticale.Maximum
            End If
        End If

    End Sub

    Public Sub monRefresh()

        On Error GoTo ErrorProcedure

        Dim R As Long

        ShapeSelettore.Visible = False

        If SospendiRefresh = False Then

            CurrentGraphics.Clear(pctSchermo.BackColor)

            For R = TopRow To TopRow + SchermoHeight
                PrintRow(R)
            Next R

            SospendiRefresh = True
            PosizionaSelettore()
            SospendiRefresh = False

            pctSchermo.Refresh()

        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

        SospendiRefresh = False

    End Sub

    Private Sub SetMaxScrollRow()

        If MaxRow + 2 > SchermoHeight Then
            ScrollVerticale.Maximum = MaxRow - SchermoHeight + 2
            ScrollVerticale.Enabled = True
        ElseIf ScrollVerticale.Enabled Then
            ScrollVerticale.Maximum = 0
            ScrollVerticale.Enabled = False
        End If

    End Sub

    Private Sub SettaMaxScrollCol()

        If MaxCol + 2 > SchermoWidth Then
            ScrollOrizzontale.Maximum = MaxCol - SchermoWidth + 2
            ScrollOrizzontale.Enabled = True
        ElseIf ScrollOrizzontale.Enabled Then
            ScrollOrizzontale.Maximum = 0
            ScrollOrizzontale.Enabled = False
        End If

    End Sub

    Private Function Riconoscimento(Parola As String, ByRef Colore As Brush, ByRef Normale As String) As Boolean

        Dim Ret As Boolean
        Dim Posiz As Long

        Ret = False
        Colore = ColConst
        Normale = Parola

        Posiz = InStr(LCase(Dizio), "," + LCase(Parola) + ",")

        If Posiz > 0 Then

            Ret = True
            Colore = ColFunct
            Normale = Mid(Dizio, Posiz + 1, Len(Parola))

        End If

        Return Ret

    End Function

    Public Overrides Property Text() As String

        Get

            On Error GoTo ErrorProcedure

            Dim I As Long

            Dim CopiaRighe(MaxRow) As String

            If MaxRow = 0 And CodeRows(0).Rg = "" Then

                Text = ""

            Else

                For I = 0 To MaxRow
                    CopiaRighe(I) = CodeRows(I).Rg
                Next I

                Text = Join(CopiaRighe, vbCrLf)

            End If

            Exit Property

ErrorProcedure:

            Resume AbortProcedure

AbortProcedure:

            Text = ""

        End Get

        Set(value As String)

            On Error GoTo ErrorProcedure

            Dim NewLine As String
            Dim BuffRows As Object
            Dim I As Long

            InizializzaContenuto()

            If InStr(value, vbCrLf) > 0 Then
                NewLine = vbCrLf
            ElseIf InStr(value, Chr(10)) > 0 Then
                NewLine = Chr(10)
            Else
                NewLine = vbCrLf
            End If

            '------------------------
            ' Scomposizione in righe
            '------------------------

            If Not value.EndsWith(NewLine) Then
                value = value + NewLine
            End If

            BuffRows = Split(value, NewLine)

            MaxCol = 0
            MaxRow = UBound(BuffRows)
            CarNumeri = Len("" & (MaxRow + 1))
            LargBordo = CarNumeri + 3
            SchermoWidth = CLng(pctSchermo.Width / WidthCharacter) - LargBordo

            ReDim CodeRows(MaxRow + 1024)

            For I = 0 To MaxRow

                CodeRows(I).Rg = BuffRows(I)
                ParseRow(I, False)

                If MaxCol < Len(CodeRows(I).Rg) Then
                    MaxCol = Len(CodeRows(I).Rg)
                End If

            Next I

            SetMaxScrollRow()
            SettaMaxScrollCol()

            monRefresh()

            Exit Property

ErrorProcedure:

            Resume AbortProcedure

AbortProcedure:

            On Error Resume Next

            SetMaxScrollRow()
            SettaMaxScrollCol()

            monRefresh()

        End Set

    End Property

    Private Sub ParseRow(Riga As Long, FlagStampa As Boolean)

        Dim Parola As String = ""
        Dim Car As String
        Dim PrevCar As String
        Dim InizioParola As Long
        Dim I As Long
        Dim OnRem As Boolean
        Dim OnString As Boolean
        Dim Colore As Brush = ColConst
        Dim Normale As String = ""
        Dim ObjId As String = ""
        Dim Separatore As String = ""
        Dim FlagInizio As Boolean
        Dim SuccPos As Long = 0
        Dim OnIf As Boolean
        Dim OnThen As Boolean
        Dim OnFor As Boolean
        Dim OnTo As Boolean
        Dim OnNext As Boolean
        Dim OnEnd As Boolean
        Dim OnExit As Boolean

        InizioParola = 0
        FlagInizio = False
        I = 0

        Do

            I = I + 1

            If I <= Len(CodeRows(Riga).Rg) Then

                If I > 1 Then
                    PrevCar = Mid(CodeRows(Riga).Rg, I - 1, 1)
                Else
                    PrevCar = ""
                End If

                Car = Mid(CodeRows(Riga).Rg, I, 1)

                If Car < " " Then

                    Mid(CodeRows(Riga).Rg, I) = " "
                    Car = " "

                End If

                If OnRem Then

                    Parola = Parola + Car

                ElseIf OnString Then

                    If Car = Chr(34) Then

                        Parola = Parola + Car
                        ParseWord(Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola)
                        OnString = False

                    Else

                        Parola = Parola + Car

                    End If

                Else

                    If Car = "'" Then

                        ParseWord(Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola)
                        OnRem = True
                        InizioParola = I
                        Parola = "'"

                    ElseIf Car = Chr(34) Then

                        ParseWord(Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola)
                        OnString = True
                        InizioParola = I
                        Parola = Chr(34)

                    Else

                        Select Case Car

                            Case "0" To "9", "a" To "z", "A" To "Z", "_"

                                FlagInizio = True
                                Parola = Parola + Car

                                If InizioParola = 0 Then
                                    InizioParola = I
                                End If

                            Case Else

                                Separatore = Car
                                ParseWord(Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola)

                        End Select

                    End If

                End If

            Else

                Exit Do

            End If

        Loop

        ParseWord(Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola)

        If FlagStampa Then
            PrintRow(Riga)
        End If

    End Sub

    Public Sub ParseWord(
    Riga As Long,
    ByRef Parola As String,
    ByRef ObjId As String,
    ByRef OnRem As Boolean,
    ByRef OnString As Boolean,
    ByRef OnIf As Boolean,
    ByRef OnThen As Boolean,
    ByRef OnFor As Boolean,
    ByRef OnTo As Boolean,
    ByRef OnNext As Boolean,
    ByRef OnEnd As Boolean,
    ByRef OnExit As Boolean,
    ByRef Separatore As String,
    ByRef InizioParola As Long
)

        Dim Colore As Brush = ColConst
        Dim Normale As String = ""

        If InizioParola > 0 Then

            If OnRem Or OnString Then

            ElseIf Riconoscimento(Parola, Colore, Normale) Then

                Parola = Normale

                Mid(CodeRows(Riga).Rg, InizioParola) = Parola

                If Colore.Equals(ColFunct) Then

                    If Parola = "If" Then
                        OnIf = True
                    End If

                    If Parola = "Then" Then
                        OnThen = True
                    End If

                    If Parola = "For" Then
                        OnFor = True
                    End If

                    If Parola = "To" Or Parola = "Each" Then
                        OnTo = True
                    End If

                    If Parola = "Next" Then
                        OnNext = True
                    End If

                    If Parola = "End" Then
                        OnEnd = True
                    End If

                    If Parola = "Exit" Then
                        OnExit = True
                    End If

                End If

            End If

            If Separatore = "." Then
                ObjId = Parola
            Else
                ObjId = ""
            End If

            Parola = ""
            InizioParola = 0

        Else

            ObjId = ""

        End If

    End Sub

    Public Property Riga() As Long

        Get

            Return CurrentRow

        End Get

        Set(value As Long)

            CurrentRow = value

        End Set

    End Property

    Public Sub SetTopRow(ByVal NewValue As Long)

        On Error GoTo ErrorProcedure

        If NewValue < ScrollVerticale.Maximum Then
            TopRow = NewValue
        Else
            TopRow = ScrollVerticale.Maximum
        End If

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Public ReadOnly Property VisibleRows() As Long

        Get

            Return SchermoHeight

        End Get

    End Property


    Private Sub DisponiSelettore(Y As Single)

        On Error Resume Next

        Dim Rg As Long

        Rg = Y \ HeightRow

        If PrevSelettoreRiga <> Rg Then

            If Rg < MaxRow - TopRow Then

                If LeftCol = 0 Then
                    ShapeSelettore.Left = -1
                    ShapeSelettore.Top = Rg * HeightRow + BorderY
                    'ShapeSelettore.Width = pctSchermo.Width
                    'ShapeSelettore.Height = HeightRow
                Else
                    ShapeSelettore.Left = -1
                    ShapeSelettore.Top = Rg * HeightRow + BorderY
                    'ShapeSelettore.Width = pctSchermo.Width + 2 * BorderY
                    'ShapeSelettore.Height = HeightRow
                End If


                If ShapeSelettore.Visible = False Then
                    ShapeSelettore.Visible = True
                End If

            Else

                ShapeSelettore.Visible = False

            End If

        End If

        PrevSelettoreRiga = Rg

    End Sub

    Public Property BreakPoint(Riga As Long) As Boolean

        Get

            Return CodeRows(Riga - 1).Bk

        End Get

        Set(value As Boolean)

            CodeRows(Riga - 1).Bk = value

        End Set

    End Property

    Public Sub ClearBreakPoint()

        On Error Resume Next

        Dim I As Long

        For I = 0 To MaxRow
            CodeRows(I).Bk = False
        Next I

        monRefresh()

    End Sub

    Public Function ListBreakPoint() As String

        On Error Resume Next

        Dim I As Long
        Dim Buffer As String = ""

        For I = 0 To MaxRow
            If CodeRows(I).Bk Then
                Buffer = Buffer & (I + 1) & "|"
            End If
        Next I

        If Buffer <> "" Then
            Buffer = "|" & Buffer
        End If

        ListBreakPoint = Buffer

    End Function

    Public ReadOnly Property GetRow(Riga As Long) As String

        Get

            On Error Resume Next

            Return CodeRows(Riga - 1).Rg

        End Get

    End Property

    Private Sub pctSchermo_MouseEnter(sender As System.Object, e As System.EventArgs) Handles pctSchermo.MouseEnter

        pctSchermo.Focus()

    End Sub

    Private Sub pctSchermo_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctSchermo.MouseWheel

        On Error GoTo ErrorProcedure

        If e.Delta > 0 Then
            If TopRow > 5 Then
                TopRow = TopRow - 5
            Else
                TopRow = 0
            End If
        Else
            If TopRow + 5 < ScrollVerticale.Maximum Then
                TopRow = TopRow + 5
            Else
                TopRow = ScrollVerticale.Maximum
            End If
        End If

        PosizionaSelettore()
        monRefresh()

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Private Sub monitor_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        On Error GoTo ErrorProcedure

        ShiftState = e.Shift

        Select Case e.KeyCode

            Case System.Windows.Forms.Keys.Left
                If LeftCol - 20 > 0 Then
                    LeftCol = LeftCol - 20
                Else
                    LeftCol = 0
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.Right
                If LeftCol + 20 <= ScrollOrizzontale.Maximum Then
                    LeftCol = LeftCol + 20
                Else
                    LeftCol = ScrollOrizzontale.Maximum
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.Up
                If TopRow - 1 > 0 Then
                    TopRow = TopRow - 1
                Else
                    TopRow = 0
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.Down
                If TopRow + 1 <= ScrollVerticale.Maximum Then
                    TopRow = TopRow + 1
                Else
                    TopRow = ScrollVerticale.Maximum
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.PageUp
                If TopRow - SchermoHeight + 1 > 0 Then
                    TopRow = TopRow - SchermoHeight + 1
                Else
                    TopRow = 0
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.PageDown
                If TopRow + SchermoHeight - 1 <= ScrollVerticale.Maximum Then
                    TopRow = TopRow + SchermoHeight - 1
                Else
                    TopRow = ScrollVerticale.Maximum
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.Home
                If ShiftState Then
                    TopRow = 0
                Else
                    LeftCol = 0
                End If
                monRefresh()

            Case System.Windows.Forms.Keys.End
                If ShiftState Then
                    TopRow = ScrollVerticale.Maximum
                Else
                    LeftCol = ScrollOrizzontale.Maximum
                End If
                monRefresh()

        End Select

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Private Sub monitor_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp

        ShiftState = False

    End Sub

    Private Sub ScrollOrizzontale_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles ScrollOrizzontale.Scroll

        LeftCol = ScrollOrizzontale.Value
        PosizionaSelettore()
        monRefresh()

    End Sub

    Private Sub ScrollOrizzontale_Enter(sender As Object, e As EventArgs) Handles ScrollOrizzontale.Enter

        pctSchermo.Focus()

    End Sub

    Private Sub ScrollVerticale_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles ScrollVerticale.Scroll

        TopRow = ScrollVerticale.Value
        monRefresh()

    End Sub

    Private Sub ScrollVerticale_Enter(sender As Object, e As EventArgs) Handles ScrollVerticale.Enter

        pctSchermo.Focus()

    End Sub

    Private Sub monitor_Enter(sender As Object, e As EventArgs) Handles MyBase.Enter

        PosizionaSelettore()

    End Sub

    Private Sub monitor_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave

        ParseRow(CurrentRow, True)

    End Sub

    Private Sub monitor_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        On Error GoTo ErrorProcedure

        ScrollVerticale.Left = Me.Width - ScrollVerticale.Width - 4
        ScrollVerticale.Top = 0
        ScrollVerticale.Height = Me.Height - ScrollOrizzontale.Height - 4

        ScrollOrizzontale.Left = 0
        ScrollOrizzontale.Top = Me.Height - ScrollOrizzontale.Height - 4
        ScrollOrizzontale.Width = Me.Width - ScrollVerticale.Width - 4

        pctSchermo.Left = 0
        pctSchermo.Top = 0
        pctSchermo.Width = ScrollVerticale.Left
        pctSchermo.Height = ScrollOrizzontale.Top

        SchermoWidth = CLng(pctSchermo.Width / WidthCharacter) - LargBordo
        SettaMaxScrollCol()

        SchermoHeight = pctSchermo.Height \ HeightRow
        SetMaxScrollRow()

        ScrollVerticale.LargeChange = SchermoHeight - 1

        monRefresh()

        Exit Sub

ErrorProcedure:

        Resume AbortProcedure

AbortProcedure:

    End Sub

    Private Sub ShapeStatement_MouseMove(sender As Object, e As MouseEventArgs) Handles ShapeStatement.MouseMove

        DisponiSelettore(ShapeStatement.Top)

    End Sub

    Private Sub ShapeSelettore_MouseDown(sender As Object, e As MouseEventArgs) Handles ShapeSelettore.MouseDown

        Try

            ClickedRow(ShapeSelettore.Top + e.Y)

            ColNumerazione = CLng(e.X \ WidthCharacter) <= LargBordo

        Catch ex As Exception

            RigaCliccata = -1

        End Try

    End Sub

    Private Sub ShapeSelettore_DoubleClick(sender As Object, e As EventArgs) Handles ShapeSelettore.DoubleClick

        On Error Resume Next

        If RigaCliccata >= 0 Then
            RaiseEvent DblClick(RigaCliccata, ColNumerazione)
        End If

    End Sub

    Private Sub ShapeStatement_DoubleClick(sender As Object, e As EventArgs) Handles ShapeStatement.DoubleClick

        On Error Resume Next

        If RigaCliccata >= 0 Then
            RaiseEvent DblClick(RigaCliccata, ColNumerazione)
        End If

    End Sub

    Private Sub ShapeStatement_MouseDown(sender As Object, e As MouseEventArgs) Handles ShapeStatement.MouseDown

        Try

            ClickedRow(ShapeSelettore.Top + e.Y)

            ColNumerazione = CLng(e.X \ WidthCharacter) <= LargBordo

        Catch ex As Exception

            RigaCliccata = -1

        End Try

    End Sub

End Class
