Imports NHunspell

Public Class NLP


    Public Shared Function DeclinationSymbols(TermToChange As String, SampleTerm As String) As List(Of String)
        Dim Result As New List(Of String)
        Dim Symbols As String = "%€$£¥"
        If Not Symbols.Contains(TermToChange.Last) Then Return Result
        Select Case SampleTerm
            Case "Megrendelőt"
                If TermToChange.Last = "%" Then
                    Result.Add(TermToChange + "-át")
                ElseIf TermToChange.Last = "£" Then
                    Result.Add(TermToChange + "-ot")
                Else
                    Result.Add(TermToChange + "-t")
                End If
            Case "Megrendelőn"
                If "%$£".Contains(TermToChange.Last) Then
                    Result.Add(TermToChange + "-on")
                ElseIf TermToChange.Last = "€" Then
                    Result.Add(TermToChange + "-n")
                ElseIf TermToChange.Last = "¥" Then
                    Result.Add(TermToChange + "-en")
                End If
            Case "Megrendelővel"
                Select Case TermToChange.Last
                    Case "%"
                        Result.Add(TermToChange + "-kal")
                    Case "€"
                        Result.Add(TermToChange + "-val")
                    Case "$"
                        Result.Add(TermToChange + "-ral")
                    Case "£"
                        Result.Add(TermToChange + "-tal")
                    Case "¥"
                        Result.Add(TermToChange + "-nel")
                End Select
            Case "Megrendelőjével"
                Select Case TermToChange.Last
                    Case "%"
                        Result.Add(TermToChange + "-ával")
                    Case "€", "$", "£"
                        Result.Add(TermToChange + "-jával")
                    Case "¥"
                        Result.Add(TermToChange + "-jével")
                End Select
            Case "Megrendelőben"
                If MélyHangrendűSymbol(TermToChange.Last) Then
                    Result.Add(TermToChange + "-ban")
                Else
                    Result.Add(TermToChange + "-ben")
                End If
            Case "Megrendelőjében"
                If MélyHangrendűSymbol(TermToChange.Last) Then
                    Result.Add(TermToChange + "-jában")
                Else
                    Result.Add(TermToChange + "-jében")
                End If
            Case "Megrendelőnek"
                If MélyHangrendűSymbol(TermToChange.Last) Then
                    Result.Add(TermToChange + "-nak")
                Else
                    Result.Add(TermToChange + "-nek")
                End If
            Case "Megrendelőre"
                If MélyHangrendűSymbol(TermToChange.Last) Then
                    Result.Add(TermToChange + "-ra")
                Else
                    Result.Add(TermToChange + "-re")
                End If
        End Select
        Return Result
    End Function
    Private Shared Function MélyHangrendűSymbol(Input) As Boolean
        If "%€$£".Contains(Input) Then Return True Else Return False
    End Function

    Friend Shared Function GenerateWord(TermToChange As String, SampleTerm As String) As String
        Dim Result As New List(Of String)
        Dim SymbolResult As List(Of String) = DeclinationSymbols(TermToChange, SampleTerm)
        If SymbolResult.Count > 0 Then Result.AddRange(SymbolResult)
        Dim HunSpellResult As List(Of String) = GenerateWordFromHunspell(TermToChange, SampleTerm)
        If HunSpellResult.Count > 0 Then Result.AddRange(HunSpellResult)

        If Result.Count = 0 Then Return String.Empty
        Dim _resultItem = Result.First
        _resultItem = CheckExceptionsforHunspell(_resultItem)
        Return _resultItem
    End Function
    Private Shared Function GenerateWordFromHunspell(TermToChange As String, SampleTerm As String) As List(Of String)
        Dim Result As New List(Of String)
        Using MyHunspell As New Hunspell(Globals.ThisAddIn.DictionaryPath & "HU_hu.aff", Globals.ThisAddIn.DictionaryPath & "HU_hu.dic")
            For Each stem As String In MyHunspell.Generate(TermToChange, SampleTerm)
                Result.Add(stem)
            Next
        End Using
        Return Result
    End Function
    Private Shared Function CheckExceptionsforHunspell(Input As String) As String
        If Globals.ThisAddIn.ExceptionDictionary.ContainsKey(Input) Then
            Dim ReturnValue As String = Globals.ThisAddIn.ExceptionDictionary(Input)
            Return ReturnValue
        End If
        Return Input
    End Function
End Class
