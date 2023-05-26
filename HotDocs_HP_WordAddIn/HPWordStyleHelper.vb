Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports HPW = HotDocs_HP_WordAddIn.HPWordHelper

Module HPWordStyleHelper
    Dim ThisApplication As Application = Globals.ThisAddIn.Application
    'Dim ActiveDoc As Document = ThisApplication.ActiveDocument

    Public Sub CommentCounter()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        MsgBox("Megjegyzések száma: " & ActiveDoc.Comments.Count)
    End Sub
    Friend Sub EnumListStyles()
        Dim ThisListTemplates As ListTemplates = Globals.ThisAddIn.Application.ActiveDocument.ListTemplates
        MsgBox(ThisListTemplates.Count)
        Dim Nevek As New StringBuilder
        For Each lst As ListTemplate In ThisListTemplates
            Nevek.AppendLine(lst.Name & " szintszám: " & lst.ListLevels.Count & "[OLNumbered: " & lst.OutlineNumbered & "]")
        Next
        MsgBox("Listák: " & Nevek.ToString)
        'TODO: törlések?
        'TODO: paragrafus alapú számozások?
    End Sub

    Public Sub FileAndCommentsCompare()
        Dim Válasz, Köztes, Régi, Új As Integer
        Dim Counter As Integer = 1
        Dim ÚjDoc As Word.Document
        Dim DocOld
        Dim DocNew

        If ThisApplication.Documents.Count > 2 Or ThisApplication.Documents.Count < 2 Then
            MsgBox("Két összehasonlítandó dokumentum legyen nyitva, ne több, ne kevesebb!")
            Exit Sub
        End If
        Régi = 1 : Új = 2
        If ThisApplication.Documents.Count = 2 Then
            Do
                Válasz = MsgBox("Régi: " & ThisApplication.Documents(Régi).Name & "; Új: " & ThisApplication.Documents(Új).Name & ". Helyes?", vbYesNoCancel)
                If Válasz = vbNo Then
                    Köztes = Régi
                    Régi = Új
                    Új = Köztes
                End If
                If Válasz = vbCancel Then Exit Sub
            Loop Until Válasz = vbYes

            DocOld = ThisApplication.Documents(Régi)
            DocNew = ThisApplication.Documents(Új)
            Dim TempFileNameBase As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "Compare_" & DateTime.Now.ToString("yyyyMMddHHmm")
            Dim TempFileNameWCounter As String = TempFileNameBase & "_" & CStr(Counter)
            Counter = Counter + 1
            ÚjDoc = ThisApplication.CompareDocuments(DocOld, DocNew, Destination:=WdCompareDestination.wdCompareDestinationNew, Granularity:=WdGranularity.wdGranularityCharLevel, CompareFormatting:=False, CompareCaseChanges:=False, CompareWhitespace:=False, CompareFields:=False, IgnoreAllComparisonWarnings:=True, RevisedAuthor:="Comparator")
            If ÚjDoc.Comments.Count > 0 Then

                WriteCommentsToFile(DocOld, TempFileNameWCounter & ".txt")
                DocOld.Close
                Dim NewDocName = TempFileNameBase & CStr(Counter) & ".txt"
                WriteCommentsToFile(DocNew, NewDocName)
                DocNew.Close
                Counter = Counter + 1
                CompareComments(TempFileNameWCounter & ".txt", NewDocName)
                ÚjDoc.DeleteAllComments()
            End If

            ÚjDoc.SaveAs2(TempFileNameBase & "_" & Counter & "_Compared.docx", CompatibilityMode:=WdCompatibilityMode.wdCurrent)
            ÚjDoc.Close()

        End If
    End Sub

    Public Sub WriteAllCommToFile()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        Dim TempFileNameBase As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "Comments_" & DateTime.Now.ToString("yyyyMMddHHmm")
        WriteCommentsToFile(ActiveDoc, TempFileNameBase & ".txt")
    End Sub
    Public Sub WriteNewComments()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        Dim fromWhen As String
        Dim convDate As DateTime
        Dim TempFileNameBase As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "Comments" & DateTime.Now.ToString("yyyyMMddHHmm")

        fromWhen = InputBox("Mikortóli kommenteket írjuk ki?", , DateAdd("d", -5, Now()))
        Try
            convDate = CType(fromWhen, DateTime)
        Catch ex As Exception
            MsgBox("Rossz a dátumformátum, művelet megszakítva. Helyes formátum: " & Now.ToLocalTime)
            Exit Sub
        End Try
        WriteCommentsToFile(ActiveDoc, TempFileNameBase & ".txt")
    End Sub


    Public Sub ApplyHeadingLS05()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("LS_05")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("LS_05"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(0.5)
    End Sub
    Public Sub ApplyHeadingLS063()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("LS_063")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("LS_063"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(0.63)
    End Sub
    Public Sub ApplyHeadingLS125()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("LS_125")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("LS_125"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(1.25)
    End Sub
    Public Sub ApplyHeadingLS15() '*
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("LS_15")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("LS_15"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(1.5)
    End Sub
    Public Sub ApplyHeadingSLS05()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("Stepped_LS_05")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("Stepped_LS_05"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(0.5)
    End Sub
    Public Sub ApplyHeadingSLS063()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("Stepped_LS_063")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("Stepped_LS_063"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(0.63)
    End Sub
    Public Sub ApplyHeadingSLS125()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("Stepped_LS_125")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("Stepped_LS_125"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(1.25)
    End Sub
    Public Sub ApplyHeadingSLS15()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        CheckListStyle("Stepped_LS_15")
        ApplyHeading1_9_ListStyle(ActiveDoc.Styles("Stepped_LS_15"))
        ActiveDoc.DefaultTabStop = ThisApplication.CentimetersToPoints(1.5)
    End Sub
    Public Sub CreateAndApplyMLevelList()
        Dim MyForm As New frmCreateAndApplyListStyle
        MyForm.ShowDialog()
    End Sub
    Public Sub Lvl1_NameNumberedBodyTextLvl()
        NameNumberedBodyTextLvl(1)
    End Sub
    Public Sub Lvl2_NameNumberedBodyTextLvl()
        NameNumberedBodyTextLvl(2)
    End Sub
    Public Sub Lvl3_NameNumberedBodyTextLvl()
        NameNumberedBodyTextLvl(3)
    End Sub
    Public Sub LvlInputNameNumberedBodyTextLvl()
        Dim Lvl As Integer
        Dim Result = InputBox("Hanyadik szinten legyen a számozott törzsszöveg?", "Számozott törzsszöveg outline szint meghatározása")
        Try
            Lvl = CInt(Result)
        Catch ex As Exception
            MsgBox("Nem adott meg értéket vagy nem számot adott meg.")
            Exit Sub
        End Try
        NameNumberedBodyTextLvl(Lvl)
    End Sub
    Public Sub LvlCurrentNameNumberedBodyTextLvl()
        Dim Lvl As Integer
        Lvl = ThisApplication.Selection.ParagraphFormat.OutlineLevel
        If Lvl = 10 Then Lvl = 1
        NameNumberedBodyTextLvl(Lvl)
    End Sub

    Public Sub CreateDefaultListStyles()
        'Dim sIndents(1 To 3) As Single
        'Dim elem As Object
        Dim mList As String
        Dim Indent As Single
        HPWordHelper.SetCursorToWaiting()
        'sIndents(1) = 0.5
        'sIndents(2) = 1.5
        'sIndents(3) = 1.25
        'sIndents(4) = 0.63

        Indent = CheckIndent()
        If Indent = 0 Then
            HPWordHelper.SetCursorToDefault()
            Exit Sub
        End If
        mList = Replace(CStr(Indent), ThisApplication.International(WdInternationalIndex.wdDecimalSeparator), "")

        'Create non-stepped list style
        CreateMlevelListStyle("LS_" & mList, Indent, False)
        'Create stepped list styles
        CreateMlevelListStyle("Stepped_LS_" & mList, Indent, True)

        CreateBulletListStyleSL("BULL_" & mList, Indent, Chr(149))
        CreateBulletListStyleSL("HYPHEN_" & mList, Indent, "-")
        CreateBulletListStyleSL("EQUAL_" & mList, Indent, "=")

        CreateNumberingSL("ArabwClosingParenthesis_" & mList, Indent, , "", ")")
        CreateNumberingSL("ArabwFullParenthesis_" & mList, Indent, , "(", ")")
        CreateNumberingSL("ArabwDot_" & mList, Indent, WdListNumberStyle.wdListNumberStyleOrdinal, "")

        CreateNumberingSL("UpperRomanwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "", ")")
        CreateNumberingSL("UpperRomanwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "(", ")")
        CreateNumberingSL("UpperRomanwDot_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "", ".")

        CreateNumberingSL("LowerRomanwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "", ")")
        CreateNumberingSL("LowerRomanwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "(", ")")
        CreateNumberingSL("LowerRomanwDot_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "", ".")

        CreateNumberingSL("UpperLetterwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "", ")")
        CreateNumberingSL("UpperLetterwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "(", ")")
        CreateNumberingSL("UpperLetterwDot_" & mList, Indent, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "", ".")

        CreateNumberingSL("LowerLetterwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "", ")")
        CreateNumberingSL("LowerLetterwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "(", ")")
        CreateNumberingSL("LowerLetterwDot_" & mList, Indent, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "", ".")

        CreateNumberingSL("CardinalText_" & mList, Indent, WdListNumberStyle.wdListNumberStyleCardinalText, "", " ")
        CreateNumberingSL("CardinalTextwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleCardinalText, "", ")")
        CreateNumberingSL("CardinalTextwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleCardinalText, "(", ")")
        CreateNumberingSL("CardinalTextwColon_" & mList, Indent, WdListNumberStyle.wdListNumberStyleCardinalText, "", ":")

        CreateNumberingSL("OrdinalText_" & mList, Indent, WdListNumberStyle.wdListNumberStyleOrdinalText, "", " ")
        CreateNumberingSL("OrdinalTextwClosingParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleOrdinalText, "", ")")
        CreateNumberingSL("OrdinalTextwFullParenthesis_" & mList, Indent, WdListNumberStyle.wdListNumberStyleOrdinalText, "(", ")")
        CreateNumberingSL("OrdinalTextwColon_" & mList, Indent, WdListNumberStyle.wdListNumberStyleOrdinalText, "", ":")
        HPWordHelper.SetCursorToDefault()
    End Sub
    Public Sub ApplyBULL()
        ApplyBulletorNumberingStyle("BULL", CheckIndentStr)
    End Sub
    Public Sub ApplyHYPHEN()
        ApplyBulletorNumberingStyle("HYPHEN", CheckIndentStr)
    End Sub
    Public Sub ApplyEQUAL()
        ApplyBulletorNumberingStyle("EQUAL", CheckIndentStr)
    End Sub
    Public Sub ApplyArabwClosingParenthesis()
        ApplyBulletorNumberingStyle("ArabwClosingParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyArabwFullParenthesis()
        ApplyBulletorNumberingStyle("ArabwFullParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyUpperRomanwDot()
        ApplyBulletorNumberingStyle("UpperRomanwDot", CheckIndentStr)
    End Sub
    Public Sub ApplyUpperRomanClosingParenthesis()
        ApplyBulletorNumberingStyle("UpperRomanwClosingParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyUpperRomanwFullParenthesis()
        ApplyBulletorNumberingStyle("UpperRomanwFullParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyLowerRomanwFullParenthesis()
        ApplyBulletorNumberingStyle("LowerRomanwFullParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyUpperLetterwClosingParenthesis()
        ApplyBulletorNumberingStyle("UpperLetterwClosingParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyUpperLetterwFullParenthesis()
        ApplyBulletorNumberingStyle("UpperLetterwFullParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyLowerLetterwClosingParenthesis()
        ApplyBulletorNumberingStyle("LowerLetterwClosingParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyLowerLetterwFullClosingParenthesis()
        ApplyBulletorNumberingStyle("LowerLetterwFullParenthesis", CheckIndentStr)
    End Sub
    Public Sub ApplyCardinalText()
        ApplyBulletorNumberingStyle("CardinalText", CheckIndentStr)
    End Sub
    Public Sub ApplyOrdinalText()
        ApplyBulletorNumberingStyle("OrdinalText", CheckIndentStr)
    End Sub
    Public Sub ApplyBodyText()
        ThisApplication.Selection.Style = WdBuiltinStyle.wdStyleBodyText
    End Sub
    Public Sub ApplyBodyTextNumbered()
        ThisApplication.Selection.Style = WdBuiltinStyle.wdStyleBodyText2
    End Sub


    Public Sub CompareComments(FileOldC As String, FileNewC As String)
        Dim ÚjDoc2, CommentOld, CommentNew As Word.Document

        CommentOld = ThisApplication.Documents.Open(FileOldC)
        CommentNew = ThisApplication.Documents.Open(FileNewC)
        Dim TempFileNameBase As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "Compare_" & DateTime.Now.ToString("yyyyMMddHHmm")
        ÚjDoc2 = ThisApplication.CompareDocuments(CommentOld, CommentNew, RevisedAuthor:="Comparator")
        ÚjDoc2.SaveAs2(TempFileNameBase & "_" & "1" & "_Compared.docx", CompatibilityMode:=WdCompatibilityMode.wdCurrent)
        ÚjDoc2.Close()
        CommentOld.Close()
        CommentNew.Close()

    End Sub
    Public Sub WriteCommentsToFile(DocToWrite As Document, OutputFileName As String, Optional SinceWhen As DateTime = Nothing)
        Dim mComm As Comment
        Dim AlreadyExists As Boolean = False
        Dim MyStrB As New StringBuilder

        For Each mComm In DocToWrite.Comments
            If IsNothing(SinceWhen) Then
                MyStrB.AppendLine("On " & mComm.Date.ToString & " " & mComm.Contact.ToString & " said: " & mComm.Range.Text)
            Else
                If mComm.Date >= SinceWhen Then
                    MyStrB.AppendLine("On " & mComm.Date.ToString & " " & mComm.Contact.ToString & " said: " & mComm.Range.Text)
                End If
            End If
        Next mComm
        If File.Exists(OutputFileName) Then AlreadyExists = True
        Try
            File.WriteAllText(OutputFileName, MyStrB.ToString)
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
        If AlreadyExists Then MsgBox(OutputFileName & " felülírva!")
    End Sub


    Public Function CheckIndent() As Single
        Dim mList As String
        Dim valasz
        Dim Result As Single = 0
        mList = HPWordHelper.CVarValue("ListUsed")
        If mList = "" Then
            valasz = MsgBox("Main list style not defined in DocVariable. Shall I create a default 1.5 cm LS?", MsgBoxStyle.OkCancel)
            If valasz = vbOK Then
                ApplyHeadingLS15()
                mList = "LS_15"
            Else
                Return 0
            End If
        End If
        mList = Right(mList, Len(mList) - InStr(1, mList, "_LS_") - Len("_LS_") + 1)
        Select Case mList
            Case "05"
                Result = 0.5
            Case "063"
                Result = 0.63
            Case "125"
                Result = 1.25
            Case "15"
                Result = 1.5
        End Select
        Return Result
    End Function

    Public Function CheckIndentStr() As String
        Dim Indent As Single
        Dim TempStr As String

        Indent = CheckIndent()
        TempStr = Replace(CStr(Indent), ThisApplication.International(WdInternationalIndex.wdDecimalSeparator), "")
        If Indent = 0 Then TempStr = "05"
        CheckIndentStr = TempStr
    End Function


    Public Sub CheckListStyle(ListToUse As String)
        Dim Indent As Single
        Dim Stepped As Boolean
        Dim IndentStr As String


        If StyleExists(ListToUse) = False Then
            If InStr(1, ListToUse, "Stepped_") <> 0 Then Stepped = True
            IndentStr = Right(ListToUse, Len(ListToUse) - InStr(1, ListToUse, "_LS_") - Len("_LS_") + 1)
            Select Case IndentStr
                Case "05"
                    Indent = 0.5
                Case "063"
                    Indent = 0.63
                Case "125"
                    Indent = 1.25
                Case "15"
                    Indent = 1.5
            End Select
            CreateMlevelListStyle(ListToUse, Indent, Stepped)
        End If

    End Sub
    Public Sub ApplyHeading1_9_ListStyle(ListToUse As Word.Style)
        Dim i As Integer
        For i = 1 To 9
            Globals.ThisAddIn.Application.ActiveDocument.Styles(ListToUse).ListTemplate.ListLevels(i).LinkedStyle = "Heading " & CStr(i)
        Next i
        HPWordHelper.CVarToFill("ListUsed", ListToUse)
    End Sub

    Public Sub DeleteHeading1_9_ListStyle(ListToUse As Word.Style)
        Dim i As Integer
        For i = 1 To 9
            Globals.ThisAddIn.Application.ActiveDocument.Styles(ListToUse).ListTemplate.ListLevels(i).LinkedStyle = ""
        Next i
        HPWordHelper.CVarToFill("ListUsed", "")
    End Sub

    Public Sub NameNumberedBodyTextLvl(Optional FirstNumberedBodyTextLevel As Integer = 1)
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        'Lvl=Milyen szintre tesszük az első SZÁMOZOTT body text 2 vagy body text 3-at, plusz alá a többi wdStyleListNumber[-5]-t?
        'Fölé headingeket
        Dim NameofListStyleUsedForHeadingList As String
        Dim i, MaximumLevelOfListNumberToSet As Integer

        NameofListStyleUsedForHeadingList = HPWordHelper.CVarValue("ListUsed")
        If NameofListStyleUsedForHeadingList = "" Then
            MsgBox("List Style Used Not Found in DocVariable. Create Heading List First!")
            Exit Sub
        End If

        ApplyHeading1_9_ListStyle(ActiveDoc.Styles(NameofListStyleUsedForHeadingList))
        If FirstNumberedBodyTextLevel > 9 Then FirstNumberedBodyTextLevel = 9
        ActiveDoc.Styles(NameofListStyleUsedForHeadingList).ListTemplate.ListLevels(FirstNumberedBodyTextLevel).LinkedStyle =
            ActiveDoc.Styles(WdBuiltinStyle.wdStyleBodyText2).NameLocal
        If FirstNumberedBodyTextLevel < 9 Then
            If FirstNumberedBodyTextLevel <= 4 Then MaximumLevelOfListNumberToSet = 5 Else MaximumLevelOfListNumberToSet = 9 - FirstNumberedBodyTextLevel '[4.3.2.1][5.6.7.8]
            For i = 1 To MaximumLevelOfListNumberToSet
                If i = 1 Then
                    ActiveDoc.Styles(NameofListStyleUsedForHeadingList).ListTemplate.ListLevels(FirstNumberedBodyTextLevel + 1).LinkedStyle =
                        ActiveDoc.Styles(WdBuiltinStyle.wdStyleListNumber).NameLocal
                Else
                    ActiveDoc.Styles(NameofListStyleUsedForHeadingList).ListTemplate.ListLevels(FirstNumberedBodyTextLevel + i).LinkedStyle =
                        ActiveDoc.Styles("List Number " & CStr(i)).NameLocal
                End If
            Next i
        End If
    End Sub

    Public Sub CreateMlevelListStyle(NameStyle As String, IndentWidth As Single, Stepped As Boolean, Optional sNumberStyle As Integer = 5, Optional bSeparator As String = "", Optional cSeparator As String = "")
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        Dim i As Integer
        Dim NFormat As String = String.Empty

        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)

        If StyleExists(NameStyle) = False Then ActiveDoc.Styles.Add(Name:=NameStyle, Type:=WdStyleType.wdStyleTypeList)

        For i = 1 To 9
            NFormat = NFormat & bSeparator & "%" & i & cSeparator
            With ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1).ListLevels(i)
                .NumberFormat = NFormat
                .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                .NumberStyle = sNumberStyle
                .Alignment = WdListLevelAlignment.wdListLevelAlignLeft

                If Stepped = True Then
                    .NumberPosition = ThisApplication.CentimetersToPoints((i - 1) * IndentWidth)
                    .TextPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)
                    .TabPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)
                Else
                    .NumberPosition = ThisApplication.CentimetersToPoints(0)
                    .TextPosition = ThisApplication.CentimetersToPoints(IndentWidth)
                    .TabPosition = ThisApplication.CentimetersToPoints(IndentWidth)
                End If

                .ResetOnHigher = (i - 1)
                .StartAt = 1
                With .Font
                    .Bold = WdConstants.wdUndefined
                    .Italic = WdConstants.wdUndefined
                    .StrikeThrough = WdConstants.wdUndefined
                    .Subscript = WdConstants.wdUndefined
                    .Superscript = WdConstants.wdUndefined
                    .Shadow = WdConstants.wdUndefined
                    .Outline = WdConstants.wdUndefined
                    .Emboss = WdConstants.wdUndefined
                    .Engrave = WdConstants.wdUndefined
                    .AllCaps = WdConstants.wdUndefined
                    .Hidden = WdConstants.wdUndefined
                    .Underline = WdConstants.wdUndefined
                    .Color = WdConstants.wdUndefined
                    .Size = WdConstants.wdUndefined
                    .Animation = WdConstants.wdUndefined
                    .DoubleStrikeThrough = WdConstants.wdUndefined
                    .Name = ""
                End With
                .LinkedStyle = ""
            End With
        Next i

        ActiveDoc.Styles(NameStyle).LinkToListTemplate(ListTemplate:=ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1), ListLevelNumber:=1)
        ActiveDoc.Styles(NameStyle).BaseStyle = ""
        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)

    End Sub

    Public Sub CreateBulletListStyleSL(NameStyle As String, IndentWidth As Single, Optional nBullet As String = "-")
        Dim i As Integer
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)
        ActiveDoc.Styles.Add(Name:=NameStyle, Type:=WdStyleType.wdStyleTypeList)
        For i = 1 To 9
            'NFormat = NFormat & "%" & i & cSeparator
            With ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1).ListLevels(i)
                .NumberFormat = nBullet
                .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                '.NumberStyle = sNumberStyle
                .Alignment = WdListLevelAlignment.wdListLevelAlignLeft

                .NumberPosition = ThisApplication.CentimetersToPoints((i - 1) * IndentWidth)
                .TextPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)
                .TabPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)

                '.ResetOnHigher = (i - 1)
                .StartAt = 1
                With .Font
                    .Bold = WdConstants.wdUndefined
                    .Italic = WdConstants.wdUndefined
                    .StrikeThrough = WdConstants.wdUndefined
                    .Subscript = WdConstants.wdUndefined
                    .Superscript = WdConstants.wdUndefined
                    .Shadow = WdConstants.wdUndefined
                    .Outline = WdConstants.wdUndefined
                    .Emboss = WdConstants.wdUndefined
                    .Engrave = WdConstants.wdUndefined
                    .AllCaps = WdConstants.wdUndefined
                    .Hidden = WdConstants.wdUndefined
                    .Underline = WdConstants.wdUndefined
                    .Color = WdConstants.wdUndefined
                    .Size = WdConstants.wdUndefined
                    .Animation = WdConstants.wdUndefined
                    .DoubleStrikeThrough = WdConstants.wdUndefined
                    .Name = ""
                End With
                .LinkedStyle = ""
            End With
        Next i

        ActiveDoc.Styles(NameStyle).LinkToListTemplate(ListTemplate:=ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1), ListLevelNumber:=1)
        ActiveDoc.Styles(NameStyle).BaseStyle = ""
        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)
    End Sub

    Public Sub CreateNumberingSL(NameStyle As String, IndentWidth As Single, Optional sNumberStyle As Integer = 5, Optional bSeparator As String = "", Optional cSeparator As String = "")
        Dim i As Integer
        HPW.SetCursorToWaiting()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)

        ActiveDoc.Styles.Add(Name:=NameStyle, Type:=WdStyleType.wdStyleTypeList)
        For i = 1 To 9
            With ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1).ListLevels(i)
                .NumberFormat = bSeparator & "%" & i & cSeparator
                .TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                .NumberStyle = sNumberStyle
                .Alignment = WdListLevelAlignment.wdListLevelAlignLeft

                .NumberPosition = ThisApplication.CentimetersToPoints((i - 1) * IndentWidth)
                .TextPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)
                .TabPosition = ThisApplication.CentimetersToPoints(i * IndentWidth)

                .ResetOnHigher = (i - 1)
                .StartAt = 1
                With .Font
                    .Bold = WdConstants.wdUndefined
                    .Italic = WdConstants.wdUndefined
                    .StrikeThrough = WdConstants.wdUndefined
                    .Subscript = WdConstants.wdUndefined
                    .Superscript = WdConstants.wdUndefined
                    .Shadow = WdConstants.wdUndefined
                    .Outline = WdConstants.wdUndefined
                    .Emboss = WdConstants.wdUndefined
                    .Engrave = WdConstants.wdUndefined
                    .AllCaps = WdConstants.wdUndefined
                    .Hidden = WdConstants.wdUndefined
                    .Underline = WdConstants.wdUndefined
                    .Color = WdConstants.wdUndefined
                    .Size = WdConstants.wdUndefined
                    .Animation = WdConstants.wdUndefined
                    .DoubleStrikeThrough = WdConstants.wdUndefined
                    .Name = ""
                End With
                .LinkedStyle = ""
            End With
        Next i
        HPWordHelper.SetCursorToDefault()
        ActiveDoc.Styles(NameStyle).LinkToListTemplate(ListTemplate:=ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1), ListLevelNumber:=1)
        ActiveDoc.Styles(NameStyle).BaseStyle = ""
        ThisApplication.ListGalleries(WdListGalleryType.wdOutlineNumberGallery).Reset(1)
    End Sub


    Public Sub ApplyBulletorNumberingStyle(ListStyle As String, mList As String)
        Dim WhereAreWe As Single
        If Not StyleExists(ListStyle & "_" & mList) Then
            Select Case ListStyle
                Case "BULL"
                    CreateBulletListStyleSL(ListStyle & "_" & mList, CSng(mList) / 100, Chr(149))
                Case "HYPHEN"
                    CreateBulletListStyleSL("HYPHEN_" & mList, CSng(mList) / 100, "-")
                Case "EQUAL"
                    CreateBulletListStyleSL("EQUAL_" & mList, CSng(mList) / 100, "=")
                Case "ArabwClosingParenthesis"
                    CreateNumberingSL("ArabwClosingParenthesis_" & mList, CSng(mList) / 100, , "", ")")
                Case "ArabwFullParenthesis"
                    CreateNumberingSL("ArabwFullParenthesis_" & mList, CSng(mList) / 100, , "(", ")")
                Case "ArabwDot"
                    CreateNumberingSL("ArabwDot_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleOrdinal, "")
                Case "UpperRomanwClosingParenthesis"
                    CreateNumberingSL("UpperRomanwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "", ")")
                Case "UpperRomanwFullParenthesis"
                    CreateNumberingSL("UpperRomanwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "(", ")")
                Case "UpperRomanwDot"
                    CreateNumberingSL("UpperRomanwDot_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseRoman, "", ".")
                Case "LowerRomanwClosingParenthesis"
                    CreateNumberingSL("LowerRomanwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "", ")")
                Case "LowerRomanwFullParenthesis"
                    CreateNumberingSL("LowerRomanwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "(", ")")
                Case "LowerRomanwDot"
                    CreateNumberingSL("LowerRomanwDot_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseRoman, "", ".")
                Case "UpperLetterwClosingParenthesis"
                    CreateNumberingSL("UpperLetterwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "", ")")
                Case "UpperLetterwFullParenthesis"
                    CreateNumberingSL("UpperLetterwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "(", ")")
                Case "UpperLetterwDot"
                    CreateNumberingSL("UpperLetterwDot_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleUppercaseLetter, "", ".")
                Case "LowerLetterwClosingParenthesis"
                    CreateNumberingSL("LowerLetterwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "", ")")
                Case "LowerLetterwFullParenthesis"
                    CreateNumberingSL("LowerLetterwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "(", ")")
                Case "LowerLetterwDot"
                    CreateNumberingSL("LowerLetterwDot_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleLowercaseLetter, "", ".")
                Case "CardinalText"
                    CreateNumberingSL("CardinalText_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleCardinalText, "", " ")
                Case "CardinalTextwClosingParenthesis"
                    CreateNumberingSL("CardinalTextwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleCardinalText, "", ")")
                Case "CardinalTextwFullParenthesis"
                    CreateNumberingSL("CardinalTextwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleCardinalText, "(", ")")
                Case "CardinalTextwColon"
                    CreateNumberingSL("CardinalTextwColon_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleCardinalText, "", ":")
                Case "OrdinalText"
                    CreateNumberingSL("OrdinalText_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleOrdinalText, "", " ")
                Case "OrdinalTextwClosingParenthesis"
                    CreateNumberingSL("OrdinalTextwClosingParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleOrdinalText, "", ")")
                Case "OrdinalTextwFullParenthesis"
                    CreateNumberingSL("OrdinalTextwFullParenthesis_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleOrdinalText, "(", ")")
                Case "OrdinalTextwColon"
                    CreateNumberingSL("OrdinalTextwColon_" & mList, CSng(mList) / 100, WdListNumberStyle.wdListNumberStyleOrdinalText, "", ":")
            End Select
        End If

        WhereAreWe = ThisApplication.Selection.ParagraphFormat.LeftIndent
        If WhereAreWe = 9999999 Then WhereAreWe = 0 Else WhereAreWe = Math.Round(WhereAreWe / ThisApplication.ActiveDocument.DefaultTabStop)
        ThisApplication.Selection.Style = ThisApplication.ActiveDocument.Styles(WdBuiltinStyle.wdStyleBodyText)
        ThisApplication.Selection.Style = ThisApplication.ActiveDocument.Styles(ListStyle & "_" & mList)

        'If WhereAreWe >= 1 Then
        'For i = 1 To WhereAreWe
        'Globals.ThisAddIn.Application.Selection.Paragraphs.Indent()
        'Next i
        'End If
    End Sub

    Public Function StyleExists(NameStyle As String) As Boolean
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        Dim mList As Word.Style
        Dim Result As Boolean = False
        For Each mList In ActiveDoc.Styles
            If mList.NameLocal = NameStyle Then Result = True
        Next mList
        Return Result
    End Function


    Public Sub ReplaceAllBodyText3ToSame2()
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        For Each ContentParagraph As Paragraph In ActiveDoc.Paragraphs
            If ContentParagraph.Style = WdBuiltinStyle.wdStyleBodyText3 Then
                ContentParagraph.Style = WdBuiltinStyle.wdStyleBodyText 'Hiba: Microsoft Wordben (COM): operator = not defined for Style for type WdBuiltinStyle
                ContentParagraph.Style = WdBuiltinStyle.wdStyleBodyText3
            End If
        Next
        Globals.ThisAddIn.logger.Info("ReplaceAllBodyText3ToSame2 vége")
    End Sub
    Public Sub ReplaceAllBodyText3ToSame()
        Globals.ThisAddIn.logger.Info("ReplaceAllBodyText3ToSame kezdete")
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim MyRange = ActiveDoc.StoryRanges(WdStoryType.wdMainTextStory)
        ReplaceAllParagraphStyles(MyRange, OldStyle:=Word.WdBuiltinStyle.wdStyleBodyText3, NewStyle:=Word.WdBuiltinStyle.wdStyleEnvelopeReturn)
        ReplaceAllParagraphStyles(MyRange, OldStyle:=Word.WdBuiltinStyle.wdStyleEnvelopeReturn, NewStyle:=Word.WdBuiltinStyle.wdStyleBodyText3)
        Globals.ThisAddIn.logger.Info("ReplaceAllBodyText3ToSame vége")
    End Sub
    Public Sub ReplaceAllParagraphStyles(RangeToChange As Range, OldStyle As Integer, NewStyle As Integer)
        Dim ActiveDoc As Document = ThisApplication.ActiveDocument
        With RangeToChange.Find
            .Text = ""
            .Replacement.Text = ""
            .ClearFormatting()
            .Style = ActiveDoc.Styles(OldStyle)
            .Replacement.Style = ActiveDoc.Styles(NewStyle)
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute(Replace:=Word.WdReplace.wdReplaceAll)
        End With
    End Sub
End Module
