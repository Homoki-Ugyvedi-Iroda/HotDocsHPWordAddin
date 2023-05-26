Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Public Class HPWordHelper
    Const CustomContractXMLNameSpace = "http://www.homoki.net/uri/customXMLcontracts/"
    Public Const AllCapStyleName = "_AllCap"
    Shared VoltTrack As Boolean

    Friend Shared Function ValidateFileName(Input As String) As String
        If String.IsNullOrWhiteSpace(Input) Then Return String.Empty
        For Each c In Path.GetInvalidFileNameChars()
            Input = Input.Replace(c, "").Trim
        Next
        Return Input
    End Function
    Public Shared Sub FindReplaceAll(ToReplace As String, ReplaceWith As String, Optional IgnoreCase As Boolean = True, Optional WildCard As Boolean = True)
        FindReplaceInternal(ToReplace, ReplaceWith, WildCard)
        If IgnoreCase = True Then FindReplaceInternal(UppercaseSecondLetter(ToReplace), UppercaseFirstLetter(ReplaceWith), WildCard)
    End Sub
    Public Shared Sub HivatkozasRaKeres(Mit As String)
        Dim Found As Boolean
        Dim ParagraphRangeToSearch As Word.Range
        Globals.ThisAddIn.Application.ActiveDocument.StoryRanges(Word.WdStoryType.wdMainTextStory).Select()
        Globals.ThisAddIn.Application.Selection.Find.ClearFormatting()
        Found = True
        Do While Found = True
            Found = Globals.ThisAddIn.Application.Selection.Find.Execute(FindText:=Mit, Forward:=True, Format:=False, MatchCase:=False)
            If Found = True Then
                ParagraphRangeToSearch = Globals.ThisAddIn.Application.ActiveDocument.Range(Globals.ThisAddIn.Application.Selection.Range.Start + 1, Globals.ThisAddIn.Application.Selection.Range.Start + 3)
                If Mit.Contains("z") Then
                    ParagraphRangeToSearch.Text = Replace(ParagraphRangeToSearch.Text, "az", "a", , , vbBinaryCompare)
                    ParagraphRangeToSearch.Text = Replace(ParagraphRangeToSearch.Text, "Az", "A", , , vbBinaryCompare)
                Else
                    ParagraphRangeToSearch.Text = Replace(ParagraphRangeToSearch.Text, "a ", "az ", , , vbBinaryCompare)
                    ParagraphRangeToSearch.Text = Replace(ParagraphRangeToSearch.Text, "A ", "Az ", , , vbBinaryCompare)
                End If
            End If
        Loop
    End Sub

    Private Shared Sub FindReplaceInternal(ToReplace As String, ReplaceWith As String, Optional WildCard As Boolean = True)
        Dim ActiveDocRange = Globals.ThisAddIn.Application.ActiveDocument.Range
        ActiveDocRange.Find.ClearFormatting()
        ActiveDocRange.Find.Replacement.ClearFormatting()
        With ActiveDocRange.Find
            .Text = ToReplace
            .Replacement.Text = ReplaceWith
            .Forward = True
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            If WildCard = True Then .MatchWildcards = True
        End With
        ActiveDocRange.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
    End Sub
    Public Shared Function LowerCaseFirstLetter(Input As String) As String
        Dim Kezdobetu As String = Input.First
        Kezdobetu = Kezdobetu.ToLower
        Dim KiskezdobetusDef As String = Kezdobetu & Strings.Mid(Input, 2)
        Return KiskezdobetusDef
    End Function
    Public Shared Function UppercaseFirstLetter(ByVal val As String) As String
        If String.IsNullOrEmpty(val) Then
            Return val
        End If
        Dim array() As Char = val.ToCharArray
        array(0) = Char.ToUpper(array(0))
        Return New String(array)
    End Function
    Public Shared Function UppercaseSecondLetter(ByVal val As String) As String
        If String.IsNullOrEmpty(val) Then
            Return val
        End If
        Dim array() As Char = val.ToCharArray
        array(1) = Char.ToUpper(array(1))
        Return New String(array)
    End Function

    Friend Shared Sub MellékletHivatkozásTörlése()
        Globals.ThisAddIn.logger.Info("MellékletHivatkozásTörlése kezdete")
        CheckTrack()
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim rng = ActiveDoc.Content
        Dim TitoktartasiFelsorolas As Bookmark = Nothing
        Dim MellekletEsTitoktartasiFelsorolas As Bookmark = Nothing
        For Each oBM As Bookmark In ActiveDoc.Bookmarks
            If oBM.Name = "SzerződésTitoktartásiFelsorolás" Then TitoktartasiFelsorolas = oBM
            If oBM.Name = "SzerződésMellékletTitoktartásiFelsorolás" Then MellekletEsTitoktartasiFelsorolas = oBM
        Next
        If IsNothing(TitoktartasiFelsorolas) OrElse IsNothing(MellekletEsTitoktartasiFelsorolas) Then Exit Sub

        rng.Find.MatchCase = True
        rng.Find.Text = "No table of contents entries found"
        rng.Find.Execute()

        While rng.Find.Found
            If IsNothing(TitoktartasiFelsorolas) And IsNothing(MellekletEsTitoktartasiFelsorolas) Then Exit While
            If Not IsNothing(TitoktartasiFelsorolas) AndAlso Not IsNothing(TitoktartasiFelsorolas.Range) AndAlso rng.InRange(TitoktartasiFelsorolas.Range) Then
                TitoktartasiFelsorolas.Range.Text = String.Empty
                Globals.ThisAddIn.logger.Info("találtam TitoktartasiFelsorolas")
                UpdateAllFields(NoWarning:=True)
                TitoktartasiFelsorolas = Nothing
            ElseIf Not IsNothing(MellekletEsTitoktartasiFelsorolas) AndAlso Not IsNothing(MellekletEsTitoktartasiFelsorolas.Range) AndAlso rng.InRange(MellekletEsTitoktartasiFelsorolas.Range) Then
                MellekletEsTitoktartasiFelsorolas.Range.Text = String.Empty
                UpdateAllFields(NoWarning:=True)
                Globals.ThisAddIn.logger.Info("találtam MellekletEsTitoktartasiFelsorolas")
                MellekletEsTitoktartasiFelsorolas = Nothing
            End If
            rng.Find.Execute()
        End While
        RestoreTrack()
        Globals.ThisAddIn.logger.Info("MellékletHivatkozásTörlése vége")
    End Sub
    Public Shared Function MezoEllenorzes(Optional NoWarning As Boolean = False) As Boolean
        Dim rngStory As Range         ' Range Object for Looping through Stories
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument

        ' Loop through each story in doc to update
        For Each rngStory In ActiveDoc.StoryRanges
            If rngStory.Find.Execute("Error!") = True Or rngStory.Find.Execute("Hiba!") = True Then
                If MsgBox("Hiba a hivatkozásokban. Kattintson az OK-ra, ha ez nem zavarja a jóváhagyásban.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
            If rngStory.Find.Execute("[") = True Or rngStory.Find.Execute("]") = True Then
                If MsgBox("Szögletes zárójel a szövegben. Kattintson az OK-ra, ha ez nem zavarja a jóváhagyásban.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
            If rngStory.Find.Execute("***") = True Then
                If MsgBox("*** a szövegben. Kattintson az OK-ra, ha ez nem zavarja a jóváhagyásban.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
            If rngStory.Find.Execute("´") = True Then
                If MsgBox("ragozást jelölő ` a szövegben. Kattintson az OK-ra, ha ez nem zavarja a jóváhagyásban.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
            If rngStory.Find.Execute(" 0. pont") = True Then
                If MsgBox("0. pontra hivatkozás a szövegben. Kattintson az OK-ra, ha ez nem zavarja a jóváhagyásban.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
            If rngStory.Find.Execute("No table of contents entries found") Then
                If MsgBox("Tartalomjegyzékben nem megfelelő elválasztó szerepel.", MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkCancel, "Mezőellenőrzés") = vbCancel Then MezoEllenorzes = True
            End If
        Next
        Globals.ThisAddIn.logger.Info("MezoEllenorzes vége: MezoEllenorzes=" & MezoEllenorzes)
        Return MezoEllenorzes
    End Function

    Public Shared Sub UpdateAllFields(Optional NoWarning As Boolean = False)
        '---------------------------------------------------------------------------------------
        ' Procedure: sUpdateFields (V2)
        ' DateTime : 20-Dec-2001
        ' Updated  : 06-Nov-2002 - Update fields in Text Boxes
        ' Updated  : Last update unnecessary (from Word 2003 onward at least)
        ' Author   : Bryan Carbonnell
        '            With code by Nancy Hutson Hale
        ' Purpose  : To update all fields in the Word Document including TOC, TOA, TOF,
        '             fields in text boxes and fields in headers/footers
        '---------------------------------------------------------------------------------------
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Dim wnd As Window = ActiveDoc.ActiveWindow           ' Pointer to Document's Window
        Dim lngMain As Long           ' Main Pane Type Holder
        Dim lngSplit As Long          ' Split Type Holder
        Dim lngActPane As Long        ' ActivePane Number
        Dim rngStory As Range         ' Range Object for Looping through Stories
        Dim TOC As TableOfContents    ' Table of Contents Object
        Dim TOA As TableOfAuthorities ' Table of Authorities Object
        Dim TOF As TableOfFigures     ' Table of Figures Object
        Dim shp As Microsoft.Office.Interop.Word.Shape              ' Shape Object to get Textboxes

        ' get Active Pane Number
        lngActPane = wnd.ActivePane.Index

        ' Hold View Type of Main pane
        lngMain = wnd.Panes(1).View.Type

        ' Hold SplitSpecial
        lngSplit = wnd.View.SplitSpecial
        ' Get Rid of any split
        wnd.View.SplitSpecial = Word.WdSpecialPane.wdPaneNone
        ' Set View to Normal
        wnd.View.Type = Word.WdViewType.wdNormalView
        ' Loop through TOC and update
        For Each TOC In ActiveDoc.TablesOfContents
            TOC.Update()
        Next
        ' Loop through TOA and update
        For Each TOA In ActiveDoc.TablesOfAuthorities
            TOA.Update()
        Next
        ' Loop through TOF and update
        For Each TOF In ActiveDoc.TablesOfFigures
            TOF.Update()
        Next

        ' Loop through each story in doc to update
        For Each rngStory In ActiveDoc.StoryRanges
            If rngStory.StoryType = WdStoryType.wdCommentsStory Then
                Globals.ThisAddIn.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone
                ' Update fields
                rngStory.Fields.Update()
                Globals.ThisAddIn.Application.DisplayAlerts = WdAlertLevel.wdAlertsAll
                rngStory.Fields.Update()
            Else : rngStory.Fields.Update()
            End If
            If rngStory.Find.Execute("Error!") = True Or rngStory.Find.Execute("Hiba!") = True Or rngStory.Find.Execute("No table of contents entries found") Then
                If NoWarning = False Then MsgBox("Hiba a hivatkozásokban.")
            End If
        Next

        'Loop through text boxes and update
        For Each shp In ActiveDoc.Shapes
            With shp.TextFrame
                If .HasText Then
                    shp.TextFrame.TextRange.Fields.Update()
                End If
            End With
        Next

        ' Return Split to original state
        wnd.View.SplitSpecial = lngSplit
        ' Return main pane to original state
        wnd.Panes(1).View.Type = lngMain

        ' Active proper pane
        wnd.Panes(lngActPane).Activate()
        Globals.ThisAddIn.logger.Info("UpdateAllFields vége")
    End Sub

    Public Shared Sub OwnWordSettingsHP()        '
        ' WordSettings
        Dim AppData As String = Environ("AppData")
        Dim TemplateWorkgroupDirectory As String = AppData & "\Microsoft\Templates\"
        Dim TemplateGlobalDirectory As String = AppData & "\Microsoft\Word\Startup\"
        'Csoportbeli CustomDictionary hozzáadása
        'Dim CustomDictionaryPath As String
        'With CustomDictionaries
        '       .ClearAll
        '       .Add(CustomDictionaryPathOwn).LanguageSpecific = False
        '       .ActiveCustomDictionary = CustomDictionaries.Item(CustomDictionaryPath)
        'Set CustomDictionaryPath = "C:\Documents and Settings\phomoki001\Application Data\Microsoft\Proof\CUSTOM.DIC"

        With Globals.ThisAddIn.Application
            .DisplayStatusBar = True
            .ShowWindowsInTaskbar = True
            .ShowStartupDialog = False
            .DisplayAutoCompleteTips = True
            .DisplayRecentFiles = True
            .DefaultSaveFormat = ""
        End With
        With Globals.ThisAddIn.Application.ActiveDocument
            .ShowGrammaticalErrors = True
            .ShowSpellingErrors = True
            .ClickAndTypeParagraphStyle = "Normal"
            .ReadOnlyRecommended = False
            .EmbedTrueTypeFonts = False
            .SaveFormsData = False
            .SaveSubsetFonts = False
            .DisableFeatures = False
            .EmbedSmartTags = True
            .EmbedTrueTypeFonts = False
            .DoNotEmbedSystemFonts = True
            .DisableFeatures = False
            .EmbedSmartTags = True
            .EmbedLinguisticData = True
        End With
        With Globals.ThisAddIn.Application.AutoCorrect
            .CorrectInitialCaps = False
            .CorrectSentenceCaps = False
            .CorrectDays = False
            .CorrectCapsLock = False
            .ReplaceText = True
            .ReplaceTextFromSpellingChecker = True
            .CorrectKeyboardSetting = False
            .DisplayAutoCorrectOptions = False
            .CorrectTableCells = False
        End With
        With Globals.ThisAddIn.Application.ActiveWindow
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayLeftScrollBar = False
            .StyleAreaWidth = Globals.ThisAddIn.Application.CentimetersToPoints(1.38)
            .DisplayVerticalRuler = True
            .DisplayRightRuler = False
            .DisplayScreenTips = True
            With .View
                .ShowAnimation = True
                .Draft = True
                .WrapToWindow = True
                .ShowPicturePlaceHolders = False
                .ShowFieldCodes = False
                .ShowBookmarks = True
                .FieldShading = WdFieldShading.wdFieldShadingAlways
                .ShowTabs = False
                .ShowSpaces = False
                .ShowParagraphs = False
                .ShowHyphens = False
                .ShowHiddenText = True
                .ShowAll = True
                .ShowDrawings = True
                .ShowObjectAnchors = True
                .ShowTextBoundaries = True
                .ShowHighlight = True
                .DisplayPageBoundaries = True
                .DisplaySmartTags = True
                .Type = WdViewType.wdPrintView
            End With
        End With
        With Globals.ThisAddIn.Application.Options
            .Pagination = True
            .WPHelp = False
            .WPDocNavKeys = False
            .ShortMenuNames = False
            .RTFInClipboard = True
            .EnableSound = False
            .ConfirmConversions = False
            .UpdateLinksAtOpen = True
            .SendMailAttach = True
            .MeasurementUnit = WdMeasurementUnits.wdCentimeters
            .AllowPixelUnits = False
            .AllowReadingMode = False
            .AnimateScreenMovements = True
            .InterpretHighAnsi = WdHighAnsiText.wdHighAnsiIsHighAnsi
            .BackgroundOpen = False
            .AutoCreateNewDrawings = False
            .ReplaceSelection = True
            .AllowDragAndDrop = False
            .INSKeyForPaste = True
            .CtrlClickHyperlinkToOpen = True
            .AutoKeyboardSwitching = False
            .DisplayPasteOptions = False
            .FormatScanning = False
            .ShowFormatError = True
            .SmartParaSelection = True
            .LocalNetworkFile = False
            .SaveInterval = 2
            .SaveNormalPrompt = True
            .DisableFeaturesbyDefault = False
            .AutoFormatAsYouTypeApplyHeadings = False
            .AutoFormatAsYouTypeApplyBorders = False
            .AutoFormatAsYouTypeApplyBulletedLists = False
            .AutoFormatAsYouTypeApplyNumberedLists = False
            .AutoFormatAsYouTypeApplyTables = False
            .AutoFormatAsYouTypeReplaceQuotes = False
            .AutoFormatAsYouTypeReplaceSymbols = False
            .AutoFormatAsYouTypeReplaceOrdinals = False
            .AutoFormatAsYouTypeReplaceFractions = False
            .AutoFormatAsYouTypeReplacePlainTextEmphasis = False
            .AutoFormatAsYouTypeReplaceHyperlinks = True
            .AutoFormatAsYouTypeFormatListItemBeginning = False
            .AutoFormatAsYouTypeDefineStyles = False
            .AutoFormatApplyHeadings = True
            .AutoFormatApplyLists = True
            .AutoFormatApplyBulletedLists = True
            .AutoFormatApplyOtherParas = True
            .AutoFormatReplaceQuotes = True
            .AutoFormatReplaceSymbols = True
            .AutoFormatReplaceOrdinals = True
            .AutoFormatReplaceFractions = True
            .AutoFormatReplacePlainTextEmphasis = True
            .AutoFormatReplaceHyperlinks = True
            .AutoFormatPreserveStyles = True
            .AutoFormatPlainTextWordMail = True
            .LabelSmartTags = False
            .DisplaySmartTagButtons = True
            '.DefaultFilePath(Path:=wdWorkgroupTemplatesPath) = TemplateWorkgroupDirectory
            '.DefaultFilePath(Path:=wdStartupPath) = TemplateGlobalDirectory
            .AutoWordSelection = True
            .INSKeyForPaste = True
            .PasteSmartCutPaste = True
            .AllowAccentedUppercase = False
            .TabIndentKey = False
            .Overtype = False
            .AllowClickAndTypeMouse = False
            .DisplayPasteOptions = False
            .PromptUpdateStyle = True
            .SmartCursoring = False
            .UpdateFieldsAtPrint = True
            .UpdateLinksAtPrint = False
            .DefaultTray = "Use printer settings"
            .PrintBackground = True
            .PrintProperties = False
            .PrintFieldCodes = False
            .PrintComments = False
            .PrintHiddenText = False
            .PrintXMLTag = False
            .PrintDrawingObjects = True
            .PrintDraft = False
            .PrintReverse = False
            .MapPaperSize = True
            .PrintOddPagesInAscendingOrder = False
            .PrintEvenPagesInAscendingOrder = False
            .PrintBackgrounds = False
            .AllowFastSave = False
            .BackgroundSave = True
            .CreateBackup = True
            .SavePropertiesPrompt = False
            .DisableFeaturesbyDefault = False
            .InsertedTextMark = WdInsertedTextMark.wdInsertedTextMarkUnderline
            .InsertedTextColor = WdColorIndex.wdByAuthor
            .DeletedTextMark = WdDeletedTextMark.wdDeletedTextMarkStrikeThrough
            .DeletedTextColor = WdColorIndex.wdByAuthor
            .RevisedPropertiesMark = WdRevisedPropertiesMark.wdRevisedPropertiesMarkNone
            .RevisedPropertiesColor = WdColorIndex.wdByAuthor
            .RevisedLinesMark = WdRevisedLinesMark.wdRevisedLinesMarkOutsideBorder
            .RevisedLinesColor = WdColorIndex.wdAuto
            .CommentsColor = WdColorIndex.wdByAuthor
            .RevisionsBalloonPrintOrientation = WdRevisionsBalloonPrintOrientation.wdBalloonPrintOrientationPreserve
            .ShowMarkupOpenSave = True
            .CheckSpellingAsYouType = False
            .CheckGrammarAsYouType = False
            .SuggestSpellingCorrections = False
            .SuggestFromMainDictionaryOnly = False
            .CheckGrammarWithSpelling = False
            .ShowReadabilityStatistics = False
            .IgnoreUppercase = False
            .IgnoreMixedDigits = False
            .IgnoreInternetAndFileAddresses = False
            .AllowCombinedAuxiliaryForms = True
            .EnableMisusedWordsDictionary = True
            .AllowCompoundNounProcessing = True
        End With
        Globals.ThisAddIn.Application.RecentFiles.Maximum = 9
        Globals.ThisAddIn.Application.Languages(WdLanguageID.wdHungarian).SpellingDictionaryType = WdDictionaryType.wdSpelling
    End Sub

    Public Shared Sub UnlinkAllReferences()
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Dim oField As Field
        'Dim oHeader As HeaderFooter
        Dim oSection As Section
        'Dim oFooter As HeaderFooter
        'Ahol az oldalszámozás van, ott nem célszerű kilinkelni, mert minden oldalon ugyanaz lesz az oldalszám?
        'Dim oShape As Shape
        Dim oMezoKod As String
        Dim MyTOC As TableOfContents

        For Each oSection In ActiveDoc.Sections
            For Each oField In oSection.Range.Fields
                If InStr(1, oField.Code.Text, "TOC", vbBinaryCompare) <> 0 And InStr(1, oField.Code.Text, "\h", vbTextCompare) <> 0 Then
                    oMezoKod = Replace(oField.Code.Text, "\h", "", , , vbTextCompare)
                    oMezoKod = Replace(oMezoKod, "\z", "", , , vbTextCompare)
                    oField.Code.Text = oMezoKod
                    For Each MyTOC In ActiveDoc.TablesOfContents
                        MyTOC.Update()
                    Next
                    oField.Unlink()
                Else : If Not (oField.Type = WdFieldType.wdFieldPage) Then oField.Unlink()
                End If
            Next oField
            If oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists And oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).LinkToPrevious = False Then
                For Each oField In oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range.Fields
                    If Not (oField.Type = WdFieldType.wdFieldPage) Then oField.Unlink()
                Next
            End If
            If oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Exists And oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = False Then
                For Each oField In oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Range.Fields
                    If Not (oField.Type = WdFieldType.wdFieldPage) Then oField.Unlink()
                Next
            End If
            If oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Exists And oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False Then
                For Each oField In oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Fields
                    If Not (oField.Type = WdFieldType.wdFieldPage) Then oField.Unlink()
                Next
            End If
        Next oSection
        'Az oldalszámokat nem jó unlinkelni, mert akkor mindenhol ugyanaz az oldalszám lesz. Ehelyett bonyolultabb megoldásra lesz szükség (beolvassa előbb oldalanként, és beírja, utána unlinkeli
        'Text Shape-ben vannak az oldalszámok sokszor!
    End Sub

    Public Shared Sub DefinitionsSort(Optional NoWarning As Boolean = False)
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Dim rngStart, rngEnd As Long 'Definíciós rész kezdete és vége
        Dim rngSortStart As Long 'Definíciós rész rendezendő kezdete (def. bevezető sorok miatt)
        Dim rngChange As Range
        Dim rngPara As Paragraph

        'a szöveg elejére ugrik
        ThisApplication.Selection.HomeKey(Unit:=WdUnits.wdStory)
        'meghatározzuk a fogalmi meghatározások sor kezdetét
        If Not KijelölTalálatStílussal({"fogalmak", "fogalom", "fogalmi", "definíció", "definition"}, ActiveDoc.Styles(WdBuiltinStyle.wdStyleHeading1)) Then
            If NoWarning = False Then MsgBox("Nem talált a program definíció/fogalommal kapcsolatos tartalmú címsort, ezért nem rendezi")
            Exit Sub
        End If
        ThisApplication.Selection.MoveDown(Unit:=WdUnits.wdLine, Count:=1)
        ThisApplication.Selection.HomeKey(Unit:=WdUnits.wdLine)
        rngStart = ThisApplication.Selection.Range.Start 'A jelenlegi kurzorhelyen kezdi a felsorolás kijelölését
        ThisApplication.Selection.Collapse(WdCollapseDirection.wdCollapseEnd)
        'Keresse meg a következő 1. szintű tételt
        If KijelölTalálatStílussal({""}, ActiveDoc.Styles(WdBuiltinStyle.wdStyleHeading1)) Then
            ThisApplication.Selection.MoveUp(Unit:=WdUnits.wdLine, Count:=1)
            'Ha nincsen, akkor az egészet jelölje ki és rendezze ...
            ThisApplication.Selection.EndKey(Unit:=WdUnits.wdLine)
            rngEnd = ThisApplication.Selection.Range.Start
            ThisApplication.Selection.Collapse(WdCollapseDirection.wdCollapseEnd)
        Else
            rngEnd = ActiveDoc.Content.End
        End If
        If rngEnd = 0 Then rngEnd = ActiveDoc.Content.End 'Ha nem volna a definíció után új címsor...

        'Rendezendő range definiálása
        rngChange = ActiveDoc.Range(rngStart, rngEnd)

        'Keresse meg a felsorolás kezdő bekezdését (az első "-"-vel jelölt bekezdést)
        For Each rngPara In rngChange.Paragraphs
            If Not rngPara.Range.ListFormat.ListTemplate Is Nothing AndAlso rngPara.Range.ListFormat.ListTemplate.ListLevels(1).NumberFormat = Chr(150) Then
                rngSortStart = rngPara.Range.Start
                Exit For
            End If
        Next
        If rngSortStart = 0 Then
            If NoWarning = False Then MsgBox("Nem talált gondolatjellel kezdődő bekezdést a definiált tartományban: " & rngStart & "-" & rngEnd)
            Exit Sub
        End If

        rngChange = ActiveDoc.Range(rngSortStart, rngEnd)
        rngChange.Select()
        'Itt kezdődik a tényleges szortírozás
        SortSelection()
        ThisApplication.Selection.Collapse(WdCollapseDirection.wdCollapseStart)
    End Sub
    Private Shared Sub SortSelection()
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application

        ThisApplication.Selection.Sort(ExcludeHeader:=False, FieldNumber:="Paragraphs", SortFieldType:=WdSortFieldType.wdSortFieldAlphanumeric,
                       SortOrder:=WdSortOrder.wdSortOrderAscending,
        FieldNumber2:="", SortFieldType2:=WdSortFieldType.wdSortFieldAlphanumeric, SortOrder2:=
        WdSortOrder.wdSortOrderAscending, FieldNumber3:="", SortFieldType3:=
        WdSortFieldType.wdSortFieldAlphanumeric, SortOrder3:=WdSortOrder.wdSortOrderAscending, Separator:=
        WdSortSeparator.wdSortSeparateByTabs, SortColumn:=False, CaseSensitive:=False, LanguageID:=WdLanguageID.wdHungarian, SubFieldNumber:="Paragraphs", SubFieldNumber2:=
        "Paragraphs", SubFieldNumber3:="Paragraphs")
    End Sub
    Private Shared Function KijelölTalálatStílussal(Keresendo As String(), StílusNeve As Style) As Boolean
        Dim Result As Boolean = False
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        ThisApplication.Selection.Find.ClearFormatting()
        ThisApplication.Selection.Find.Style = ActiveDoc.Styles(StílusNeve)
        For Each SzovegTetel In Keresendo
            With ThisApplication.Selection.Find
                .Text = SzovegTetel
                'Ha angol nyelvű, más szöveggel kell helyettesíteni, pl. "Definitions"!
                'Ha az első fejezetben szerepel a "fogal" és nem fogalommeghatározás, akkor hibás eredményt adhat, mert megkeverheti a sorokat
                .Forward = True
                .Wrap = WdFindWrap.wdFindStop
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            If ThisApplication.Selection.Find.Execute = True Then Result = True
        Next
        Return Result
    End Function

    Public Shared Sub InsertPictureBelowFooter(PicturePath As String, ThisDoc As Document)
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        If ThisDoc.ActiveWindow.View.SplitSpecial <> WdSpecialPane.wdPaneNone Then
            ThisDoc.ActiveWindow.Panes(2).Close()
        End If
        If ThisDoc.ActiveWindow.ActivePane.View.Type = WdViewType.wdNormalView Or ThisDoc.ActiveWindow.ActivePane.View.Type = WdViewType.wdOutlineView Then
            ThisDoc.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
        End If
        ThisDoc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
        Globals.ThisAddIn.Application.Selection.Collapse()
        Dim MyShape = ThisDoc.InlineShapes.AddPicture(FileName:=PicturePath, LinkToFile:=False, SaveWithDocument:=True, Range:=Globals.ThisAddIn.Application.Selection.Range)
        Dim MyShape2 = MyShape.ConvertToShape()
        MyShape2.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionBottomMarginArea
        MyShape2.TopRelative = Globals.ThisAddIn.Application.CentimetersToPoints(0.1)
        MyShape2.LeftRelative = Globals.ThisAddIn.Application.CentimetersToPoints(0.1)

    End Sub

    Public Shared Sub CVarToFill(NameToFill, ValueToFill)
        Dim Prop As Variable
        For Each Prop In Globals.ThisAddIn.Application.ActiveDocument.Variables
            If Prop.Name = NameToFill Then Prop.Delete()
        Next
        Globals.ThisAddIn.Application.ActiveDocument.Variables.Add(Name:=NameToFill, Value:=ValueToFill)
    End Sub
    Public Shared Function CVarValue(NameToFill) As String

        Dim Prop As Variable
        Dim Result As String = String.Empty
        For Each Prop In Globals.ThisAddIn.Application.ActiveDocument.Variables
            If Prop.Name = NameToFill Then Result = Prop.Value
        Next
        Return Result
    End Function

    Public Shared Sub RemoveDoubleParagraphs()
        Globals.ThisAddIn.logger.Info("RemoveDoubleParagraphs kezdete")
        CheckTrack()
        FindReplaceInternal("^p^p", "^p", WildCard:=False)
        RestoreTrack()
        Globals.ThisAddIn.logger.Info("RemoveDoubleParagraphs vége")
    End Sub

    Public Shared Sub OldRemoveDoubleParagraphs()
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Dim i As Integer
        Dim myBool As Boolean
        Dim oIdeVissza As Range
        'Dim ActiveDoc = Globals.ThisAddIn.Application.ActiveDocument
        oIdeVissza = ThisApplication.Selection.Range
        ThisApplication.Selection.Collapse()
        ActiveDoc.StoryRanges(WdStoryType.wdMainTextStory).Select()
        myBool = True
        With ThisApplication.Selection.Find
            .ClearFormatting()
            .MatchCase = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .Replacement.ClearFormatting()
            .Text = "^p^p"
            .Replacement.Text = "^p"
            Do Until myBool = False
                myBool = .Execute(Replace:=WdReplace.wdReplaceAll, Forward:=True, Format:=False)
                i = i + 1
                If i >= 100 Then Exit Sub 'Végtelen ciklus miatt, ha a dokumentumban csak paragraphok vannak.
            Loop
        End With
        ThisApplication.Selection.Collapse(WdCollapseDirection.wdCollapseStart)
        oIdeVissza.Select()
    End Sub
    Public Shared Function GetFileNameFromContractDetails()
        Dim Result As String
        Result = DateTime.Now.ToString("yyMMddHHmm") & "_"
        Dim ClientNameAsFileName = Left(HPWordHelper.ValidateFileName(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties("Company").Value), 100)
        Dim PartnerNameAsFileName = Left(HPWordHelper.ValidateFileName(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties("Category").Value), 100)
        Result += ClientNameAsFileName & PartnerNameAsFileName & Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties("Subject").Value
        Return Result
    End Function

    Public Shared Sub SetApprovalPath(Language As String)
        Dim SettingsToUpdate As Object = Nothing
        Dim LanguageInTitle As String = "a magyar"
        If Language = "English" Then
            SettingsToUpdate = My.Settings.ApprovalPicEnglish
            LanguageInTitle = "az angol"
        Else
            SettingsToUpdate = My.Settings.ApprovalPicHungarian
        End If
        Dim dialogresult As New DialogResult
        Dim ValasztottFajl = String.Empty
        Dim openFileDialog1 As New OpenFileDialog With {
            .Filter = "PNG-képfájl|*.png",
            .Title = "Válassza ki, hogy honnan tölti be " & LanguageInTitle & " nyelvű jóváhagyás képét!",
            .RestoreDirectory = True,
            .DefaultExt = "png",
            .Multiselect = False
        }
        If File.Exists(SettingsToUpdate) Then
            openFileDialog1.InitialDirectory = Path.GetDirectoryName(SettingsToUpdate)
        End If
        dialogresult = openFileDialog1.ShowDialog()
        If dialogresult = DialogResult.OK Then
            If Language = "English" Then My.Settings.ApprovalPicEnglish = openFileDialog1.FileName Else My.Settings.ApprovalPicHungarian = openFileDialog1.FileName
            My.Settings.Save()
        End If
    End Sub

    Public Shared Sub SetCursorToWaiting()
        Globals.ThisAddIn.Application.System.Cursor = WdCursorType.wdCursorWait
    End Sub
    Public Shared Sub SetCursorToDefault()
        Globals.ThisAddIn.Application.System.Cursor = WdCursorType.wdCursorNormal
    End Sub

    Public Shared Sub ChangeTerms(TermtoChangeFrom As String, TermtoChangeTo As String, Optional DefinialtFogalom As Boolean = False)
        CheckTrack()
        If Globals.ThisAddIn.Application.Documents.Count < 1 Then Exit Sub
        If String.IsNullOrWhiteSpace(TermtoChangeTo) Then Exit Sub
        If TermtoChangeFrom = TermtoChangeTo Then Exit Sub

        SetCursorToWaiting()
        If TermtoChangeFrom.Contains(" ") Or TermtoChangeTo.Contains(" ") Then
            Globals.ThisAddIn.logger.Info("If TermtoChangeFrom.Contains( ) Or TermtoChangeTo.Contains( ) Then " & TermtoChangeFrom & " _ " & TermtoChangeTo)
            ReplaceAllButLastTerms(TermtoChangeFrom, TermtoChangeTo)
            Globals.ThisAddIn.logger.Info("ReplaceAllButLastTerms sikeres " & TermtoChangeFrom & " _ " & TermtoChangeTo)
        End If
        TermtoChangeFrom = LastWord(TermtoChangeFrom)
        TermtoChangeTo = LastWord(TermtoChangeTo)

        Globals.ThisAddIn.logger.Info("ChangeTerminDoc előtt " & TermtoChangeFrom & " _ " & TermtoChangeTo)
        ChangeTerminDoc(TermtoChangeFrom, TermtoChangeTo)
        If DefinialtFogalom = True Then ChangeDefinedTerm(TermtoChangeFrom, TermtoChangeTo)
        RestoreTrack()
        SetCursorToDefault()
    End Sub
    Private Shared Sub ChangeDefinedTerm(TermToChangeFrom As String, TermtoChangeTo As String)
        TermtoChangeTo = LowerCaseFirstLetter(TermtoChangeTo)
        TermToChangeFrom = LowerCaseFirstLetter(TermToChangeFrom)
        FindReplaceAll("mint " & TermToChangeFrom, "mint " & TermtoChangeTo, IgnoreCase:=False, WildCard:=False)
        FindReplaceAll(TermToChangeFrom & " (a továbbiakban ", TermtoChangeTo & " (a továbbiakban", IgnoreCase:=False, WildCard:=False)
        FindReplaceAll(TermToChangeFrom & " (továbbiakban ", TermtoChangeTo & " (továbbiakban", False, False)
        FindReplaceAll(TermToChangeFrom & ", a továbbiakban ", TermtoChangeTo & ", a továbbiakban", False, False)
        FindReplaceAll(TermToChangeFrom & ", továbbiakban ", TermtoChangeTo & ", továbbiakban", False, False)
        FindReplaceAll(TermToChangeFrom & " (" & TermtoChangeTo, TermtoChangeTo & " (" & TermtoChangeTo, False, False)
    End Sub
    Private Shared Function LastWord(TermToChange As String) As String
        If Not TermToChange.Contains(" ") Then Return TermToChange
        Dim ResultList As List(Of String) = TermToChange.Split(" ").ToList
        Return ResultList.Last
    End Function
    Private Shared Sub ReplaceAllButLastTerms(TermToChangeFrom As String, TermToChangeTo As String)
        Dim IsTermToChangeFromMultiWord As Boolean = TermToChangeFrom.Contains(" ")
        Dim IsTermToChangeToMultiWord As Boolean = TermToChangeTo.Contains(" ")
        If IsTermToChangeFromMultiWord AndAlso IsTermToChangeToMultiWord Then
            Dim AllButLastFrom As List(Of String) = TermToChangeFrom.Split(" ", options:=StringSplitOptions.RemoveEmptyEntries).ToList
            AllButLastFrom.RemoveAt(AllButLastFrom.Count - 1)
            Dim AllButLastTo As List(Of String) = TermToChangeTo.Split(" ", options:=StringSplitOptions.RemoveEmptyEntries).ToList
            AllButLastTo.RemoveAt(AllButLastTo.Count - 1)
            Globals.ThisAddIn.logger.Info("AllButLastFrom/To: " & String.Join(" ", AllButLastFrom.ToArray) & String.Join(" ", AllButLastTo.ToArray))
            FindReplaceAll(String.Join(" ", AllButLastFrom.ToArray), String.Join(" ", AllButLastTo.ToArray), IgnoreCase:=False)
        End If
        If IsTermToChangeFromMultiWord AndAlso Not IsTermToChangeToMultiWord Then
            Dim AllButLastFrom As List(Of String) = TermToChangeFrom.Split(" ").ToList
            AllButLastFrom.RemoveAt(AllButLastFrom.Count - 1)
            FindReplaceAll(String.Join(" ", AllButLastFrom.ToArray), String.Empty, IgnoreCase:=False)
        End If
        If Not IsTermToChangeFromMultiWord AndAlso IsTermToChangeToMultiWord Then
            Dim AllButLastTo As List(Of String) = TermToChangeTo.Split(" ").ToList
            AllButLastTo.RemoveAt(AllButLastTo.Count - 1)
            HPWordHelper.FindReplaceAll(TermToChangeFrom, String.Join(" ", AllButLastTo.ToArray) & " " & TermToChangeFrom, IgnoreCase:=False)
        End If
    End Sub

    Friend Shared Sub ChangeTerminDoc(TermToChangeFrom As String, TermtoChangeTo As String)
        CheckTrack()
        Dim rng = Globals.ThisAddIn.Application.ActiveDocument.Content
        If TermToChangeFrom = TermtoChangeTo Then
            Globals.ThisAddIn.logger.Info("Azonos a két cserélendő szöveg, ezért kilép a ChangeTerminDoc sub.")
            Exit Sub
        End If
        rng.Find.MatchCase = True
        rng.Find.MatchPrefix = True 'ezzel a módosítással próbáltam a "vállalkozó" "megbízottra" csere esetén fellépő alvállalkozó törléseket elkerülni
        rng.Find.Text = TermToChangeFrom
        rng.Find.Execute()
        Dim CouldNotGenerate As New Dictionary(Of String, String)
        While rng.Find.Found
            rng.Expand(Word.WdUnits.wdWord)
            Dim key As String = rng.Text
            key = key.TrimEnd
            'If key <> rng.Text Then rng.MoveEnd(Word.WdUnits.wdCharacter, -1)

            Dim TermtoChangeFromFull = key
            Dim ChangeTo = NLP.GenerateWord(TermtoChangeTo, TermtoChangeFromFull)
            If ChangeTo <> String.Empty Then
                Dim rng2 = Globals.ThisAddIn.Application.ActiveDocument.Content
                rng2.Find.Text = TermtoChangeFromFull
                rng2.Find.MatchCase = True
                rng2.Find.MatchWholeWord = True
                rng2.Find.Replacement.Text = ChangeTo
                rng2.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            Else
                ChangeTo = Guid.NewGuid.ToString().Replace("-", "")
                CouldNotGenerate.Add(ChangeTo, key)
                Dim rng2 = Globals.ThisAddIn.Application.ActiveDocument.Content
                rng2.Find.Text = TermtoChangeFromFull
                rng2.Find.MatchWholeWord = True
                rng2.Find.MatchCase = True
                rng2.Find.Replacement.Text = ChangeTo
                rng2.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            End If
            rng = Globals.ThisAddIn.Application.ActiveDocument.Content
            rng.Find.MatchCase = True
            rng.Find.Text = TermToChangeFrom
            rng.Find.Execute()
        End While

        Dim CouldNotGenerateWords As String = String.Empty
        For Each TempName As String In CouldNotGenerate.Keys
            Dim Value = CouldNotGenerate.Item(TempName)
            CouldNotGenerateWords += Value + "; "
            Dim rng2 = Globals.ThisAddIn.Application.ActiveDocument.Content
            rng2.Find.Text = TempName
            rng2.Find.MatchWholeWord = True
            rng2.Find.Replacement.Text = Value
            rng2.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            Dim rng3 = Globals.ThisAddIn.Application.ActiveDocument.Content
            rng2.Find.Text = TermToChangeFrom
            rng2.Find.MatchWholeWord = False
            rng2.Find.Replacement.Text = TermtoChangeTo
            rng2.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        Next
        If CouldNotGenerate.Count > 0 Then
            MsgBox("A következő szavak átalakításának helyességét ellenőrizze:" & vbCrLf & CouldNotGenerateWords)
        End If
        Marshal.ReleaseComObject(rng)
        RestoreTrack()
    End Sub
    Public Shared Sub RagozzMindent()
        Globals.ThisAddIn.logger.Info("RagozzMindent() kezdete")
        CheckTrack()

        For i = 1 To 4
            DeclinationNounInWord("t", "Megrendelőt") 'Accusativus
            DeclinationNounInWord("on", "Megrendelőn") 'Superessivus
            DeclinationNounInWord("ben", "Megrendelőben") 'Inessivus
            DeclinationNounInWord("nek", "Megrendelőnek") 'Dativus
            DeclinationNounInWord("re", "Megrendelőre")  'Sublativus
            DeclinationNounInWord("vel", "Megrendelővel") 'Instrumentalis
            DeclinationNounInWord("jével", "Megrendelőjével") 'Instrumentalis + possessivus
            DeclinationNounInWord("jét", "Megrendelőjét") 'Accusativus + possessivus
            DeclinationNounInWord("ében", "Megrendelőjében") 'Inessivus + possessivus
        Next i
        RestoreTrack()
        Globals.ThisAddIn.logger.Info("RagozzMindent() vége")
    End Sub
    Friend Shared Sub DeclinationNounInWord(NounForm As String, SampleForm As String)
        CheckTrack()
        Dim CodeName As String = "`" & NounForm & "`"
        Dim rng = Globals.ThisAddIn.Application.ActiveDocument.Content
        rng.Find.MatchCase = True
        rng.Find.Text = CodeName
        rng.Find.MatchPrefix = True 'csak a szóeleji egyezőségre keres rá
        rng.Find.Execute()
        While rng.Find.Found
            Dim OriginalRangeStart = rng.Start
            'rng.Expand(Word.WdUnits.wdWord) 'nem jó: bár az utána lévő szóközt is kijelöli, de nem jelöli ki, ha a `CodeName`-mel jelölt szó első a sorban vagy első a dokumentumban
            rng.MoveStartUntil(" ", Word.WdConstants.wdBackward)
            If OriginalRangeStart = rng.Start Then 'Ha a visszaugrás is ugyanaz maradt a kijelölés kezdete, azaza `CodeName` kezdőpontja, akkor vagy a sorban első vagy a dokumentumban első szóról van szó
                rng.MoveStartUntil(vbCrLf, Word.WdConstants.wdBackward) 'Sorban első szóról van szó
                If rng.Start = OriginalRangeStart Then rng.Start = 0 'Dokumentumban első szóról van szó
            End If
            rng.Find.MatchCase = True
            Dim ToChange As String = rng.Text
            rng.Find.Text = ToChange
            Dim WordForGeneration As String = ToChange.Replace(CodeName, "")
            Dim Mire As String = NLP.GenerateWord(WordForGeneration, SampleForm)
            rng.Find.Replacement.Text = Mire
            rng.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            rng = Globals.ThisAddIn.Application.ActiveDocument.Content
            rng.Find.Text = CodeName
        End While
        Marshal.ReleaseComObject(rng)
        RestoreTrack()
    End Sub

    Public Sub RegexAAzCsere()
        'Dim MatchPattern As String = "\baz (?=[bcdfghjklmnpqrstvwxyz]\w*)"
        Dim MatchPattern As String = "(?!(az\snem))(\baz (?=[bcdfghjklmnpqrstvwxyz]\w*))"
        Dim ReplacementString = "a "
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument.Range
        Dim Cserelendo = ActiveDoc.Text
        Dim Cserelt As String = Regex.Replace(Cserelendo, MatchPattern, ReplacementString, RegexOptions.IgnoreCase)
        ActiveDoc.Text = Cserelt
    End Sub

    Private Shared Sub CheckTrack()
        VoltTrack = Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions
        If VoltTrack = True Then Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False
    End Sub
    Private Shared Sub RestoreTrack()
        If VoltTrack = True AndAlso Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False Then
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = True
        End If
    End Sub
    Public Shared Sub HibasAAzCsere()
        Globals.ThisAddIn.logger.Info("AAzcsere kezdete")
        HPWordHelper.SetCursorToWaiting()
        CheckTrack()
        'számozás
        Globals.ThisAddIn.logger.Info("HivatkozasSzamAAzCsere kezdete")
        HivatkozasSzamAAzCsere()
        ''Az alábbi metódus a kereszthivatkozások esetén nem jól helyezte el a névelőt, pl. "a [1. melléklet]" esetét nem érzékelte jól, ezért van a számmal kezdődőkre külön metódus
        ''A mutató névmás és az "az" névelő nem számok esetén azonos alakú, így félrevezető lehet az alábbi, egyébként általában jól működő algoritmus
        'Globals.ThisAddIn.logger.Info("<a (<[FLMNRSXY][BCDFGHJKLMNPQRSTVWXYZ]) kezd")
        'HPWordHelper.FindReplaceAll("<a (<[FLMNRSXY][BCDFGHJKLMNPQRSTVWXYZ])", "az \1", True)
        'Globals.ThisAddIn.logger.Info("<az (<[bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZ] kezd")
        'HPWordHelper.FindReplaceAll("<az (<[bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZ])", "a \1", True)
        'Globals.ThisAddIn.logger.Info("<a (<[aáeéiíoóöőuúüűAÁEÉIÍOÓÖŐUÚÜŰ] kezd")
        'HPWordHelper.FindReplaceAll("<a (<[aáeéiíoóöőuúüűAÁEÉIÍOÓÖŐUÚÜŰ])", "az \1", True)
        HPWordHelper.SetCursorToDefault()
        RestoreTrack()
        Globals.ThisAddIn.logger.Info("AAzcsere vége")
    End Sub

    Public Shared Sub HivatkozasSzamAAzCsere()
        Dim i As Integer
        Dim oIdeVissza As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        For i = 6 To 9
            HPWordHelper.HivatkozasRaKeres(" az " & CStr(i))
            HPWordHelper.HivatkozasRaKeres("^paz " & CStr(i))
        Next i
        For i = 2 To 4
            HPWordHelper.HivatkozasRaKeres(" az " & CStr(i))
            HPWordHelper.HivatkozasRaKeres("^paz " & CStr(i))
        Next i
        HivatkozasRaKeresSzamjegy(0, "1", Az:=False, Cardinal:=False)
        HivatkozasRaKeresSzamjegy(0, "1", Az:=False, Cardinal:=True)
        HivatkozasRaKeresSzamjegy(3, "1", Az:=False, Cardinal:=False)
        HivatkozasRaKeresSzamjegy(6, "1", Az:=False, Cardinal:=False)
        HivatkozasRaKeresSzamjegy(0, "5", Az:=False, Cardinal:=False)
        HivatkozasRaKeresSzamjegy(0, "5", Az:=False, Cardinal:=True)
        Dim SzamjegyLista As Array = {1, 2, 4, 5, 7, 8}
        For Each Szamjegy In SzamjegyLista
            HivatkozasRaKeresSzamjegy(Szamjegy, "1", Az:=True, Cardinal:=True)
        Next
    End Sub

    Private Shared Sub HivatkozasRaKeresSzamjegy(HanyJegyElsoUtan As Integer, MilyenJegy As String, Az As Boolean, Cardinal As Boolean)
        If MilyenJegy = String.Empty Then Exit Sub
        Dim KeresoSzoveg As String = String.Empty
        Dim Kezdoresz As String = "a"
        Dim Lezaroresz As String = String.Empty
        If Az = True Then Kezdoresz = "az"
        If Cardinal = True Then Lezaroresz = "." Else Lezaroresz = " "
        Dim Counter As Integer = 0
        Do While Counter < HanyJegyElsoUtan
            KeresoSzoveg = KeresoSzoveg + "^#"
            Counter = Counter + 1
        Loop
        HPWordHelper.HivatkozasRaKeres(" " + Kezdoresz + " " + MilyenJegy + KeresoSzoveg + Lezaroresz)
        HPWordHelper.HivatkozasRaKeres("^p" + Kezdoresz + " " + MilyenJegy + KeresoSzoveg + Lezaroresz)
        If Cardinal = False Then
            Dim Lezarok = {",", "^p"}
            For Each LezaroChar As String In Lezarok
                Lezaroresz = LezaroChar
                HPWordHelper.HivatkozasRaKeres(" " + Kezdoresz + " " + MilyenJegy + KeresoSzoveg + Lezaroresz)
                HPWordHelper.HivatkozasRaKeres("^p" + Kezdoresz + " " + MilyenJegy + KeresoSzoveg + Lezaroresz)
            Next
        End If
    End Sub

    Friend Shared Function AddRichTextControl(Where As Range, ControlName As String) As RichTextContentControl
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim extendedDocument = Globals.Factory.GetVstoObject(ActiveDoc)
        Dim richTextControlTemp = extendedDocument.Controls.AddRichTextContentControl(Where, ControlName)
        richTextControlTemp.PlaceholderText = "Enter the value"
        Return richTextControlTemp
    End Function

    Friend Shared Function AddContentControl(Where As Range, Type As WdContentControlType, Optional Title As String = "", Optional PlaceHolderText As String = "Enter the value", Optional Style As String = "") As ContentControl
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim Result As ContentControl
        Dim myRange As Range
        Dim currentSelection As Range
        'If Not IsNothing(Globals.ThisAddIn.Application.Selection) Then
        'currentSelection = Globals.ThisAddIn.Application.Selection.Range
        'End If
        myRange = Where
        myRange.Text = "" 'Azért nem myRange.Delete, mert akkor letörli az egymás után következő két space-t a Word
        myRange.Select()
        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseStart)
        Try
            Result = ActiveDoc.ContentControls.Add(Type)
        Catch ex As Exception
            Globals.ThisAddIn.logger.Info("Nem tudta beállítani ezt: " & Where.Start & " _End: " & Where.End)
            Return Nothing
        End Try
        Result.Title = Title
        If Not String.IsNullOrEmpty(Style) Then
            Try
                Result.DefaultTextStyle = Style
            Catch ex As Exception
                Globals.ThisAddIn.logger.Info("Nem tudta beállítani ezt a karakterstílust: " & Style)
            End Try
        End If
        'If Not IsNothing(currentSelection) Then currentSelection.Select() : Selection
        Return Result
    End Function

    Friend Shared Function AddCustomXmlPartToDocument() As Office.CustomXMLPart
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim ExistingXMLPart As Office.CustomXMLParts = ActiveDoc.CustomXMLParts.SelectByNamespace(CustomContractXMLNameSpace)
        'For Each cxml As Microsoft.Office.Core.CustomXMLPart In ActiveDoc.CustomXMLParts
        '    Globals.ThisAddIn.logger.Info("CustomXMLPart1:" & cxml.NamespaceURI)
        '    Globals.ThisAddIn.logger.Info("CustomXMLPart2:" & cxml.XML)
        'Next
        If ExistingXMLPart.Count > 0 Then
            For Each customXML As Office.CustomXMLPart In ExistingXMLPart
                If Not customXML.BuiltIn Then customXML.Delete()
            Next
            'Dim e = ActiveDoc.CustomXMLParts.GetEnumerator()
            'Dim p As Microsoft.Office.Core.CustomXMLPart
            'While e.MoveNext()
            '    p = DirectCast(e.Current, Microsoft.Office.Core.CustomXMLPart)
            '    'p.BuiltIn will be true for internal buildin excel parts 
            '    If p IsNot Nothing AndAlso Not p.BuiltIn AndAlso p.NamespaceURI = NS.NamespaceName Then
            '        p.Delete()
            '    End If
        End If
        Dim xmlString As String = My.Resources.blank_customXMLContracts
        Dim ContractCustomXML As Office.CustomXMLPart = ActiveDoc.CustomXMLParts.Add(xmlString)
        Globals.ThisAddIn.logger.Info("CustomXMLPart rögzítve")
        'ContractCustomXML.NamespaceManager.AddNamespace("cxl", CustomContractXMLNameSpace)
        Return ContractCustomXML
    End Function
    Friend Shared Sub AddAllCapStyle()
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Try
            ActiveDoc.Styles(AllCapStyleName).Font.AllCaps = True
        Catch ex As Exception
            Dim AllCapStyle = ActiveDoc.Styles.Add(AllCapStyleName, Word.WdStyleType.wdStyleTypeCharacter)
            AllCapStyle.Font.AllCaps = True
        End Try
    End Sub
End Class
