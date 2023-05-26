Imports HPW = HotDocs_HP_WordAddIn.HPWordHelper
Imports HD = HotDocs
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Diagnostics

Public Class _client_SpecificChanges
    Friend Shared Sub _client_SpecificTransformations(AnswerCollectionUsed As HD.AnswerCollection)
        Globals.ThisAddIn.logger.Info("_client_SpecificTransformations kezdete")
        Dim frmTerms As New frmChange3Terms
        frmTerms.ShowDialog()
        HPW.DefinitionsSort(NoWarning:=True)
        HPW.UpdateAllFields(NoWarning:=True)
        HPW.MellékletHivatkozásTörlése()
        AdjustNumbering(AnswerCollectionUsed)
        Globals.ThisAddIn.logger.Info("_client_SpecificTransformations vége")
    End Sub
    Friend Shared Sub AdjustNumbering(AnswerCollectionUsed As HD.AnswerCollection)
        Dim ContractType As String = String.Empty
        Dim IsAmendment As Boolean = False
        Globals.ThisAddIn.logger.Info("AdjustNumbering kezdete")
        Try
            ContractType = AnswerCollectionUsed.Item("ContractMainType", HD.HDVARTYPE.HD_MULTCHOICETYPE).Value
            IsAmendment = AnswerCollectionUsed.Item("IsAmendment", HD.HDVARTYPE.HD_TRUEFALSETYPE).Value
        Catch ex As Exception
            Globals.ThisAddIn.logger.Info("ContractType or IsAmendment hiba:" & ex.Message)
            If Not IsNothing(ContractType) Then Globals.ThisAddIn.logger.Info("ContractType = " & ContractType.ToString)
        End Try

        If ContractType = "Lending" Then LinkListLevelsto_client_(CByte(1))
        If IsAmendment Then
            LinkListLevelsto_client_(CByte(0))
            Globals.ThisAddIn.logger.Info("LinkListLevelsto_client_(CByte(0))")
        End If
        Globals.ThisAddIn.logger.Info("AdjustNumbering vége")
    End Sub
    Friend Shared Sub SetBodyTextLevel()
        Dim Meddig = InputBox("Meddig legyen címsor?")
        Dim MeddigSzam As Byte = 0
        Try
            MeddigSzam = CByte(Meddig)
        Catch ex As Exception
            MsgBox("A megadott érték legyen 0-9 közötti szám!", MsgBoxStyle.Critical)
            Exit Sub
        End Try
        If MeddigSzam > 9 Then MeddigSzam = 9
        LinkListLevelsto_client_(CByte(Meddig))
    End Sub
    Friend Shared Sub Apply_client_BodyTextToSelection()
        Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleBodyText3
    End Sub
    Public Shared Sub LinkListLevelsto_client_(MeddigCímsor As Byte)
        Globals.ThisAddIn.logger.Info("LinkListLevelsto_client_ kezdete")
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        HPW.SetCursorToWaiting()
        Dim i As Byte
        Dim BodyText3Name = ActiveDoc.Styles(WdBuiltinStyle.wdStyleBodyText3).NameLocal
        If MeddigCímsor = 0 Then
            LinkListLevelTo_client_Style(BodyText3Name, 1) 'Nem lehet wdBuiltinStyle.wdStyleBodyText3, mert az számot ad, és a LinkedStyle-nak String kell inputként
            ReplaceAllBodyText3ToSame()
            Globals.ThisAddIn.logger.Info("LinkListLevelsto_client_ (BodyText3Name,1) lefutott")
            Exit Sub
        End If

        For i = 1 To MeddigCímsor
            LinkListLevelTo_client_Style("Heading " & CStr(i), i)
            'Mivel stringben kéri a LinkedStyle, ezért az angol nevet írtuk, de általánosabbá teheto így: "ActiveDocument.Styles(wdStyleBodyText3).NameLocal"
        Next i
        i = 1
        If MeddigCímsor < 9 Then LinkListLevelTo_client_Style(BodyText3Name, MeddigCímsor + 1)
        If MeddigCímsor < 8 Then LinkListLevelTo_client_Style("List", MeddigCímsor + 2)
        If MeddigCímsor < 7 Then
            For i = MeddigCímsor + 3 To 5
                '9-ig kellene, de csak 5 lista van maximum beépített stílusban
                LinkListLevelTo_client_Style("List " & CStr(i - MeddigCímsor - 1), i)
            Next i
        End If
        ReplaceAllBodyText3ToSame()
        HPW.SetCursorToDefault()
        Globals.ThisAddIn.logger.Info("LinkListLevelsto_client_ vége")
    End Sub
    Public Shared Sub LinkListLevelTo_client_Style(BekezdésStílusNeve As String, Szint As Integer)
        Globals.ThisAddIn.Application.ActiveDocument.Styles("_client__lista").ListTemplate.ListLevels(Szint).LinkedStyle = BekezdésStílusNeve
    End Sub
    Public Shared Sub FileSaveAs_client_Name()
        MsgBox(HPW.GetFileNameFromContractDetails() & " NOT IMPLEMENTED YET") 'NOT IMPLEMENTED YET!  dialog box needed
    End Sub

    Public Shared Sub Copy_client_Template()
        Dim PathToCopy As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "Custom Office Templates"
        Dim FileDestination As String = Path.Combine(PathToCopy, My.Settings._client_TemplateName)
        Dim ContentPath As String = Path.Combine(Globals.ThisAddIn.DictionaryPath, My.Settings._client_TemplateName)
        If Not File.Exists(ContentPath) Then
            MsgBox("Nem található: " & ContentPath)
            Exit Sub
        End If
        If Not Directory.Exists(PathToCopy) Then
            Try
                Directory.CreateDirectory(PathToCopy)
            Catch ex As Exception
                MsgBox(PathToCopy & " könyvtár létrehozása sikertelen: " & ex.Message)
            End Try
        End If
        File.Copy(ContentPath, FileDestination, True)
        With Globals.ThisAddIn.Application.ActiveDocument
            .UpdateStylesOnOpen = True
            .AttachedTemplate = FileDestination
        End With
    End Sub

    Public Shared Sub InsertApproval(FilePathToInsert As String)
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Globals.ThisAddIn.logger.Info("Jóváhagyás elindult")
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Globals.ThisAddIn.logger.Info("ActiveDoc neve: " & ActiveDoc.FullName)
        ThisApplication.ActiveDocument.Content.HighlightColorIndex = WdColorIndex.wdNoHighlight
        ThisApplication.Selection.Collapse()
        If ActiveDoc.TrackRevisions = True Then ActiveDoc.TrackRevisions = False
        ActiveDoc.Revisions.AcceptAll()
        For Each oComment In ActiveDoc.Comments
            oComment.Delete
        Next
        HPW.UpdateAllFields()
        Globals.ThisAddIn.logger.Info("Jóváhagyás updateallfields sikeres")
        If HPW.MezoEllenorzes() = True Then MsgBox("A program formázási hiba miatt a jóváhagyást nem folytatta.") : Exit Sub
        HPW.UnlinkAllReferences()
        Globals.ThisAddIn.logger.Info("Jóváhagyás unlink references sikeres")
        Dim oSection As Section
        Dim currentView As WdViewType
        currentView = ActiveDoc.ActiveWindow.ActivePane.View.Type
        ActiveDoc.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
        If Not File.Exists(FilePathToInsert) Then
            MsgBox("A megadott jóváhagyási kép útvonalon nincsen fájl, állítsa át a jóváhagyási kép adatait!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        ThisApplication.Selection.Collapse()
        For Each oSection In ActiveDoc.Sections
            HPW.InsertPictureBelowFooter(FilePathToInsert, ActiveDoc)
        Next
        ActiveDoc.ActiveWindow.View.Type = currentView
        Globals.ThisAddIn.logger.Info("Jóváhagyás MentesExportig elfutottt")
        JovahagyasMentesExporttal(ActiveDoc)
    End Sub
    Public Shared Sub JovahagyasMentesExporttal(ThisDoc As Document)

        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        'ThisDoc.Save()
        Dim CurrentName = ThisDoc.FullName
        Globals.ThisAddIn.logger.Info("JovahagyasMentesExporttal kezdés, CurrentName=" & CurrentName)
        Dim ApprovalPath As String = String.Empty
#If LKT = "Y" Then
        ApprovalPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Jovahagyas")
        Globals.ThisAddIn.logger.Info("JovahagyasMentesExporttal, ApprovalPath= " & ApprovalPath)

        If Not Directory.Exists(ApprovalPath) Then
            Try
                Directory.CreateDirectory(ApprovalPath)
            Catch ex As Exception
                MsgBox("Jóváhagyási mappa létrehozási hiba: " & ex.Message & vbCrLf & "Mentés a meglévő fájl mellé")
            End Try
        End If
#Else
        ApprovalPath = Path.GetDirectoryName(CurrentName)
#End If
        CurrentName = Path.Combine(ApprovalPath, Path.GetFileName(CurrentName))
        Dim NoExtName = Path.Combine(Path.GetDirectoryName(CurrentName), Path.GetFileNameWithoutExtension(CurrentName) & "_fin")
        Dim ApprovedName = NoExtName & Path.GetExtension(CurrentName)
        HPW.CVarToFill("Jóváhagyott", "-1")
        HPW.CVarToFill("JóváhagyásDátuma", CStr(Now.ToShortDateString))
        Dim PDFFileName = NoExtName & ".pdf"
        Try
            ThisDoc.SaveAs(Path.Combine(ApprovalPath, ApprovedName))
        Catch ex As Exception
            MsgBox("ThisDoc.SaveAs - Mentési hiba")
            Globals.ThisAddIn.logger.Info("Mentési hiba - CurrentName: " & CurrentName & vbCrLf & "Path.Combine(ApprovalPath, ApprovedName)" & Path.Combine(ApprovalPath, ApprovedName))
            Exit Sub
        End Try
        Try
            ThisDoc.ExportAsFixedFormat(OutputFileName:=PDFFileName, ExportFormat:=
                WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range:=WdExportRange.wdExportAllDocument, From:=1, To:=1, Item:=WdExportItem.wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True,
                CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=False,
                BitmapMissingFonts:=False, UseISO19005_1:=False)
        Catch
            MsgBox("ThisDoc.SaveAs - PDF készítési hiba")
            Globals.ThisAddIn.logger.Info("PDF készítési hiba - PDFFileName: " & PDFFileName)
            Exit Sub
        End Try

        ES3.CreateES3(NoExtName, PDFFileName)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
            .FileName = Path.Combine(ApprovalPath, NoExtName & ".es3"),
            .UseShellExecute = True,
            .Verb = "open"
            })
        Process.Start(Path.GetDirectoryName(PDFFileName))

        'TODO: a Word és a PDF és az ES3-ból "Indulhat a csomag készítése?" zippelje őket össze és fileexplorer popupolja?
    End Sub

End Class
