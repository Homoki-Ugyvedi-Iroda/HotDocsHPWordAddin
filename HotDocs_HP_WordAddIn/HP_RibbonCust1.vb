Imports Microsoft.Office.Tools.Ribbon
#If NoHotDocs <> "Y" Then
Imports HD = HotDocs
Imports HDCust = HotDocs_HP_WordAddIn.HDCustomizations
#End If
Imports HPW = HotDocs_HP_WordAddIn.HPWordHelper
Imports HPWStyle = HotDocs_HP_WordAddIn.HPWordStyleHelper
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Windows.Forms
Imports System.Deployment.Application
Imports System.Reflection

Public Class HP_RibbonCust1
    Public Shared SelectedTitle As String = String.Empty
    Public Shared SelectedDescription As String = String.Empty
#If NoHotDocs <> "Y" Then
    Friend WithEvents MyHotDocsApp As New HD.Application
#End If

    Private Sub TurnOffScreen()
        Globals.ThisAddIn.Application.System.Cursor = WdCursorType.wdCursorWait
        Globals.ThisAddIn.Application.ScreenUpdating = False
    End Sub
    Private Sub TurnOnScreen()
        Globals.ThisAddIn.Application.System.Cursor = WdCursorType.wdCursorNormal
        Globals.ThisAddIn.Application.ScreenUpdating = True
        Globals.ThisAddIn.logger.Info("TurnOnScreen vége")
    End Sub
#If NoHotDocs <> "Y" Then
    Private Sub btnSelectHotDocsTemplate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnChooseTemplate.Click
        Dim tplPath As String = String.Empty
        Try
            MyHotDocsApp.SelectTemplate(My.Settings.HotDocsTemplate, tplPath:=tplPath, tplTitle:=SelectedTitle, tplDesc:=SelectedDescription)
            If Not String.IsNullOrEmpty(tplPath) Then MyHotDocsApp.Assemblies.AddToQueue(tplPath, True)
            Marshal.ReleaseComObject(MyHotDocsApp)
        Catch ex As Exception
            Globals.ThisAddIn.logger.Info("HotDocs nincsen telepítve")
            Exit Sub
        End Try
        If Not File.Exists(My.Settings.HotDocsTemplate) Then
            MsgBox("Nincsen kijelölve sablon, addig nem folytathatom!", MsgBoxStyle.Critical)
            Exit Sub
        End If
    End Sub
    Private Sub app_OnAssemblyCompleteEvent(tplTitle As String, tplPath As String, docPath As String, ansCall As Object, assemblyHandle As Integer) Handles MyHotDocsApp.AssemblyCompleteEvent
        Dim ActiveDoc As Document = Globals.ThisAddIn.Application.ActiveDocument
        If ActiveDoc Is Nothing Then
            Globals.ThisAddIn.logger.Error("app_OnAssemblyCompleteEvent hiányzik ActiveDoc")
            Exit Sub
        End If
        Dim MyAnswerColl = HDCust.GetAnswerCollection(ansCall)
        If MyAnswerColl Is Nothing Then Exit Sub
        Globals.ThisAddIn.LastUsedAnswerCollection = MyAnswerColl
        Dim TransformationLevel As HDCust.TransformationType = HDCust.GetTransformationLevelRequired(MyAnswerColl)
        Select Case TransformationLevel
            Case HDCust.TransformationType.None
                Exit Sub
            Case HDCust.TransformationType.Ask
                Dim StartAssembly As DialogResult = MsgBox("Megkezdjem az összeállítás utáni Word szintű átalakításokat?", MsgBoxStyle.YesNo)
                If Not StartAssembly = MsgBoxResult.Yes Then Exit Sub
            Case HDCust.TransformationType.CIBSpecific
            Case HDCust.TransformationType.Full
        End Select
        TurnOffScreen()
        HDCust.SetNonXMLDocPropertiesBasedOnHDAnswers(MyAnswerColl)
        HDCust.SetXMLDocPropertiesCreateContentControlsBasedOnHDAnswers(MyAnswerColl)
        HPW.RemoveDoubleParagraphs()
        If TransformationLevel = HDCust.TransformationType.CIBSpecific Then
            CIBSpecificChanges.CIBSpecificTransformations(MyAnswerColl)
        Else
            HPW.UpdateAllFields(NoWarning:=True)
        End If
        HPW.RemoveDoubleParagraphs()
        HPW.RagozzMindent()
        HPW.HibasAAzCsere()
        TurnOnScreen()
        MsgBox("Tervezet kész!", MsgBoxStyle.ApplicationModal)
    End Sub
#End If

    Private Sub btnReplaceTerms_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReplaceTerms.Click
        Dim frmTerms As New frmChange3Terms
        frmTerms.ShowDialog()
        Globals.ThisAddIn.logger.Info("btnReplaceTerms vége")
    End Sub

    Private Sub btnRagozzMindent_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRagozzMindent.Click
        HPW.SetCursorToWaiting()
        HPW.RagozzMindent()
        Globals.ThisAddIn.logger.Info("btnRagozzMindent_Click vége")
        HPW.SetCursorToDefault()
        'MyHDCust.HibasAAzCsere()
    End Sub

    Private Sub btnRagozzTesztelj_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRagozzTesztelés.Click
        Dim Input As String = InputBox("Mit ragozzak?")
        Dim InputSample As String = InputBox("Miként ragozzam?")
        MsgBox(NLP.GenerateWord(Input, InputSample))
    End Sub

    Private Sub btnAAzCsere_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAAzCsere.Click
        HPW.SetCursorToWaiting()
        HPW.HibasAAzCsere()
        HPW.SetCursorToDefault()
    End Sub

    Private Sub btnApprovalSettingHU_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApprovalSettingHU.Click
        HPW.SetApprovalPath("Hungarian")
    End Sub

    Private Sub btnApprovalSettingEN_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApprovalSettingEN.Click
        HPW.SetApprovalPath("English")
    End Sub

    Private Sub btnApprovalHu_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApprovalHu.Click
        CIBSpecificChanges.InsertApproval(My.Settings.ApprovalPicHungarian)
    End Sub

    Private Sub btnApprovalEN_Click(sender As Object, e As RibbonControlEventArgs)
        CIBSpecificChanges.InsertApproval(My.Settings.ApprovalPicEnglish)
    End Sub

    Private Sub btnCVarView_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCVarView.Click
        Dim VarName = InputBox("Neve a változónak, amit ellenőrizni kíván?")
        If VarName <> String.Empty Then MsgBox(HPWordHelper.CVarValue(VarName))
    End Sub

    Private Sub btnCVarToFill_Click(sender As Object, e As RibbonControlEventArgs)
        Dim VarName = InputBox("Neve a változónak, amit kitölteni kíván?")
        Dim VarValue = InputBox("Változó értéke?")
        If VarName <> String.Empty AndAlso VarValue <> String.Empty Then HPWordHelper.CVarToFill(VarName, VarValue)
    End Sub

    Private Sub btnDefinitionsSort_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDefinitionsSort.Click
        HPW.DefinitionsSort()
    End Sub

    Private Sub btnUnlinkAllReferences_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUnlinkAllReferences.Click
        HPW.UnlinkAllReferences()
    End Sub

    Private Sub btnOwnWordSettingsHP_Click(sender As Object, e As RibbonControlEventArgs) Handles btnOwnWordSettingsHP.Click
        HPW.OwnWordSettingsHP()
    End Sub

    Private Sub btnUpdateAllFields_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateAllFields.Click
        HPW.UpdateAllFields()
    End Sub

    Private Sub btnMezoEllenorzes_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMezoEllenorzes.Click
        HPW.MezoEllenorzes()
    End Sub

    Private Sub btnCommentCounter_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCommentCounter.Click
        HPWStyle.CommentCounter()
    End Sub

    Private Sub btnFileAndCommentsCompare_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFileAndCommentsCompare.Click
        HPWStyle.FileAndCommentsCompare()
    End Sub

    Private Sub btnCommentsWriteOut_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCommentsWriteOut.Click
        HPWStyle.WriteAllCommToFile()
    End Sub

    Private Sub btnWriteNewComments_Click(sender As Object, e As RibbonControlEventArgs) Handles btnWriteNewComments.Click
        HPWStyle.WriteNewComments()
    End Sub

    Private Sub btnCreateDefaultListStyles_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCreateDefaultListStyles.Click
        HPWStyle.CreateDefaultListStyles()
    End Sub

    Private Sub btnCreateAndApplyMLevelList_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCreateAndApplyMLevelList.Click
        HPWStyle.CreateAndApplyMLevelList()
    End Sub

    Private Sub btnLvl2_NameNumberedBodyTextLvl_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLvl2_NameNumberedBodyTextLvl.Click
        HPWStyle.Lvl2_NameNumberedBodyTextLvl()
    End Sub

    Private Sub btnLvl1_NameNumberedBodyTextLvl_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLvl1_NameNumberedBodyTextLvl.Click
        HPWStyle.Lvl1_NameNumberedBodyTextLvl()
    End Sub

    Private Sub btnLvl3_NameNumberedBodyTextLvl_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLvl3_NameNumberedBodyTextLvl.Click
        HPWStyle.Lvl3_NameNumberedBodyTextLvl()
    End Sub

    Private Sub btnSetLvlNameNumberedBodyTextLvl_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetLvlNameNumberedBodyTextLvl.Click
        HPWStyle.LvlInputNameNumberedBodyTextLvl()
    End Sub

    Private Sub btnCurrentLvlBodyText_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCurrentLvlBodyText.Click
        HPWStyle.LvlCurrentNameNumberedBodyTextLvl()
    End Sub

    Private Sub btnCIBStyleBodyTextLvl_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCIBStyleBodyTextLvl.Click
        CIBSpecificChanges.SetBodyTextLevel()
    End Sub

    Private Sub btnApplyBodyText_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApplyBodyText.Click
        HPWStyle.ApplyBodyText()
    End Sub

    Private Sub btnBodyText2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnBodyText2.Click
        HPWStyle.ApplyBodyTextNumbered()
    End Sub

    Private Sub btnBodyText3_Click(sender As Object, e As RibbonControlEventArgs) Handles btnBodyText3.Click
        CIBSpecificChanges.ApplyCIBBodyTextToSelection()
    End Sub

    Private Sub btnClickSeveralSingleLevel(sender As Object, e As RibbonControlEventArgs) Handles btnEQUAL.Click, btnBULL.Click, btnHYPHEN.Click, btnLetterCapsClose.Click, btnLetterCapsFull.Click, btnLetterSmallClose.Click, btnLetterSmallFull.Click, btnArabClosing.Click, btnArabFull.Click, btnROMANdot.Click, btnRomanCapFull.Click, btnRomanSmallFull.Click, btnOrdinal.Click, btnCardinal.Click
        Dim SenderButton As RibbonButton = Nothing
        SenderButton = TryCast(sender, RibbonButton)
        If IsNothing(SenderButton) Then Exit Sub
        Select Case SenderButton.Name
            Case "btnEQUAL"
                ApplyEQUAL()
            Case "btnHYPHEN"
                ApplyHYPHEN()
            Case "btnBULL"
                ApplyBULL()
            Case "btnLetterCapsClose"
                ApplyUpperLetterwClosingParenthesis()
            Case "btnLetterCapsFull"
                ApplyUpperLetterwFullParenthesis()
            Case "btnLetterSmallClose"
                ApplyLowerLetterwClosingParenthesis()
            Case "btnLetterSmallFull"
                ApplyLowerLetterwFullClosingParenthesis()
            Case "btnArabClosing"
                ApplyArabwClosingParenthesis()
            Case "btnArabFull"
                ApplyArabwFullParenthesis()
            Case "btnROMANdot"
                ApplyUpperRomanwDot()
            Case "btnRomanCapFull"
                ApplyUpperRomanwFullParenthesis()
            Case "btnRomanSmallFull"
                ApplyLowerRomanwFullParenthesis()
            Case "btnOrdinal"
                ApplyOrdinalText()
            Case "btnCardinal"
                ApplyCardinalText()
        End Select
    End Sub

    Private Sub btnClickSeveralListNumber(sender As Object, e As RibbonControlEventArgs) Handles btnListNumber.Click, btnListNumber2.Click, btnListNumber3.Click, btnListNumber4.Click, btnListNumber5.Click
        Dim SenderButton As RibbonButton = Nothing
        SenderButton = TryCast(sender, RibbonButton)
        If IsNothing(SenderButton) Then Exit Sub
        Select Case SenderButton.Name
            Case "btnListNumber"
                Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleListNumber
            Case "btnListNumber2"
                Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleListNumber2
            Case "btnListNumber3"
                Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleListNumber3
            Case "btnListNumber4"
                Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleListNumber4
            Case "btnListNumber5"
                Globals.ThisAddIn.Application.Selection.Style = Word.WdBuiltinStyle.wdStyleListNumber5
        End Select
    End Sub

    Private Sub btnSaveAsCIB_Click(sender As Object, e As RibbonControlEventArgs)
        CIBSpecificChanges.FileSaveAsCIBName()
    End Sub
    Private Sub btnCopyCIBtemplate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCopyCIBtemplate.Click
        CIBSpecificChanges.CopyCIBTemplate()
    End Sub

    Private Sub btnRemoveDoubleParagraph_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRemoveDoubleParagraph.Click
        HPWordHelper.RemoveDoubleParagraphs()
    End Sub

    Private Sub btnHotDocsTemplate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHotDocsTemplate.Click
        Dim dialogresult As New DialogResult
        Dim ValasztottFajl = String.Empty
        Dim openFileDialog1 As New OpenFileDialog With {
            .Filter = "HotDocs sablon fájl|*.HDL",
            .Title = "Válassza ki, hogy honnan tölti be a HotDocs sablont!",
            .RestoreDirectory = True,
            .DefaultExt = "hdl",
            .Multiselect = False
        }
        If File.Exists(My.Settings.HotDocsTemplate) Then
            openFileDialog1.InitialDirectory = Path.GetDirectoryName(My.Settings.HotDocsTemplate)
        End If
        dialogresult = openFileDialog1.ShowDialog()
        If dialogresult = DialogResult.OK Then
            My.Settings.HotDocsTemplate = openFileDialog1.FileName
            My.Settings.Save()
        End If
    End Sub

    Private Sub btnHeading1_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHeading1.Click
        Globals.ThisAddIn.Application.Selection.Style = WdBuiltinStyle.wdStyleHeading1
    End Sub

    Private Sub btnHeading2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHeading2.Click
        Globals.ThisAddIn.Application.Selection.Style = WdBuiltinStyle.wdStyleHeading2
    End Sub

    Private Sub btnNumStylesEnumerate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnNumStylesEnumerate.Click
        HPWStyle.EnumListStyles()
    End Sub

    Private Sub btnMellekletHivatakozasokEllenorzeseTorlese_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMellekletHivatakozasokEllenorzeseTorlese.Click
        HPW.MellékletHivatkozásTörlése()
    End Sub

    Private Sub btnAnnexCIB_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAnnexCIB.Click
        Dim MellekletStyle As Style = Globals.ThisAddIn.Application.ActiveDocument.Styles("Mellékletcím")
        If Not MellekletStyle Is Nothing Then Globals.ThisAddIn.Application.Selection.Style = MellekletStyle
    End Sub

    Private Sub btnChangeTerm_Click(sender As Object, e As RibbonControlEventArgs) Handles btnChangeTerm.Click
        Dim TermToChangeFrom As String = InputBox("Milyen fogalmat cseréljek le?" & vbCrLf & "(Ragozott tő nélkül is próbálja meg, pl. 'Tárgy' esetén a 'Tárg' kifejezés vezet helyes eredményre", "Lecserélendő fogalom")
        Dim TermToChangeTo As String = InputBox("Milyen fogalomra cseréljük le?", "Új fogalom")
        Dim MsgBoxRes As MsgBoxResult = MsgBox("Szövegben definiált fogalom (kisbetűs változatok keresése miatt)", MsgBoxStyle.YesNo)
        Dim DefFogalom As Boolean
        If MsgBoxRes = MsgBoxResult.Yes Then DefFogalom = True Else DefFogalom = False
        If TermToChangeFrom = String.Empty OrElse TermToChangeTo = String.Empty Then Exit Sub
        HPW.ChangeTerms(TermToChangeFrom.Trim, TermToChangeTo.Trim, DefFogalom)
    End Sub

    Private Sub btnChangeAll_Click(sender As Object, e As RibbonControlEventArgs) Handles btnChangeAll.Click
        Dim TermToChangeFrom As String = InputBox("Mi legyen a `rag`, amire keressen a szövegben?", "Input")
        Dim TermToChangeTo As String = InputBox("Mire cseréljük le, mi a ragozási minta?", "Mintaalak")
        HPW.DeclinationNounInWord(TermToChangeFrom, TermToChangeTo)
    End Sub

    Private Function GetVersion() As String
        Dim Result As String = String.Empty
        If ApplicationDeployment.IsNetworkDeployed Then
            Result = "AD+" & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
#If NoHotDocs = "Y" Then
            Result += "/NoHotDoc"
#End If
#If LKT = "Y" Then
            Result += "/LKT"
#End If

        Else
            Result = "EA+" & Assembly.GetExecutingAssembly().GetName().Version.ToString
#If NoHotDocs = "Y" Then
            Result += "/NoHotDoc"
#End If
#If LKT = "Y" Then
            Result += "/LKT"
#End If
        End If
        If Not String.IsNullOrWhiteSpace(Result) Then Return Result Else Return "Ismeretlen"
    End Function
#If NoHotDocs <> "Y" Then
    Private Sub btnCustomXML_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCustomXML.Click
        MsgBox("Ennyi válasz van eltárolva: " & Globals.ThisAddIn.LastUsedAnswerCollection.Count)
        HDCust.SetXMLDocPropertiesCreateContentControlsBasedOnHDAnswers(Globals.ThisAddIn.LastUsedAnswerCollection)
    End Sub
#End If
End Class
