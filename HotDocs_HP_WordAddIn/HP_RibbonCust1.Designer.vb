Partial Class HP_RibbonCust1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()
        Me.lblVersion.Label = GetVersion()
    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.tbHP = Me.Factory.CreateRibbonTab
        Me.gpWordHelper = Me.Factory.CreateRibbonGroup
        Me.btnRemoveDoubleParagraph = Me.Factory.CreateRibbonButton
        Me.btnDefinitionsSort = Me.Factory.CreateRibbonButton
        Me.btnUnlinkAllReferences = Me.Factory.CreateRibbonButton
        Me.btnUpdateAllFields = Me.Factory.CreateRibbonButton
        Me.btnMezoEllenorzes = Me.Factory.CreateRibbonButton
        Me.grpHotDocs = Me.Factory.CreateRibbonGroup
        Me.btnRagozzMindent = Me.Factory.CreateRibbonButton
        Me.btnReplaceTerms = Me.Factory.CreateRibbonButton
        Me.btnChooseTemplate = Me.Factory.CreateRibbonButton
        Me.grpCompareComments = Me.Factory.CreateRibbonGroup
        Me.btnFileAndCommentsCompare = Me.Factory.CreateRibbonButton
        Me.spbtnComments = Me.Factory.CreateRibbonSplitButton
        Me.btnCommentCounter = Me.Factory.CreateRibbonButton
        Me.btnCommentsWriteOut = Me.Factory.CreateRibbonButton
        Me.btnWriteNewComments = Me.Factory.CreateRibbonButton
        Me.grpApprovals = Me.Factory.CreateRibbonGroup
        Me.btnApprovalHu = Me.Factory.CreateRibbonButton
        Me.Tesztelésre = Me.Factory.CreateRibbonGroup
        Me.btnRagozzTesztelés = Me.Factory.CreateRibbonButton
        Me.btnAAzCsere = Me.Factory.CreateRibbonButton
        Me.btnCVarView = Me.Factory.CreateRibbonButton
        Me.btnMellekletHivatakozasokEllenorzeseTorlese = Me.Factory.CreateRibbonButton
        Me.btnChangeTerm = Me.Factory.CreateRibbonButton
        Me.btnChangeAll = Me.Factory.CreateRibbonButton
        Me.btnCustomXML = Me.Factory.CreateRibbonButton
        Me.grpSettings = Me.Factory.CreateRibbonGroup
        Me.btnCopyCIBtemplate = Me.Factory.CreateRibbonButton
        Me.btnHotDocsTemplate = Me.Factory.CreateRibbonButton
        Me.sbtnSetApproval = Me.Factory.CreateRibbonSplitButton
        Me.btnApprovalSettingEN = Me.Factory.CreateRibbonButton
        Me.btnApprovalSettingHU = Me.Factory.CreateRibbonButton
        Me.btnOwnWordSettingsHP = Me.Factory.CreateRibbonButton
        Me.lblVersion = Me.Factory.CreateRibbonLabel
        Me.tbWordStyles = Me.Factory.CreateRibbonTab
        Me.grpGeneric = Me.Factory.CreateRibbonGroup
        Me.btnCreateDefaultListStyles = Me.Factory.CreateRibbonButton
        Me.btnCreateAndApplyMLevelList = Me.Factory.CreateRibbonButton
        Me.grpHeadingStylesApply = Me.Factory.CreateRibbonGroup
        Me.ddNonNumbered = Me.Factory.CreateRibbonDropDown
        Me.btnBULL = Me.Factory.CreateRibbonButton
        Me.btnHYPHEN = Me.Factory.CreateRibbonButton
        Me.btnEQUAL = Me.Factory.CreateRibbonButton
        Me.ddNumbered = Me.Factory.CreateRibbonDropDown
        Me.btnCardinal = Me.Factory.CreateRibbonButton
        Me.btnOrdinal = Me.Factory.CreateRibbonButton
        Me.btnArabClosing = Me.Factory.CreateRibbonButton
        Me.btnArabFull = Me.Factory.CreateRibbonButton
        Me.btnROMANdot = Me.Factory.CreateRibbonButton
        Me.btnRomanSmallFull = Me.Factory.CreateRibbonButton
        Me.btnRomanCapFull = Me.Factory.CreateRibbonButton
        Me.btnLetterSmallClose = Me.Factory.CreateRibbonButton
        Me.btnLetterSmallFull = Me.Factory.CreateRibbonButton
        Me.btnLetterCapsClose = Me.Factory.CreateRibbonButton
        Me.btnLetterCapsFull = Me.Factory.CreateRibbonButton
        Me.grpBodyTextLvl = Me.Factory.CreateRibbonGroup
        Me.btnLvl1_NameNumberedBodyTextLvl = Me.Factory.CreateRibbonButton
        Me.btnLvl2_NameNumberedBodyTextLvl = Me.Factory.CreateRibbonButton
        Me.btnLvl3_NameNumberedBodyTextLvl = Me.Factory.CreateRibbonButton
        Me.btnSetLvlNameNumberedBodyTextLvl = Me.Factory.CreateRibbonButton
        Me.btnCurrentLvlBodyText = Me.Factory.CreateRibbonButton
        Me.btnCIBStyleBodyTextLvl = Me.Factory.CreateRibbonButton
        Me.grpOtherStyles = Me.Factory.CreateRibbonGroup
        Me.btnApplyBodyText = Me.Factory.CreateRibbonButton
        Me.btnBodyText2 = Me.Factory.CreateRibbonButton
        Me.btnBodyText3 = Me.Factory.CreateRibbonButton
        Me.btnListNumberApply = Me.Factory.CreateRibbonDropDown
        Me.btnListNumber = Me.Factory.CreateRibbonButton
        Me.btnListNumber2 = Me.Factory.CreateRibbonButton
        Me.btnListNumber3 = Me.Factory.CreateRibbonButton
        Me.btnListNumber4 = Me.Factory.CreateRibbonButton
        Me.btnListNumber5 = Me.Factory.CreateRibbonButton
        Me.btnHeading1 = Me.Factory.CreateRibbonButton
        Me.btnHeading2 = Me.Factory.CreateRibbonButton
        Me.btnNumStylesEnumerate = Me.Factory.CreateRibbonButton
        Me.btnAnnexCIB = Me.Factory.CreateRibbonButton
        Me.tbHP.SuspendLayout()
        Me.gpWordHelper.SuspendLayout()
        Me.grpHotDocs.SuspendLayout()
        Me.grpCompareComments.SuspendLayout()
        Me.grpApprovals.SuspendLayout()
        Me.Tesztelésre.SuspendLayout()
        Me.grpSettings.SuspendLayout()
        Me.tbWordStyles.SuspendLayout()
        Me.grpGeneric.SuspendLayout()
        Me.grpHeadingStylesApply.SuspendLayout()
        Me.grpBodyTextLvl.SuspendLayout()
        Me.grpOtherStyles.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbHP
        '
        Me.tbHP.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.tbHP.Groups.Add(Me.gpWordHelper)
        Me.tbHP.Groups.Add(Me.grpHotDocs)
        Me.tbHP.Groups.Add(Me.grpCompareComments)
        Me.tbHP.Groups.Add(Me.grpApprovals)
        Me.tbHP.Groups.Add(Me.Tesztelésre)
        Me.tbHP.Groups.Add(Me.grpSettings)
        Me.tbHP.Label = "HP"
        Me.tbHP.Name = "tbHP"
        '
        'gpWordHelper
        '
        Me.gpWordHelper.Items.Add(Me.btnRemoveDoubleParagraph)
        Me.gpWordHelper.Items.Add(Me.btnDefinitionsSort)
        Me.gpWordHelper.Items.Add(Me.btnUnlinkAllReferences)
        Me.gpWordHelper.Items.Add(Me.btnUpdateAllFields)
        Me.gpWordHelper.Items.Add(Me.btnMezoEllenorzes)
        Me.gpWordHelper.Label = "WordHelper"
        Me.gpWordHelper.Name = "gpWordHelper"
        '
        'btnRemoveDoubleParagraph
        '
        Me.btnRemoveDoubleParagraph.Label = "Remove Dbl Para"
        Me.btnRemoveDoubleParagraph.Name = "btnRemoveDoubleParagraph"
        '
        'btnDefinitionsSort
        '
        Me.btnDefinitionsSort.Label = "Definíciók rendezése"
        Me.btnDefinitionsSort.Name = "btnDefinitionsSort"
        '
        'btnUnlinkAllReferences
        '
        Me.btnUnlinkAllReferences.Label = "Kilinkelés"
        Me.btnUnlinkAllReferences.Name = "btnUnlinkAllReferences"
        '
        'btnUpdateAllFields
        '
        Me.btnUpdateAllFields.Label = " Mezőfrissítés"
        Me.btnUpdateAllFields.Name = "btnUpdateAllFields"
        '
        'btnMezoEllenorzes
        '
        Me.btnMezoEllenorzes.Label = "Mezőellenőrzés"
        Me.btnMezoEllenorzes.Name = "btnMezoEllenorzes"
        '
        'grpHotDocs
        '
        Me.grpHotDocs.Items.Add(Me.btnRagozzMindent)
        Me.grpHotDocs.Items.Add(Me.btnReplaceTerms)
        Me.grpHotDocs.Items.Add(Me.btnChooseTemplate)
        Me.grpHotDocs.Label = "HotDoc helper"
        Me.grpHotDocs.Name = "grpHotDocs"
        '
        'btnRagozzMindent
        '
        Me.btnRagozzMindent.Label = "RagozzMindet (HD)"
        Me.btnRagozzMindent.Name = "btnRagozzMindent"
        '
        'btnReplaceTerms
        '
        Me.btnReplaceTerms.Label = "Replace Terms"
        Me.btnReplaceTerms.Name = "btnReplaceTerms"
        '
        'btnChooseTemplate
        '
        Me.btnChooseTemplate.Label = "HD template kiválasztása"
        Me.btnChooseTemplate.Name = "btnChooseTemplate"
        '
        'grpCompareComments
        '
        Me.grpCompareComments.Items.Add(Me.btnFileAndCommentsCompare)
        Me.grpCompareComments.Items.Add(Me.spbtnComments)
        Me.grpCompareComments.Label = "Compare és kommentek"
        Me.grpCompareComments.Name = "grpCompareComments"
        '
        'btnFileAndCommentsCompare
        '
        Me.btnFileAndCommentsCompare.Label = "Doc & comments comp"
        Me.btnFileAndCommentsCompare.Name = "btnFileAndCommentsCompare"
        '
        'spbtnComments
        '
        Me.spbtnComments.Items.Add(Me.btnCommentCounter)
        Me.spbtnComments.Items.Add(Me.btnCommentsWriteOut)
        Me.spbtnComments.Items.Add(Me.btnWriteNewComments)
        Me.spbtnComments.Label = "Kommentek"
        Me.spbtnComments.Name = "spbtnComments"
        '
        'btnCommentCounter
        '
        Me.btnCommentCounter.Label = "Komment számláló"
        Me.btnCommentCounter.Name = "btnCommentCounter"
        Me.btnCommentCounter.ShowImage = True
        '
        'btnCommentsWriteOut
        '
        Me.btnCommentsWriteOut.Label = "Kommentek kiírása"
        Me.btnCommentsWriteOut.Name = "btnCommentsWriteOut"
        Me.btnCommentsWriteOut.ShowImage = True
        '
        'btnWriteNewComments
        '
        Me.btnWriteNewComments.Label = "Új komment kiírása"
        Me.btnWriteNewComments.Name = "btnWriteNewComments"
        Me.btnWriteNewComments.ShowImage = True
        '
        'grpApprovals
        '
        Me.grpApprovals.Items.Add(Me.btnApprovalHu)
        Me.grpApprovals.Label = "Jóváhagyás"
        Me.grpApprovals.Name = "grpApprovals"
        '
        'btnApprovalHu
        '
        Me.btnApprovalHu.Label = "Magyar"
        Me.btnApprovalHu.Name = "btnApprovalHu"
        '
        'Tesztelésre
        '
        Me.Tesztelésre.Items.Add(Me.btnRagozzTesztelés)
        Me.Tesztelésre.Items.Add(Me.btnAAzCsere)
        Me.Tesztelésre.Items.Add(Me.btnCVarView)
        Me.Tesztelésre.Items.Add(Me.btnMellekletHivatakozasokEllenorzeseTorlese)
        Me.Tesztelésre.Items.Add(Me.btnChangeTerm)
        Me.Tesztelésre.Items.Add(Me.btnChangeAll)
        Me.Tesztelésre.Items.Add(Me.btnCustomXML)
        Me.Tesztelésre.Label = "Tesztelésre"
        Me.Tesztelésre.Name = "Tesztelésre"
        '
        'btnRagozzTesztelés
        '
        Me.btnRagozzTesztelés.Label = "Ragozz teszt"
        Me.btnRagozzTesztelés.Name = "btnRagozzTesztelés"
        '
        'btnAAzCsere
        '
        Me.btnAAzCsere.Label = "A/az csere"
        Me.btnAAzCsere.Name = "btnAAzCsere"
        '
        'btnCVarView
        '
        Me.btnCVarView.Label = "CVarValue"
        Me.btnCVarView.Name = "btnCVarView"
        '
        'btnMellekletHivatakozasokEllenorzeseTorlese
        '
        Me.btnMellekletHivatakozasokEllenorzeseTorlese.Label = "MellekletHivTorles"
        Me.btnMellekletHivatakozasokEllenorzeseTorlese.Name = "btnMellekletHivatakozasokEllenorzeseTorlese"
        '
        'btnChangeTerm
        '
        Me.btnChangeTerm.Label = "ChangeTerm"
        Me.btnChangeTerm.Name = "btnChangeTerm"
        '
        'btnChangeAll
        '
        Me.btnChangeAll.Label = "ChangeAllTst"
        Me.btnChangeAll.Name = "btnChangeAll"
        '
        'btnCustomXML
        '
        Me.btnCustomXML.Label = "CustXMLchk"
        Me.btnCustomXML.Name = "btnCustomXML"
        '
        'grpSettings
        '
        Me.grpSettings.Items.Add(Me.btnCopyCIBtemplate)
        Me.grpSettings.Items.Add(Me.btnHotDocsTemplate)
        Me.grpSettings.Items.Add(Me.sbtnSetApproval)
        Me.grpSettings.Items.Add(Me.btnOwnWordSettingsHP)
        Me.grpSettings.Items.Add(Me.lblVersion)
        Me.grpSettings.Label = "Beállítások"
        Me.grpSettings.Name = "grpSettings"
        '
        'btnCopyCIBtemplate
        '
        Me.btnCopyCIBtemplate.Label = "CIB Word sablon bemásolása"
        Me.btnCopyCIBtemplate.Name = "btnCopyCIBtemplate"
        '
        'btnHotDocsTemplate
        '
        Me.btnHotDocsTemplate.Label = "Set Hotdocs Library"
        Me.btnHotDocsTemplate.Name = "btnHotDocsTemplate"
        '
        'sbtnSetApproval
        '
        Me.sbtnSetApproval.Items.Add(Me.btnApprovalSettingEN)
        Me.sbtnSetApproval.Items.Add(Me.btnApprovalSettingHU)
        Me.sbtnSetApproval.Label = "Jóváhagyó fájl beállítása"
        Me.sbtnSetApproval.Name = "sbtnSetApproval"
        '
        'btnApprovalSettingEN
        '
        Me.btnApprovalSettingEN.Label = "Beállítás angol"
        Me.btnApprovalSettingEN.Name = "btnApprovalSettingEN"
        Me.btnApprovalSettingEN.ShowImage = True
        '
        'btnApprovalSettingHU
        '
        Me.btnApprovalSettingHU.Label = "Beállítás magyar"
        Me.btnApprovalSettingHU.Name = "btnApprovalSettingHU"
        Me.btnApprovalSettingHU.ShowImage = True
        '
        'btnOwnWordSettingsHP
        '
        Me.btnOwnWordSettingsHP.Label = "HPWordSettings"
        Me.btnOwnWordSettingsHP.Name = "btnOwnWordSettingsHP"
        '
        'lblVersion
        '
        Me.lblVersion.Label = "Verzió: ismeretlen"
        Me.lblVersion.Name = "lblVersion"
        '
        'tbWordStyles
        '
        Me.tbWordStyles.Groups.Add(Me.grpGeneric)
        Me.tbWordStyles.Groups.Add(Me.grpHeadingStylesApply)
        Me.tbWordStyles.Groups.Add(Me.grpBodyTextLvl)
        Me.tbWordStyles.Groups.Add(Me.grpOtherStyles)
        Me.tbWordStyles.Label = "Word Styles"
        Me.tbWordStyles.Name = "tbWordStyles"
        '
        'grpGeneric
        '
        Me.grpGeneric.Items.Add(Me.btnCreateDefaultListStyles)
        Me.grpGeneric.Items.Add(Me.btnCreateAndApplyMLevelList)
        Me.grpGeneric.Label = "Generic (nemCIB)"
        Me.grpGeneric.Name = "grpGeneric"
        '
        'btnCreateDefaultListStyles
        '
        Me.btnCreateDefaultListStyles.Label = "CreateDefaultListStyles"
        Me.btnCreateDefaultListStyles.Name = "btnCreateDefaultListStyles"
        '
        'btnCreateAndApplyMLevelList
        '
        Me.btnCreateAndApplyMLevelList.Label = "CreateAndApplyMLevelList"
        Me.btnCreateAndApplyMLevelList.Name = "btnCreateAndApplyMLevelList"
        '
        'grpHeadingStylesApply
        '
        Me.grpHeadingStylesApply.Items.Add(Me.ddNonNumbered)
        Me.grpHeadingStylesApply.Items.Add(Me.ddNumbered)
        Me.grpHeadingStylesApply.Label = "Apply Heading Styles (nemCIB)"
        Me.grpHeadingStylesApply.Name = "grpHeadingStylesApply"
        '
        'ddNonNumbered
        '
        Me.ddNonNumbered.Buttons.Add(Me.btnBULL)
        Me.ddNonNumbered.Buttons.Add(Me.btnHYPHEN)
        Me.ddNonNumbered.Buttons.Add(Me.btnEQUAL)
        Me.ddNonNumbered.Label = "Non numbered"
        Me.ddNonNumbered.Name = "ddNonNumbered"
        '
        'btnBULL
        '
        Me.btnBULL.Label = "BULLET*"
        Me.btnBULL.Name = "btnBULL"
        '
        'btnHYPHEN
        '
        Me.btnHYPHEN.Label = "HYPHEN-"
        Me.btnHYPHEN.Name = "btnHYPHEN"
        '
        'btnEQUAL
        '
        Me.btnEQUAL.Label = "EQUAL="
        Me.btnEQUAL.Name = "btnEQUAL"
        '
        'ddNumbered
        '
        Me.ddNumbered.Buttons.Add(Me.btnCardinal)
        Me.ddNumbered.Buttons.Add(Me.btnOrdinal)
        Me.ddNumbered.Buttons.Add(Me.btnArabClosing)
        Me.ddNumbered.Buttons.Add(Me.btnArabFull)
        Me.ddNumbered.Buttons.Add(Me.btnROMANdot)
        Me.ddNumbered.Buttons.Add(Me.btnRomanSmallFull)
        Me.ddNumbered.Buttons.Add(Me.btnRomanCapFull)
        Me.ddNumbered.Buttons.Add(Me.btnLetterSmallClose)
        Me.ddNumbered.Buttons.Add(Me.btnLetterSmallFull)
        Me.ddNumbered.Buttons.Add(Me.btnLetterCapsClose)
        Me.ddNumbered.Buttons.Add(Me.btnLetterCapsFull)
        Me.ddNumbered.Label = "Numbered"
        Me.ddNumbered.Name = "ddNumbered"
        '
        'btnCardinal
        '
        Me.btnCardinal.Label = "Cardinal"
        Me.btnCardinal.Name = "btnCardinal"
        '
        'btnOrdinal
        '
        Me.btnOrdinal.Label = "Ordinal"
        Me.btnOrdinal.Name = "btnOrdinal"
        '
        'btnArabClosing
        '
        Me.btnArabClosing.Label = "Arab)"
        Me.btnArabClosing.Name = "btnArabClosing"
        '
        'btnArabFull
        '
        Me.btnArabFull.Label = "(Arab)"
        Me.btnArabFull.Name = "btnArabFull"
        '
        'btnROMANdot
        '
        Me.btnROMANdot.Label = "ROMAN."
        Me.btnROMANdot.Name = "btnROMANdot"
        '
        'btnRomanSmallFull
        '
        Me.btnRomanSmallFull.Label = "(roman)"
        Me.btnRomanSmallFull.Name = "btnRomanSmallFull"
        '
        'btnRomanCapFull
        '
        Me.btnRomanCapFull.Label = "(ROMAN)"
        Me.btnRomanCapFull.Name = "btnRomanCapFull"
        '
        'btnLetterSmallClose
        '
        Me.btnLetterSmallClose.Label = "letter)"
        Me.btnLetterSmallClose.Name = "btnLetterSmallClose"
        '
        'btnLetterSmallFull
        '
        Me.btnLetterSmallFull.Label = "(letter)"
        Me.btnLetterSmallFull.Name = "btnLetterSmallFull"
        '
        'btnLetterCapsClose
        '
        Me.btnLetterCapsClose.Label = "LETTER)"
        Me.btnLetterCapsClose.Name = "btnLetterCapsClose"
        '
        'btnLetterCapsFull
        '
        Me.btnLetterCapsFull.Label = "(LETTER)"
        Me.btnLetterCapsFull.Name = "btnLetterCapsFull"
        '
        'grpBodyTextLvl
        '
        Me.grpBodyTextLvl.Items.Add(Me.btnLvl1_NameNumberedBodyTextLvl)
        Me.grpBodyTextLvl.Items.Add(Me.btnLvl2_NameNumberedBodyTextLvl)
        Me.grpBodyTextLvl.Items.Add(Me.btnLvl3_NameNumberedBodyTextLvl)
        Me.grpBodyTextLvl.Items.Add(Me.btnSetLvlNameNumberedBodyTextLvl)
        Me.grpBodyTextLvl.Items.Add(Me.btnCurrentLvlBodyText)
        Me.grpBodyTextLvl.Items.Add(Me.btnCIBStyleBodyTextLvl)
        Me.grpBodyTextLvl.Label = "Body Text Level Set"
        Me.grpBodyTextLvl.Name = "grpBodyTextLvl"
        '
        'btnLvl1_NameNumberedBodyTextLvl
        '
        Me.btnLvl1_NameNumberedBodyTextLvl.Label = "Lvl 1"
        Me.btnLvl1_NameNumberedBodyTextLvl.Name = "btnLvl1_NameNumberedBodyTextLvl"
        '
        'btnLvl2_NameNumberedBodyTextLvl
        '
        Me.btnLvl2_NameNumberedBodyTextLvl.Label = "Lvl 2"
        Me.btnLvl2_NameNumberedBodyTextLvl.Name = "btnLvl2_NameNumberedBodyTextLvl"
        '
        'btnLvl3_NameNumberedBodyTextLvl
        '
        Me.btnLvl3_NameNumberedBodyTextLvl.Label = "Lvl 3"
        Me.btnLvl3_NameNumberedBodyTextLvl.Name = "btnLvl3_NameNumberedBodyTextLvl"
        '
        'btnSetLvlNameNumberedBodyTextLvl
        '
        Me.btnSetLvlNameNumberedBodyTextLvl.Label = "Style Body (nem CIB)"
        Me.btnSetLvlNameNumberedBodyTextLvl.Name = "btnSetLvlNameNumberedBodyTextLvl"
        '
        'btnCurrentLvlBodyText
        '
        Me.btnCurrentLvlBodyText.Label = "Current"
        Me.btnCurrentLvlBodyText.Name = "btnCurrentLvlBodyText"
        '
        'btnCIBStyleBodyTextLvl
        '
        Me.btnCIBStyleBodyTextLvl.Label = "CIB Style Body"
        Me.btnCIBStyleBodyTextLvl.Name = "btnCIBStyleBodyTextLvl"
        '
        'grpOtherStyles
        '
        Me.grpOtherStyles.Items.Add(Me.btnApplyBodyText)
        Me.grpOtherStyles.Items.Add(Me.btnBodyText2)
        Me.grpOtherStyles.Items.Add(Me.btnBodyText3)
        Me.grpOtherStyles.Items.Add(Me.btnListNumberApply)
        Me.grpOtherStyles.Items.Add(Me.btnHeading1)
        Me.grpOtherStyles.Items.Add(Me.btnHeading2)
        Me.grpOtherStyles.Items.Add(Me.btnNumStylesEnumerate)
        Me.grpOtherStyles.Items.Add(Me.btnAnnexCIB)
        Me.grpOtherStyles.Label = "Apply Other Styles"
        Me.grpOtherStyles.Name = "grpOtherStyles"
        '
        'btnApplyBodyText
        '
        Me.btnApplyBodyText.Label = "BodyText"
        Me.btnApplyBodyText.Name = "btnApplyBodyText"
        '
        'btnBodyText2
        '
        Me.btnBodyText2.Label = "Body Text 2"
        Me.btnBodyText2.Name = "btnBodyText2"
        '
        'btnBodyText3
        '
        Me.btnBodyText3.Label = "Body Text 3 (CIB)"
        Me.btnBodyText3.Name = "btnBodyText3"
        '
        'btnListNumberApply
        '
        Me.btnListNumberApply.Buttons.Add(Me.btnListNumber)
        Me.btnListNumberApply.Buttons.Add(Me.btnListNumber2)
        Me.btnListNumberApply.Buttons.Add(Me.btnListNumber3)
        Me.btnListNumberApply.Buttons.Add(Me.btnListNumber4)
        Me.btnListNumberApply.Buttons.Add(Me.btnListNumber5)
        Me.btnListNumberApply.Label = "ListNumber"
        Me.btnListNumberApply.Name = "btnListNumberApply"
        '
        'btnListNumber
        '
        Me.btnListNumber.Label = "List Number"
        Me.btnListNumber.Name = "btnListNumber"
        '
        'btnListNumber2
        '
        Me.btnListNumber2.Label = "List Number 2"
        Me.btnListNumber2.Name = "btnListNumber2"
        '
        'btnListNumber3
        '
        Me.btnListNumber3.Label = "List Number 3"
        Me.btnListNumber3.Name = "btnListNumber3"
        '
        'btnListNumber4
        '
        Me.btnListNumber4.Label = "List Number 4"
        Me.btnListNumber4.Name = "btnListNumber4"
        '
        'btnListNumber5
        '
        Me.btnListNumber5.Label = "List Number 5"
        Me.btnListNumber5.Name = "btnListNumber5"
        '
        'btnHeading1
        '
        Me.btnHeading1.Label = "Heading 1"
        Me.btnHeading1.Name = "btnHeading1"
        '
        'btnHeading2
        '
        Me.btnHeading2.Label = "Heading 2"
        Me.btnHeading2.Name = "btnHeading2"
        '
        'btnNumStylesEnumerate
        '
        Me.btnNumStylesEnumerate.Label = "NumStyles"
        Me.btnNumStylesEnumerate.Name = "btnNumStylesEnumerate"
        '
        'btnAnnexCIB
        '
        Me.btnAnnexCIB.Label = "MellékletCIB"
        Me.btnAnnexCIB.Name = "btnAnnexCIB"
        '
        'HP_RibbonCust1
        '
        Me.Name = "HP_RibbonCust1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.tbHP)
        Me.Tabs.Add(Me.tbWordStyles)
        Me.tbHP.ResumeLayout(False)
        Me.tbHP.PerformLayout()
        Me.gpWordHelper.ResumeLayout(False)
        Me.gpWordHelper.PerformLayout()
        Me.grpHotDocs.ResumeLayout(False)
        Me.grpHotDocs.PerformLayout()
        Me.grpCompareComments.ResumeLayout(False)
        Me.grpCompareComments.PerformLayout()
        Me.grpApprovals.ResumeLayout(False)
        Me.grpApprovals.PerformLayout()
        Me.Tesztelésre.ResumeLayout(False)
        Me.Tesztelésre.PerformLayout()
        Me.grpSettings.ResumeLayout(False)
        Me.grpSettings.PerformLayout()
        Me.tbWordStyles.ResumeLayout(False)
        Me.tbWordStyles.PerformLayout()
        Me.grpGeneric.ResumeLayout(False)
        Me.grpGeneric.PerformLayout()
        Me.grpHeadingStylesApply.ResumeLayout(False)
        Me.grpHeadingStylesApply.PerformLayout()
        Me.grpBodyTextLvl.ResumeLayout(False)
        Me.grpBodyTextLvl.PerformLayout()
        Me.grpOtherStyles.ResumeLayout(False)
        Me.grpOtherStyles.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tbHP As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents gpWordHelper As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnReplaceTerms As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRagozzMindent As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRagozzTesztelés As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAAzCsere As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tesztelésre As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpApprovals As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnApprovalHu As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnApprovalSettingHU As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnApprovalSettingEN As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRemoveDoubleParagraph As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCVarView As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDefinitionsSort As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUnlinkAllReferences As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnOwnWordSettingsHP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUpdateAllFields As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMezoEllenorzes As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpCompareComments As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCommentCounter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFileAndCommentsCompare As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCommentsWriteOut As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnWriteNewComments As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents tbWordStyles As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpGeneric As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCreateDefaultListStyles As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpHeadingStylesApply As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCreateAndApplyMLevelList As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnLvl1_NameNumberedBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLvl2_NameNumberedBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLvl3_NameNumberedBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSetLvlNameNumberedBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCurrentLvlBodyText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCIBStyleBodyTextLvl As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpOtherStyles As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnApplyBodyText As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBodyText2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBodyText3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddNonNumbered As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnBULL As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHYPHEN As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEQUAL As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddNumbered As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnCardinal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnOrdinal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnArabClosing As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnArabFull As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnROMANdot As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRomanSmallFull As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRomanCapFull As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLetterSmallClose As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLetterSmallFull As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLetterCapsClose As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLetterCapsFull As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents sbtnSetApproval As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents spbtnComments As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents grpHotDocs As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnChooseTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnListNumberApply As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnListNumber As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnListNumber2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnListNumber3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnListNumber4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnListNumber5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCopyCIBtemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHotDocsTemplate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHeading1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHeading2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnNumStylesEnumerate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpSettings As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnMellekletHivatakozasokEllenorzeseTorlese As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAnnexCIB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChangeTerm As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChangeAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents lblVersion As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents btnCustomXML As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property HP_RibbonCust1() As HP_RibbonCust1
        Get
            Return Me.GetRibbon(Of HP_RibbonCust1)()
        End Get
    End Property
End Class
