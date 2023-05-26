#If NoHotDocs <> "Y" Then
Imports HPW = HotDocs_HP_WordAddIn.HPWordHelper
Imports HD = HotDocs
Imports Microsoft.Office.Interop.Word

Public Class HDCustomizations
    Public Enum TransformationType
        Ask
        CIBSpecific
        Full
        None
    End Enum

    Friend Shared Function GetAnswerCollection(ansCall As Object) As HD.AnswerCollection
        Dim Result As HD.AnswerCollection = Nothing
        Try
            Result = TryCast(ansCall, HD.AnswerCollection)
        Catch ex As Exception
            Globals.ThisAddIn.HotDocsInstalled = False
            Globals.ThisAddIn.logger.Info("HotDocs nincsen telepítve")
        End Try
        Return Result
    End Function
    Friend Shared Function GetTransformationLevelRequired(AnswerCollectionUsed As HD.AnswerCollection) As TransformationType
        Dim Result As New TransformationType
        Try
            Dim ShallRun = AnswerCollectionUsed.Item("RunVSTOLinguistic", HD.HDVARTYPE.HD_TEXTTYPE).Value
            If Not IsNothing(ShallRun) AndAlso ShallRun = "0" Then
                Result = TransformationType.None
            ElseIf Not IsNothing(ShallRun) AndAlso ShallRun = "1" Then
                Result = TransformationType.CIBSpecific
                Globals.ThisAddIn.logger.Info("RunVSTOLinguistic=1 miatt Ügyfél#1 specifikus átalakításokkal")
            ElseIf Not IsNothing(ShallRun) AndAlso ShallRun = "2" Then
                Result = TransformationType.Full
                Globals.ThisAddIn.logger.Info("RunVSTOLinguistic=2 miatt Full transformation, de nem Ügyfél#1 specifikus")
            Else
                Result = TransformationType.Ask
            End If
        Catch ex As Exception
            Globals.ThisAddIn.logger.Info("RunVSTOLinguistic HDVARTYPE lekérdezési hiba")
            Result = TransformationType.Ask
        End Try
        Return Result
    End Function
    Friend Shared Function GetAnswerNames(AnswerCollectionUsed As HD.AnswerCollection) As List(Of String)
        Dim AnswerNames As New List(Of String)
        For Each SingleAnswer As HD.Answer In AnswerCollectionUsed
            AnswerNames.Add(SingleAnswer.Name)
        Next
        Return AnswerNames
    End Function

    Friend Shared Sub SetNonXMLDocPropertiesBasedOnHDAnswers(AnswerCollectionUsed As HD.AnswerCollection)
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim ContractTitleAsFileName, ContractTitle, ClientName, PartnerName, ClientNameAsFileName, PartnerNameAsFileName As String
        ContractTitleAsFileName = String.Empty : ContractTitle = String.Empty : ClientName = String.Empty : PartnerName = String.Empty : ClientNameAsFileName = String.Empty : PartnerNameAsFileName = String.Empty
        Try
            ContractTitle = AnswerCollectionUsed.Item("ContractTitle", HD.HDVARTYPE.HD_TEXTTYPE).Value
            ClientName = AnswerCollectionUsed.Item("ClientName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            PartnerName = AnswerCollectionUsed.Item("PartnerName", HD.HDVARTYPE.HD_TEXTTYPE).Value
            ActiveDoc.BuiltInDocumentProperties("Title").Value = ContractTitle
            ActiveDoc.BuiltInDocumentProperties("Company").Value = ClientName
            ActiveDoc.BuiltInDocumentProperties("Category").Value = PartnerName
        Catch ex As Exception
            Globals.ThisAddIn.logger.Info("HDVARTYPE Hiba at SetNonXMLDocPropertiesBasedOnHDAnswers: " & ex.Message)
        End Try
        ActiveDoc.BuiltInDocumentProperties("Subject").Value = HPW.GetFileNameFromContractDetails
    End Sub
    Private Shared Function GetBookmarkNamesSplitByDoubleUnderscore() As List(Of String)
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim BMNames As New List(Of String)
        For Each BM As Bookmark In ActiveDoc.Bookmarks
            BMNames.Add(BM.Name.Split({"__"}, StringSplitOptions.RemoveEmptyEntries).First)
        Next
        Return BMNames
    End Function

    Friend Shared Sub SetXMLDocPropertiesCreateContentControlsBasedOnHDAnswers(AnswerCollectionUsed As HD.AnswerCollection)
        If IsNothing(AnswerCollectionUsed) Then
            Globals.ThisAddIn.logger.Error("Saját AnswerCollection hiányzik")
            Exit Sub
        End If
        Dim BMList As List(Of String) = GetBookmarkNamesSplitByDoubleUnderscore()
        Dim AnswerList As List(Of String) = GetAnswerNames(AnswerCollectionUsed)
        Dim IntersectedList As List(Of String) = BMList.Intersect(AnswerList).ToList
        Globals.ThisAddIn.logger.Info("IntersectedList = " & IntersectedList.Count)
        Globals.ThisAddIn.logger.Info("BMList = " & BMList.Count)
        Dim AddinCustomXML As Office.CustomXMLPart = HPW.AddCustomXmlPartToDocument()
        If IsNothing(AddinCustomXML) Then
            Globals.ThisAddIn.logger.Error("A saját CustomXML hiányzik")
            Exit Sub
        End If
        For Each ListItem As String In IntersectedList
            Dim AnswerValue As String = String.Empty
            Try
                AnswerValue = AnswerCollectionUsed.Item(ListItem, HD.HDVARTYPE.HD_TEXTTYPE).Value
            Catch ex As Exception
                Globals.ThisAddIn.logger.Error(ListItem & " szerinti HDVARTYPE hibát okozott: " & ex.Message)
            End Try
            If String.IsNullOrWhiteSpace(AnswerValue) Then Continue For
            If ListItem = "ContractTitle" Then
                IterateBMsWithNameStarting("ContractTitle", Word.WdContentControlType.wdContentControlText, AnswerValue)
                'keresse végig BM-eket
                Continue For
            End If
            Dim AddinCustomXMLNode As Office.CustomXMLNode = Nothing
            AddinCustomXMLNode = AddinCustomXML.SelectSingleNode("/ns0:root/ns0:HotDocsReferences/ns0:" & ListItem)
            If IsNothing(AddinCustomXMLNode) Then
                Globals.ThisAddIn.logger.Info("Ez a ListItem nem létezik a CustomXML-ben: " & ListItem)
                Continue For
            End If
            AddinCustomXMLNode.Text = AnswerValue 'Nem mentjük el egyelőre külön a Node értékét, ha hibát okoz, ellenőrizni...
            IterateBMsWithNameStarting(ListItem, Word.WdContentControlType.wdContentControlText, AnswerValue, AddinCustomXMLNode)
        Next
    End Sub
    Friend Shared Sub IterateBMsWithNameStarting(ListItem As String, Type As Word.WdContentControlType, AnswerValue As String, Optional XMLNodeToMap As Office.CustomXMLNode = Nothing)
        'RichText egy soron következő ContentControlt, és összekötni az adott ContentControlt a megfelelő XML mezővel
        Dim ActiveDoc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim BMsToCheck = From BM As Bookmark In ActiveDoc.Bookmarks Where BM.Name.StartsWith(ListItem) Select BM
        For Each BM As Bookmark In BMsToCheck
            Dim counter As Integer = 0
            If BM.Name = "ContractTitle__1" Then Globals.ThisAddIn.logger.Info("Megvan ContractTitle__1 ekkor")
            Dim BMNameList = BM.Name.Split({"__"}, StringSplitOptions.RemoveEmptyEntries)
            Dim CharStyle As String = CheckFormattingRequirementsInBMName(BMNameList.Last)
            If BMNameList.Count = 0 Then Continue For
            Dim WhereToInsert As Word.Range = BM.Range
            Globals.ThisAddIn.logger.Info("BM name: " & BM.Name & ", starting at " & BM.Range.Start & " value=" & BM.Range.Text)
            Globals.ThisAddIn.logger.Info("Text to replace to: " & AnswerValue)
            Dim PlaceToInsertContentControl As Word.Range = ActiveDoc.Range(Start:=BM.Range.Start, [End]:=BM.End)
            BM.Delete()
            Dim ContentControlCreated As ContentControl = HPW.AddContentControl _
                (PlaceToInsertContentControl, Type, Title:=BMNameList.First, Style:=CharStyle)
            If IsNothing(ContentControlCreated) Then
                Globals.ThisAddIn.logger.Info("Hiba, nem jött létre a ContentControl: " & ListItem & " at " & PlaceToInsertContentControl.Start & ":" & PlaceToInsertContentControl.End)
                Continue For
            End If
            If IsNothing(XMLNodeToMap) Then 'A beépített property-kre itt kel egyedi kezelés case alapján, hiszen azok nem lehetnek Rich Text Controlok, csak simple-k
                Select Case ListItem
                    Case "ContractTitle"
                        MapToBuiltInProperty(ContentControlCreated, "title")
                End Select
                Continue For
            End If
            Dim MapSuccessful As Boolean = ContentControlCreated.XMLMapping.SetMappingByNode(XMLNodeToMap)
            If MapSuccessful Then
                Globals.ThisAddIn.logger.Info(XMLNodeToMap.XPath & " szerinti CustomXML-hez kötve a következő Control: " & ContentControlCreated.Title)
            Else
                Globals.ThisAddIn.logger.Error(XMLNodeToMap.XPath & " szerinti CustomXML-hez nem lett kötve a következő Control: " & ContentControlCreated.Title)
            End If

            counter += 1
        Next
    End Sub
    Private Shared Function CheckFormattingRequirementsInBMName(Last As String) As String
        Dim Result As String = String.Empty
        If Last = "AC" Then
            HPW.AddAllCapStyle()
            Result = HPW.AllCapStyleName
        End If
        Return Result
    End Function
    Private Shared Sub MapToBuiltInProperty(ControlToMap As ContentControl, Name As String)
        Dim CorePropertiesURI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        Dim ExtendedPropertiesURI = "http://schemas.openxmlformats.org/package/2006/metadata/extended-properties"
        Dim CoverPagePropertiesURI = "https://schemas.microsoft.com/office/2006/coverPageProps"
        Dim XPathToSet As String = "/ns"
        Dim URIToSelect As String = String.Empty
        Select Case Name
            Case "creator", "keywords", "description", "subject", "title", "category", "contentStatus"
                XPathToSet += "1:coreProperties/ns0:" & Name
                URIToSelect = CorePropertiesURI
            Case "Company", "Manager"
                XPathToSet += "0:Properties/ns0:" & Name
                URIToSelect = ExtendedPropertiesURI
            Case "PublishDate", "Abstract", "CompanyAddress", "CompanyPhone", "CompanyFax", "CompanyEmail"
                XPathToSet += "0:CoverPageProperties/ns0:" & Name
                URIToSelect = CoverPagePropertiesURI
            Case Else
                Globals.ThisAddIn.logger.Error("Ilyen típusú builtinproperty név nem lett implementálva" & Name)
                Exit Sub
        End Select
        For Each ThisCustomXmlPart As Office.CustomXMLPart In Globals.ThisAddIn.Application.ActiveDocument.CustomXMLParts
            If ThisCustomXmlPart.NamespaceURI = URIToSelect Then
                Dim MapSuccessful As Boolean
                MapSuccessful = ControlToMap.XMLMapping.SetMapping(XPath:=XPathToSet, PrefixMapping:="", Source:=ThisCustomXmlPart)
                If MapSuccessful Then
                    Globals.ThisAddIn.logger.Info(XPathToSet & " BuiltinPropertyhez kötve a " & ControlToMap.Title & " nevű contentcontrol")
                Else
                    Globals.ThisAddIn.logger.Error(XPathToSet & " BuiltinPropertyhez nincsen kötve a " & ControlToMap.Title & " nevű contentcontrol")
                End If
            End If
        Next
    End Sub

End Class
#End If
