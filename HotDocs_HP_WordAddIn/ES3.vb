Imports System.Diagnostics
Imports System.IO
Imports System.Xml
Imports System.Xml.XPath

Public Class ES3
    Public Shared Sub CreateES3(NameWithoutExtension, PDFFileNameToInsert)
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument
        Dim ES3DestinationName = Path.Combine(Path.GetDirectoryName(ActiveDoc.Path), NameWithoutExtension & ".es3")
        Dim ContentFile As String = Path.Combine(Globals.ThisAddIn.DictionaryPath, "Empty_NonSigned.es3")
        Try
            File.Copy(ContentFile, ES3DestinationName, True)
        Catch ex As Exception
            MsgBox("Jóváhagyás során ES3 létrehozása nem sikerült. Hibaok: " & ex.Message)
            Exit Sub
        End Try
        If ManipulateES3(PDFFileNameToInsert, ES3DestinationName) = False Then MsgBox("ES3 módosítása sikertelen")
        Process.Start(Path.Combine(Environment.GetEnvironmentVariable("ProgramFiles(x86)"), "Microsec", "eszigno3", "eszigno3.exe"), ES3DestinationName)
    End Sub
    Public Shared Function ManipulateES3(PDFName As String, ES3Name As String) As Boolean
        Dim ThisApplication As Word.Application = Globals.ThisAddIn.Application
        Dim ActiveDoc As Word.Document = ThisApplication.ActiveDocument

        Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(New NameTable())
        Dim es As XNamespace = "https://www.microsec.hu/ds/e-szigno30#"
        Dim ds As XNamespace = "http://www.w3.org/2000/09/xmldsig#"
        namespaces.AddNamespace("es", es.NamespaceName)
        namespaces.AddNamespace("ds", ds.NamespaceName)
        Dim XES3 As XDocument = XDocument.Load(ES3Name)
        Dim ES3Title As String = ActiveDoc.BuiltInDocumentProperties("Title").Value & " jóváhagyása " & Now.ToString("yyMMddHHmm")

        Try
            Dim PDFAsBase64 As String = Convert.ToBase64String(File.ReadAllBytes(PDFName))
            XES3.XPathSelectElement("//es:DossierProfile/es:Title", namespaces).Value = ES3Title
            XES3.XPathSelectElement("//es:DocumentProfile/es:Title", namespaces).Value = Path.GetFileName(PDFName)
            XES3.XPathSelectElement("//ds:Object", namespaces).Value = PDFAsBase64
            XES3.Save(ES3Name)

        Catch ex As Exception
            MsgBox("ES3 átalakítási hiba. Hibakód: " & ex.Message)
            Return False
        End Try
        Return True
    End Function
End Class
