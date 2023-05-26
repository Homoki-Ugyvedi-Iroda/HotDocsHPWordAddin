Imports System.IO
Imports System.IO.Compression
Imports System.IO.Compression.ZipFile
Imports System.Xml
Imports System.Xml.XPath

Module Module1
    Dim LaptopPath = "C:\Users\PéterHomoki\Documents\HotDocs\Templates\"
    Dim _client_Path = "O:\Office\HotDocs\TemplateSets\"
    Dim _client_CMPPath = "___client____IT_Library"
    Dim _client_CMP = Path.Combine(_client_CMPPath, "shared.cmp")
    Dim __client__BUDLNew = Path.Combine(_client_CMPPath, "_client_list.udl")
    Dim _client_BUDLOld = Path.Combine(LaptopPath, "_client_list.udl")


    Sub Main()
        Dim FilesToProcess As String() = {_client_CMP}
        Dim XPathToFind As String = "/hd:componentLibrary/hd:components/hd:database/hd:udlFile" 'több is van
        Dim hd As XNamespace = "http://www.hotdocs.com/schemas/component_library/2009"

        'Dim Found = CopyHotDocZipsToPlace() '''már a kitömörítés és másolás funkcióra nincsen szükség, csak az XML átírásra
        'If Found = False Then Console.WriteLine("Nincsen OneDrive.zip") : Console.ReadKey() : Exit Sub

        Dim ChangedFlag As Boolean = False
        For Each toProc In FilesToProcess
            Console.WriteLine("filename: " & toProc & Environment.NewLine)
            Dim xDoc As XDocument = XDocument.Load(toProc)
            Dim namespaceManager As XmlNamespaceManager = New XmlNamespaceManager(New NameTable)
            namespaceManager.AddNamespace("hd", "http://www.hotdocs.com/schemas/component_library/2009")
            Dim ElementList As IEnumerable(Of XElement) = xDoc.XPathSelectElements(XPathToFind, namespaceManager)
            For Each xe As XElement In ElementList
                If xe.Value.Equals(_client_BUDLOld) Then
                    xe.Value = _client_BUDLNew
                    Console.WriteLine("Changed :" & xe.Value)
                    ChangedFlag = True
                End If
                If xe.Value.Equals(_client_BUDLOld) Then
                    xe.Value = _client_BUDNew
                    Console.WriteLine("Changed :" & xe.Value)
                    ChangedFlag = True
                End If
            Next
            If ChangedFlag = True Then
                If File.Exists(Path.GetFileName(toProc) + ".bak") Then
                    Console.WriteLine("Már létező backup fájl, törölve:" & Path.GetFileName(toProc) + ".bak")
                    Try
                        File.Delete(Path.GetFileName(toProc) + ".bak")
                    Catch ex As Exception
                        Console.WriteLine("Törlési hiba: " & ex.Message)
                    End Try
                End If
                Try
                    My.Computer.FileSystem.RenameFile(toProc, Path.GetFileName(toProc) + ".bak")
                Catch ex As Exception
                    Console.WriteLine("Átnevezési hiba: " & ex.Message & vbCrLf & Path.GetFileName(toProc) + ".bak")
                End Try
                xDoc.Save(toProc)
                Console.WriteLine("Átírva: " & toProc)
            Else
                Console.WriteLine("Nem módosultak a konfigurációs fájlok")
            End If
        Next
        Console.ReadKey()
    End Sub

    Function GetFirstOneDriveUnzip() As String
        Dim SearchDirectory As New IO.DirectoryInfo(My.Computer.FileSystem.SpecialDirectories.MyDocuments)
        Dim FilesArray As IO.FileSystemInfo() = SearchDirectory.GetFileSystemInfos("OneDrive_*.zip")
        If FilesArray.Count = 0 Then Return Nothing
        Return FilesArray.OrderBy(Function(f) f.CreationTime).First.FullName
    End Function

    Function CopyHotDocZipsToPlace() As Boolean
        Dim ZipFileToExtract As String = GetFirstOneDriveUnzip()
        If String.IsNullOrWhiteSpace(ZipFileToExtract) Then
            Return False
            Exit Function
        End If
        Using archive As Compression.ZipArchive = New ZipArchive(File.OpenRead(ZipFileToExtract), ZipArchiveMode.Read)
            For Each entry As Compression.ZipArchiveEntry In archive.Entries
                If entry.Name = "HotDocsToCopy1.zip" Then
                ElseIf entry.Name = "HotDocsToCopy2.zip" Then
                End If
            Next
        End Using
        Return True
    End Function

    Sub ExtractHotDocsZipsToDestination(WhichFile As String, WhereTo As String)
        If Not File.Exists(WhichFile) Then Console.WriteLine(WhichFile & " does not exist") : Exit Sub
        If Not Directory.Exists(WhereTo) Then Directory.CreateDirectory(WhereTo)
        Using archive As ZipArchive = New ZipArchive(File.OpenRead(WhichFile), ZipArchiveMode.Read)
            For Each entry As Compression.ZipArchiveEntry In archive.Entries
                If Not entry.Name.EndsWith("*.mdb", StringComparison.OrdinalIgnoreCase) Then
                    Dim FullNameToExtract = WhereTo & Path.DirectorySeparatorChar & entry.Name
                    If File.Exists(FullNameToExtract) Then File.Delete(FullNameToExtract)
                    entry.ExtractToFile(FullNameToExtract)
                    Console.WriteLine("Kitömörítve: " & entry.FullName)
                Else
                    Console.WriteLine("Nem lett felülírva mdb fájl")
                End If
            Next
        End Using
    End Sub

End Module
