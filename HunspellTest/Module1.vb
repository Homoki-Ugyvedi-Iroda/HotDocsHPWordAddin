Imports System.IO
Imports NHunspell
Module Module1

    Sub Main()
        Console.WriteLine("Milyen fájlt olvasson be?")
        Dim filenev = Console.ReadLine
        If filenev = String.Empty Or Not File.Exists(filenev) Then
            Console.WriteLine("empty filename")
            Exit Sub
        End If
        Dim Content As String = String.Empty
        Content = File.ReadAllText(filenev, encoding:=Text.Encoding.UTF8)
        Dim Dict = GetDictionaryPath()
        Dim resp As List(Of String)
        Console.WriteLine("exists:" & File.Exists(Dict & "HU_hu.aff"))
        Console.WriteLine(Content)

        Using MyHunspell As New Hunspell(Dict & "HU_hu.aff", Dict & "HU_hu.dic")
            resp = MyHunspell.Analyze(Content)
            Console.WriteLine(MyHunspell.Analyze("kutyám").FirstOrDefault)
        End Using
        Console.WriteLine("sorok száma:" & resp.Count)
        For Each line In resp
            Console.WriteLine(line)
        Next
        Console.ReadKey()
    End Sub
    Private Function GenerateWordFromHunspell(TermToChange As String, SampleTerm As String) As List(Of String)
        Dim Result As New List(Of String)
        Dim Dict = GetDictionaryPath()
        Using MyHunspell As New Hunspell(Dict & "HU_hu.aff", Dict & "HU_hu.dic")
            For Each stem As String In MyHunspell.Generate(TermToChange, SampleTerm)
                Result.Add(stem)
            Next
        End Using
        Return Result
    End Function
    Private Function GetDictionaryPath() As String
        'You have to check IsNetworkDeployed, because the DataDirectory property does not exist 
        '  unless you are running as a ClickOnce installed application/add-in.
        'If Deploy.ApplicationDeployment.IsNetworkDeployed Then Return Deploy.ApplicationDeployment.CurrentDeployment.DataDirectory
        'If it Then 's not data, you can look at the assembly information to find your files:
        Dim assemblyInfo As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        If (assemblyInfo IsNot Nothing) Then
            Dim UriPath As String = assemblyInfo.CodeBase
            Dim LocalPath As String = New Uri(UriPath).LocalPath
            Return Path.GetDirectoryName(LocalPath) & Path.DirectorySeparatorChar
        End If
        Return String.Empty
    End Function

End Module
