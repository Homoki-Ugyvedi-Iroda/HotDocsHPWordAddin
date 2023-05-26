Imports System.Runtime.InteropServices
Imports NHunspell
Imports Deploy = System.Deployment.Application
Imports System.Diagnostics
Imports System.IO
#If NoHotDocs <> "Y" Then
Imports HD = HotDocs
#End If

Public Class ThisAddIn
    Public HotDocsInstalled As Boolean
    Public DictionaryPath As String
    Public ExceptionDictionary As New Dictionary(Of String, String)
    Public logger As NLog.Logger
#If NoHotDocs <> "Y" Then
    Public LastUsedAnswerCollection As HD.AnswerCollection
#End If

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        DictionaryPath = GetDictionaryPath()
        LoadExceptionDictionary()
        SetLogger()
        SetMySettings
    End Sub

    Private Sub LoadExceptionDictionary()
        Dim LineStringArray As String() = My.Resources.DeclinationConjugationExceptions.Split(vbCrLf)
        For Each line In LineStringArray
            Dim linesplit = line.Split(";")
            If linesplit.Count < 2 Then Continue For
            ExceptionDictionary.Add(linesplit(0).Replace(vbLf, ""), linesplit(1))
        Next
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub
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
    Private Sub SetLogger()
        'NLog.LogManager.ThrowExceptions = True
        'NLog.Common.InternalLogger.LogFile = "c:\\temp\internallog2.txt"
        'NLog.Common.InternalLogger.LogLevel = NLog.LogLevel.Trace
        Dim config = New NLog.Config.LoggingConfiguration
        Dim logfile = New NLog.Targets.FileTarget("logfile")
        logfile.FileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & Path.DirectorySeparatorChar & "HP_WordAddinLogFile.txt"
        config.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, logfile)
        NLog.LogManager.Configuration = config
        logger = NLog.LogManager.GetCurrentClassLogger
        logger.Info("ThisAddin_Startup")
    End Sub
    Private Sub SetMySettings()
        Dim DefaultCIBTemplateFileName = "CIB_teszt.hdl"
        Dim DefaultPicName = "jh.png"
        Dim DefaultHotDocsTemplateRemotePathLKT = "O:\Office"
        Dim DefaultPicRemotePathLKT = "O:\Client documents\CIB\"
        Dim Changed As Boolean
#If LKT = "Y" Then
        Dim DefaultHotDocsTemplatePath = Path.Combine(DefaultHotDocsTemplateRemotePathLKT, "HotDocs", "Templates", DefaultCIBTemplateFileName)
        Dim DefaultPicPath = Path.Combine(DefaultPicRemotePathLKT, DefaultPicName)
#Else
        Dim DefaultHotDocsTemplatePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "HotDocs", "Templates", DefaultCIBTemplateFileName)
        Dim DefaultPicPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), DefaultPicName)
#End If
        If String.IsNullOrWhiteSpace(My.Settings.HotDocsTemplate) Then
            My.Settings.HotDocsTemplate = DefaultHotDocsTemplatePath
            Changed = True
        End If
        If String.IsNullOrWhiteSpace(My.Settings.ApprovalPicEnglish) Then
            My.Settings.ApprovalPicEnglish = DefaultPicPath
            Changed = True
        End If
        If String.IsNullOrWhiteSpace(My.Settings.ApprovalPicHungarian) Then
            My.Settings.ApprovalPicHungarian = DefaultPicPath
            Changed = True
        End If
        If String.IsNullOrWhiteSpace(My.Settings.CIBTemplateName) Then
            My.Settings.CIBTemplateName = "CIB_HUI_iratok.dotx"
            Changed = True
        End If
        If Changed = True Then My.Settings.Save()
    End Sub
End Class
