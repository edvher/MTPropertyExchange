Imports System
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization
Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports msvbe = Microsoft.Vbe.Interop
Imports System.IO.Packaging
Imports System.Runtime.InteropServices
Imports log4net
Imports log4net.Appender
Imports log4net.Repository
Imports log4net.Core

'Using a config file besides app.config/web.config
<Assembly: log4net.Config.Repository("MT_PropertyExchangeDLL.dll")> 
<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="MT_PropertyExchangeLog.config", Watch:=True)> 

'<Guid("41A13AC0-103B-40ec-9A52-AEE2C6C846C6"), ComVisible(True), ProgId("MT_PropertyExchange.Globals"), ClassInterface(ClassInterfaceType.AutoDispatch)> _
<Microsoft.VisualBasic.ComClass("41A13AC0-103B-40ec-9A52-AEE2C6C846C6"), ComVisible(True), ProgId("MT_PropertyExchange.Globals")> _
Public Class Globals
    Public Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    ''' <summary>
    ''' FilePorpIO Error codes
    ''' </summary>
    <ComVisible(False)> _
        Public Enum Errors
        ''' <summary>
        ''' Unknown error cause.
        ''' </summary>
        UnknownError = -1
        ''' <summary>
        ''' Success. No error happened.
        ''' </summary>
        NoError = 0
        ''' <summary>
        ''' File could not be found.
        ''' </summary>
        FileNotFound = 9000
        ''' <summary>
        ''' File could be found, but is read only.
        ''' </summary>
        FileReadOnly = 9001
        ''' <summary>
        ''' Unable to convert property to valid value. 
        ''' </summary>
        InvalidData = 9002
        ''' <summary>
        ''' General automation error.
        ''' </summary>
        ApplicationError = 9003
        ''' <summary>
        ''' Unknown file extension.
        ''' </summary>
        ExtensionUnknown = 9004
        ''' <summary>
        ''' Mode type unknown. (command line only)
        ''' </summary>
        WrongMethodCode = 9005
        ''' <summary>
        ''' Parameter missing. (command line only)
        ''' </summary>
        ParameterMissing = 9006
        ''' <summary>
        ''' Office macro could not be found or executed.
        ''' </summary>
        MacroNotFound = 9007
        ''' <summary>
        ''' File in use.
        ''' </summary>
        FileInUse = 9008
        ''' <summary>
        ''' Cannot read file with encoding provided.
        ''' </summary>
        EncodingNotSupported = 9009

    End Enum

    Protected Friend Shared Function HandleEx(ByRef e As Exception) As Integer
        Dim rc As Integer = 0
        Try
            Throw e
        Catch ioe As System.IO.IOException
            System.Console.Error.WriteLine(ioe.Message)
            If TypeOf ioe Is System.IO.FileNotFoundException Then
                log.Fatal("File not found.", e)
                rc = Errors.FileNotFound
            Else
                log.Fatal("File in use.", e)
                rc = Errors.FileInUse
            End If
        Catch fro As FileReadOnlyException
            System.Console.Error.WriteLine(fro.Message)
            log.Fatal("File is read-only.", e)
            rc = Errors.FileReadOnly
        Catch appex As ApplicationException
            System.Console.Error.WriteLine(appex.Message)
            log.Fatal("Unexpected application error.", e)
            If rc = 0 Then rc = Errors.ApplicationError
        Catch comex As COMException
            System.Console.Error.WriteLine(comex.Message)
            log.Fatal("Error handling COM object.", e)
            rc = comex.ErrorCode
        Catch ex As Exception
            System.Console.Error.WriteLine(ex.Message)
            rc = Errors.UnknownError
            log.Fatal("Unexpected error", ex)
        End Try
        GC.Collect()
        log.WarnFormat("Setting Exit code to {0}", rc)
        System.Environment.ExitCode = rc
        Return rc
    End Function

    ''' <summary>
    ''' Class factory generating the correct instance object to manipulate Office
    ''' document properties.
    ''' </summary>
    ''' <param name="filename">Full path to Office document</param>
    ''' <returns>
    ''' <see cref="T:MT_PropertyExchange.FilePropIO2003">FilePropIO2003</see> or <see
    ''' cref="T:MT_PropertyExchange.FilePropIO2007">FilePropIO2007</see> object, if
    ''' succesful,
    ''' <para><see langword="null"/> otherwise</para>
    ''' </returns>
    Protected Shared Function FilePropIOFactory(ByVal filename As String) As IFilePropIO
        'log.InfoFormat("FilePropIOFactory(filename={0}", filename)
        Dim rv As IFilePropIO
        rv = New FilePropIO2007
        'log.Info("Trying Office 2k7/2k10 file format")
        If rv.IsExcel(filename) Then Return rv
        If rv.IsWord(filename) Then Return rv
        rv = Nothing
        GC.Collect()
        'Change by RAF - handle ALL files other than office 200/1020 using DSOFile
        'log.Info("Fallback: Office 2k3 file format")
        rv = New FilePropIO2003
        Return rv
    End Function

    Protected Friend Shared Function RunMacro(ByVal fileName As String, ByVal macroName As String, ByVal template As String, ByRef caller As IFilePropIO) As Integer
        Return RunMacro(fileName, macroName, template, caller, False)
    End Function

    Protected Friend Shared Function RunMacro(ByVal fileName As String, ByVal macroName As String, ByVal template As String, ByRef caller As IFilePropIO, ByVal visible As Boolean) As Integer
        Dim rc As Integer = 0
        Dim bStarted As Boolean = False
        macroName = macroName.Trim()

        log.InfoFormat("RunMacro(fileName={0},macroName={1},template={2},visible={3}", fileName, macroName, template, visible)

        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing

            If caller.IsWord(fileName) Then
                log.Info("Entering Word branch")
                Dim wdApp As Word.Application = Nothing
                If wdApp Is Nothing Then
                    log.Debug("Creating Word application object")
                    System.Console.WriteLine("Creating Word.Application instance.")
                    wdApp = New Word.Application
                    bStarted = True
                End If
                If wdApp Is Nothing Then
                    log.Fatal("Could not create Word application object.")
                    rc = Errors.MacroNotFound
                Else
                    log.DebugFormat("Word application V{0} instantiated", wdApp.Version)
                    System.Console.WriteLine("Setting properties of Application.")
                    wdApp.NormalTemplate.Saved = True : wdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                    wdApp.Visible = visible
                    wdApp.ScreenUpdating = visible
                    wdApp.DisplayAlerts = IIf(visible, Word.WdAlertLevel.wdAlertsAll, Word.WdAlertLevel.wdAlertsNone)
                    Try
                        log.DebugFormat("wdApp.Run(macroName={0}, fileName={1})", macroName, fileName)
                        wdApp.Run(macroName, fileName)
                    Catch comex As Exception
                        log.Error("Error in COM Automation. Trying a second time ...")
                        Try
                            wdApp.Documents.Open(fileName, False, False, False)
                            wdApp.Run(macroName, fileName)
                        Catch comex2 As Exception
                            log.Error("Error in COM Automation", comex2)
                            System.Console.Error.WriteLine("Cannot run macro. Macro not found or erroneous.")
                            System.Console.Error.WriteLine(comex2.Message)
                            System.Console.Error.WriteLine("Loaded templates:")
                            Try
                                log.Error("Loaded Templates")
                                For Each ai As Object In wdApp.Templates
                                    log.Error(Path.Combine(ai.Path, ai.Name))
                                    System.Console.Error.WriteLine(Path.Combine(ai.Path, ai.Name))
                                Next
                            Catch ex As Exception
                                log.Fatal("Unhandled Exception during RunMacro", ex)
                            End Try
                            rc = Errors.MacroNotFound
                        End Try
                    End Try

                    wdApp.NormalTemplate.Saved = True : wdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                    If bStarted Then
                        log.Debug("Quitting Word application")
                        wdApp.Quit(False)
                        log.Debug("Release COM Object")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp)
                        wdApp = Nothing
                    End If
                    log.Info("Exiting Word branch")
                End If
            ElseIf caller.IsExcel(fileName) Then
                Dim xlApp As Excel.Application = Nothing
                Dim xlAddin As Object = Nothing
                Dim xlDoc As Excel.Workbook = Nothing
                Dim xlaPath As String = template
                Dim bFound As Boolean = False

                Try
                    log.Info("Entering Excel branch")

                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

                    If xlApp Is Nothing Then
                        log.Debug("Creating Excel application object")
                        System.Console.WriteLine("Creating Excel.Application instance.")
                        xlApp = New Excel.Application()
                        System.Console.WriteLine(String.Format("Created Excel.Application version {0}.", xlApp.Version))
                        bStarted = True
                        xlApp.DisplayAlerts = False
                        log.DebugFormat("Excel application V{0} instantiated", xlApp.Version)
                    End If
                    If xlApp Is Nothing Then
                        log.Fatal("Could not create Excel application object.")
                        rc = Errors.MacroNotFound
                    End If
                    If xlApp IsNot Nothing AndAlso xlDoc Is Nothing Then
                        xlApp.Visible = visible
                        xlApp.ScreenUpdating = visible
                        xlApp.DisplayAlerts = visible

                        log.Debug("Building up search path inventory...")
                        System.Console.WriteLine("Building up search path inventory...")

                        Dim paths As New Collections.ArrayList()
                        paths.Add(xlApp.StartupPath)
                        paths.Add(Path.Combine(xlApp.StartupPath, "Library\pt"))
                        paths.Add(Path.Combine(xlApp.StartupPath, "Library\mt"))
                        paths.Add(Path.Combine(xlApp.StartupPath, "Library\vai"))
                        paths.Add(xlApp.LibraryPath)
                        paths.Add(Path.Combine(xlApp.LibraryPath, "pt"))
                        paths.Add(Path.Combine(xlApp.LibraryPath, "mt"))
                        paths.Add(Path.Combine(xlApp.LibraryPath, "vai"))
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "Microsoft Office\Library\pt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "Microsoft Office\Library\mt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "Microsoft Office\Library\vai")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Library\pt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Library\mt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Library\vai")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Primetals\Office")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Siemens\MT\Office")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + " (x86)\Library\pt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + " (x86)\Library\mt")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + " (x86)\Library\vai")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + " (x86)\Primetals\Office")
                        paths.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + " (x86)\Siemens\MT\Office")

                        log.Debug("Building up template file names inventory...")
                        System.Console.WriteLine("Building up template file names inventory...")
                        Dim macros As New Collections.ArrayList()
                        macros.Add("PT_Addin_01_PLM.xlam")
                        macros.Add("MT_Addin_01_PLM.xlam")
                        macros.Add("plmUpdateServer.xla")

                        log.Debug("Building path-file permutations...")
                        System.Console.WriteLine("Building path-file permutations...")
                        Dim probe As New Collections.ArrayList()
                        If Not template.Equals(String.Empty) Then
                            fi = New FileInfo(template)
                            If Not fi.Exists Then
                                log.ErrorFormat("Template ({0}) not found. Trying defaults.", template)
                                System.Console.Error.WriteLine(String.Format("Template ({0}) not found. Trying defaults.", template))
                            Else
                                probe.Add(template)
                            End If
                            fi = Nothing
                        End If
                        For Each m As String In macros
                            For Each p As String In paths
                                probe.Add(Path.Combine(p, m))
                            Next
                        Next
                        macros.Clear()
                        paths.Clear()

                        log.Debug("Start seeking addin...")
                        System.Console.WriteLine("Start seeking addin...")
                        For Each pth As String In probe

                            xlaPath = pth
                            log.DebugFormat("Probing for {0} ...", xlaPath)
                            System.Console.WriteLine(String.Format("Probing for {0} ...", xlaPath))
                            fi = New FileInfo(xlaPath)

                            'For Each addIn As Excel.AddIn In xlApp.AddIns
                            '    If addIn.Name.Equals(fi.Name, StringComparison.InvariantCultureIgnoreCase) Then
                            '        xlaPath = addIn.FullName
                            '        xlAddin = addIn
                            '        bFound = True
                            '        System.Console.Error.WriteLine("Template seems to have been loaded on startup!")
                            '        Exit For
                            '    End If
                            'Next

                            If Not bFound AndAlso fi.Exists Then
                                log.Debug("File exists. Creating Object instance.")
                                System.Console.WriteLine("File exists. Creating Object instance.")
                                Try
                                    log.DebugFormat("Trying xlApp.Workbooks.Open(xlaPath={0})", xlaPath)
                                    xlAddin = xlApp.Workbooks.Open(xlaPath)
                                Catch ex1 As COMException
                                    Try
                                        log.DebugFormat("Trying xlApp.AddIns.Add(xlaPath={0})", xlaPath)
                                        xlAddin = xlApp.AddIns.Add(xlaPath)
                                    Catch comex As COMException
                                        log.Error("COM Exception. Cannot continue:", comex)
                                        System.Console.Error.WriteLine("COM Exception. Cannot continue:")
                                        System.Console.Error.WriteLine(comex.Message)
                                        rc = comex.ErrorCode
                                        xlAddin = Nothing
                                    Catch ex As Exception
                                        log.Error("Unknown Exception. Cannot continue:", ex)
                                        System.Console.Error.WriteLine("Unknown  Exception. Cannot continue:")
                                        System.Console.Error.WriteLine(ex.Message)
                                        rc = Errors.UnknownError
                                        xlAddin = Nothing
                                    End Try
                                End Try
                                bFound = (xlAddin IsNot Nothing)
                                If bFound Then
                                    log.Debug("Addin/Template has been found. Exiting loop.")
                                    Exit For
                                End If
                            End If
                            fi = Nothing
                            If (bFound) Then
                                log.Debug("Addin/Template has been found. Exiting loop.")
                                Exit For
                            End If
                            'System.Console.Error.WriteLine(String.Empty)
                        Next
                        probe.Clear()

                        bFound = (xlAddin IsNot Nothing)
                        If bFound Then
                            log.Debug("Addin/Template has been found.")
                            System.Console.WriteLine(String.Format("Success.", xlaPath))
                            Try
                                TryCast(xlAddin, Excel.Workbook).IsAddin = True
                            Catch ex As Exception
                                'System.Console.Error.WriteLine(String.Format("Could not set IsAddin property.", xlaPath))
                            End Try
                            log.DebugFormat("Execute xlApp.Run({0}{1})", "'" + xlaPath + "'!" + macroName, fileName)
                            System.Console.WriteLine(String.Format("Executing macro [{0}] on [{1}].", "'" + xlaPath + "'!" + macroName, fileName))
                            xlApp.Run("'" + xlaPath + "'!" + macroName, fileName)
                            log.Debug("Execute finished")
                        Else
                            log.Error("Addin/Template has NOT been found.")
                            System.Console.Error.WriteLine("Template could not be loaded!")
                            rc = Errors.MacroNotFound
                        End If
                    End If
                Catch comex As COMException
                    log.Error("COM Exception. Cannot continue:", comex)
                    System.Console.Error.WriteLine("COM Exception. Cannot continue:")
                    System.Console.Error.WriteLine(comex.Message)
                    rc = comex.ErrorCode
                Finally
                    log.Debug("Release COM Object for xlAddin")
                    If bFound Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAddin)
                    xlAddin = Nothing
                    If bStarted Then
                        log.Debug("Quitting Excel")
                        xlApp.Quit()
                        log.Debug("Release COM Object for xlApp")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                        xlApp = Nothing
                    End If
                End Try
                log.Info("Exiting Excel branch")

            Else
                log.Error("Calling macro on unsupported file format")
                rc = Errors.ExtensionUnknown
            End If
        Catch ex As Exception
            rc = HandleEx(ex)
        End Try
        GC.Collect()
        log.InfoFormat("Return code=", rc)
        Return rc
    End Function

    Public Function WriteProperty(ByVal fileName As String, ByVal propertyName As String, ByVal propertyValue As String) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim util As IFilePropIO = Nothing


        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rc = util.WriteProp(fileName, propertyName, propertyValue)

        Return rc
    End Function

    Public Function ReadProperty(ByVal fileName As String, ByVal propertyName As String) As String
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rs = util.ReadProperty(fileName, propertyName)
        Return rs
    End Function

    Public Function ExportProperties(ByVal fileName As String) As String
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rs = util.ReadAllProperties(fileName)

        Dim p As DocumentProperties = DocumentProperties.FromXml(rs)
        rs = Path.GetTempFileName()
        File.Delete(rs)
        rs = System.IO.Path.ChangeExtension(rs, ".xml")
        p.SaveToFile(rs)
        Return rs
    End Function

    Public Function GetAllProperties(ByVal fileName As String) As String
        Return GetAllProperties(fileName, vbCrLf)
    End Function

    Public Function GetAllProperties(ByVal fileName As String, ByVal delim As String) As String
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rs = util.ReadAllProperties(fileName)
        Dim p As DocumentProperties = DocumentProperties.FromXml(rs)
        rs = String.Empty

        For Each prop As DocumentProperty In p
            If rs.Equals(String.Empty) Then
                rs = prop.Name
            Else
                rs = rs + delim + prop.Name
            End If
        Next
        Return rs
    End Function

    Public Function ImportProperties(ByVal importFileName As String, ByVal fileName As String) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        Dim p As DocumentProperties = DocumentProperties.LoadFromFile(importFileName)
        If p Is Nothing Then rc = Errors.InvalidData
        If rc = 0 Then rc = util.WriteAllProperties(p.ToXml(), fileName)
        Return rc
    End Function

    Public Function DeleteProperty(ByVal fileName As String, ByVal propertyName As String) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rc = util.DeleteProp(fileName, propertyName)
        Return rc
    End Function

    Public Function RunMacro(ByVal fileName As String, ByVal macroName As String, ByVal visible As Boolean) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rc = util.RunMacro(fileName, macroName, String.Empty, visible)
        Return rc
    End Function

    Public Function RunMacro(ByVal fileName As String, ByVal macroName As String, ByVal template As String, ByVal visible As Boolean) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        If rc = 0 Then rc = util.RunMacro(fileName, macroName, template, visible)
        Return rc
    End Function

    Private _prAllProp As String = String.Empty
    Public ReadOnly Property prAllProp() As String
        Get
            Return _prAllProp
        End Get
    End Property

    Public Function GetAllCusProperties(ByVal fileName As String) As String
        Return GetAllCusProperties(fileName, "@|")
    End Function

    Public Function GetAllCusProperties(ByVal fileName As String, ByVal delim As String) As String
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        ' Locate the log configuration relative to the installed assembly instead
        ' of a hard coded (32-bit) Program Files path.
        InitializeLogging()

        log.InfoFormat("GetAllCusProperties(fileName={0}, delim={1}", fileName, delim)

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            log.Error("File extension unknown. Cannot proceed.")
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
            log.Error("Returning 'File not found.'")
            Return "File not found."
        End If

        log.DebugFormat("rs = util.ReadAllProperties(fileName={0})", fileName)
        If rc = 0 Then rs = util.ReadAllProperties(fileName)
        'log.DebugFormat("rs={0}...", rs.Substring(0, 25))

        log.Debug("Creating DocumentProperties from XML")
        Dim p As DocumentProperties = DocumentProperties.FromXml(rs)

        _prAllProp = Path.GetTempFileName()
        log.DebugFormat("Temporary File: {0}", _prAllProp)

        log.Debug("Initializing XML output stream")
        Dim wr As New System.IO.StreamWriter(prAllProp, True, Encoding.UTF8)
        log.Debug("Output stream initialised. Writing properties")
        For Each prop As DocumentProperty In p
            wr.WriteLine(prop.Name + delim + prop.Value)
            log.DebugFormat("   {0}={1}", prop.Name, prop.Value)
            wr.Flush()
        Next
        log.DebugFormat("{0} properties written", p.Count)
        log.Debug("Flushing output file.")
        wr.Flush()
        log.Debug("Closing file.")
        wr.Close()
        log.Debug("File closed.")
        wr = Nothing
        p = Nothing
        util = Nothing
        log.InfoFormat("Returning file path: {0}", prAllProp)
        Return prAllProp
    End Function

    Public Function ImportPropertiesFromText(ByVal importFileName As String, ByVal fileName As String) As Integer
        Return ImportPropertiesFromText(importFileName, fileName, "utf-8") '"iso-8859-1")
    End Function

    Public Function ImportPropertiesFromText(ByVal importFileName As String, ByVal fileName As String, ByVal encodingName As String) As Integer
        Dim rc As Errors : rc = Errors.NoError
        Dim rs As String = String.Empty
        Dim util As IFilePropIO = Nothing

        util = FilePropIOFactory(fileName)
        If util Is Nothing Then
            System.Console.Error.WriteLine("File extension unknown. Cannot proceed.")
            rc = Errors.ExtensionUnknown
        End If

        Dim p As DocumentProperties = DocumentProperties.LoadFromTextFile(importFileName, encodingName)
        If System.Environment.ExitCode <> 0 Then Return System.Environment.ExitCode
        If p Is Nothing Then rc = Errors.InvalidData
        If rc = 0 Then rc = util.WriteAllProperties(p.ToXml(encodingName), fileName)
        Return rc
    End Function

    Public Sub InitializeLogging()
        Dim fi As FileInfo
        Dim settingsFileName As String = "MT_PropertyExchangeLog.config"
        Dim subdir As String = Path.DirectorySeparatorChar & "Siemens" & Path.DirectorySeparatorChar & "MT_PropertyExchange"
        Dim settingsFileNameWithPath As String = subdir & Path.DirectorySeparatorChar & settingsFileName

        Dim pf86 As String = Environment.GetEnvironmentVariable("ProgramFiles(x86)")

        fi = New FileInfo(settingsFileName)
        If Not fi.Exists Then fi = New FileInfo(Directory.GetCurrentDirectory & Path.DirectorySeparatorChar & settingsFileName)
        If Not fi.Exists Then fi = New FileInfo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location) & Path.DirectorySeparatorChar & settingsFileName)
        If Not fi.Exists AndAlso Not String.IsNullOrEmpty(pf86) Then fi = New FileInfo(pf86 & settingsFileNameWithPath)
        If Not fi.Exists Then fi = New FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & settingsFileNameWithPath)
        If Not fi.Exists Then Exit Sub

        log4net.Config.XmlConfigurator.Configure(fi)
        Dim logAppender As RollingFileAppender = GetLogAppender()
        If logAppender IsNot Nothing Then
            logAppender.ImmediateFlush = True
            logAppender.AppendToFile = True
            logAppender.ActivateOptions()
        End If
        log.Info("Logging initialized.")

    End Sub

    Private Function GetLogAppender() As RollingFileAppender
        Dim repository = TryCast(log.Logger.Repository, Hierarchy.Hierarchy)
        If repository IsNot Nothing Then
            Dim appenders As log4net.Appender.IAppender() = repository.GetAppenders()
            If appenders IsNot Nothing Then
                For Each var As log4net.Appender.IAppender In appenders
                    If TypeOf var Is RollingFileAppender Then
                        Return TryCast(var, RollingFileAppender)
                    End If
                Next
            End If
        End If
        Return Nothing
    End Function

    Public Sub SetLogfile(ByVal logfile As String)
        Dim logAppender As RollingFileAppender = GetLogAppender()
        logAppender.File = logfile
        logAppender.AppendToFile = True
        logAppender.ActivateOptions()
    End Sub
    Public Sub ClearLogfile()
        Dim logAppender As RollingFileAppender = GetLogAppender()
        logAppender.AppendToFile = True
        logAppender.ActivateOptions()
        Dim fi As FileInfo = New FileInfo(logAppender.File)
        If fi.Exists Then fi.Delete()
    End Sub
    Public Sub SetLogLevel(ByVal loglevel As String)
        Dim repository = TryCast(log.Logger.Repository, Hierarchy.Hierarchy)
        If repository IsNot Nothing Then
            Dim lvl As Level = Level.Off
            Select Case loglevel.ToUpper
                Case "DEBUG" : lvl = Level.Debug
                Case "INFO" : lvl = Level.Info
                Case "WARN" : lvl = Level.Warn
                Case "ERROR" : lvl = Level.Error
                Case "FATAL" : lvl = Level.Fatal
            End Select
            repository.Root.Level = lvl
        End If
    End Sub

    ''' <summary>
    ''' If a file "version-override.txt" exists next to the installed DLL,
    ''' its (trimmed) content is returned by Version() instead of the
    ''' built-in value. Allows adjusting the reported version without a
    ''' rebuild. GetVersion() is NOT affected - it must always return "V2".
    ''' </summary>
    Private Function VersionOverride() As String
        Try
            Dim dir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim f As String = Path.Combine(dir, "version-override.txt")
            If File.Exists(f) Then
                Dim s As String = File.ReadAllText(f).Trim()
                If s.Length > 0 Then Return s
            End If
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

    Public Function Version() As String
        Dim o As String = VersionOverride()
        If o IsNot Nothing Then Return o
        Return "2026-07-20 12:00:00"
    End Function

    ''' <summary>
    ''' Marks this DLL as the "new" PMT Office DLL. The SAP transaction
    ''' ZBATIMP (FORM check_dlls) probes for this method; the official 2017
    ''' build (OfficePropertyExchange_35, 2017-05-08) returns exactly "V2",
    ''' matching the 2-character field on the ABAP side - keep it identical.
    ''' </summary>
    Public Function GetVersion() As String

        System.Console.Error.WriteLine("V2")
        Return "V2"

    End Function
End Class
