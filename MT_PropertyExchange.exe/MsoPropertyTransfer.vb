Imports MT_PropertyExchange
Imports System
Imports System.Text
Imports CommandLine
Imports CommandLine.Utility
Imports MT_PropertyExchange.Globals

'Using a config file besides app.config/web.config
<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="MT_PropertyExchangeLog.config", Watch:=True)> 

Module MsoPropertyTransfer

    Sub Main()
        'Dim log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

        Dim rc As Integer = 0
        Dim mode As String = String.Empty
        Dim fileName, propertyName, propertyValue, macroName, template, xmlFile, delim, txtFile, encoding As String
        Dim splitCommandLine As String() = CommandLine.Utility.Arguments.SplitCommandLine(Environment.CommandLine)
        Dim args As Arguments = New Arguments(splitCommandLine)
        Dim err As StringBuilder = New StringBuilder()

        fileName = args.Single("file")
        xmlFile = args.Single("xml")
        txtFile = args.Single("txt")
        propertyName = args.Single("prop") : If propertyName Is Nothing Then propertyName = args.ForceSingle("property")
        propertyValue = args.ForceSingle("val") : If propertyValue Is Nothing Then propertyValue = args.ForceSingle("value")
        macroName = args.Single("macro")
        mode = args.Single("mode")
        template = args.Single("template")
        delim = args.Single("delim")
        encoding = args.Single("encoding")
        Dim visible As Boolean = args.IsTrue("visible")

        Dim util As New Globals() : util.InitializeLogging()
        Dim log As log4net.ILog = Globals.log

        Dim logfile As String = args.Single("logfile") : If Not String.IsNullOrEmpty(logfile) Then util.SetLogfile(logfile)
        If args.IsTrue("clearlog") Then util.ClearLogfile()
        Dim loglevel As String = args.Single("loglevel") : If Not String.IsNullOrEmpty(loglevel) Then util.SetLogLevel(loglevel)

        log.Info("Started EXE Module")
        log.DebugFormat("Command line: {0}", Environment.CommandLine)

        If args.IsTrue("dll") OrElse args.IsTrue("dllhelp") OrElse args.IsTrue("help-dll") Then
            Console.Write(My.Resources.dllhelp.ToString())
            End
        End If

        If args.IsTrue("help-enc") OrElse args.IsTrue("help-encodings") Then
            For Each enc As System.Text.EncodingInfo In System.Text.Encoding.GetEncodings()
                Console.WriteLine(enc.Name)
            Next
            End
        End If

        If args.IsTrue("h") OrElse args.IsTrue("help") OrElse args.IsTrue("?") Then
            Console.Write(My.Resources.help.ToString())
            End
        End If

        If args.Count = 1 AndAlso (args.IsTrue("v") OrElse args.IsTrue("version")) Then
            System.Console.Out.WriteLine(util.Version)
            System.Environment.Exit(0)
        End If

        If args.IsTrue("interactive") OrElse args.Count = 0 OrElse args.Count = 1 AndAlso splitCommandLine(0).ToLower.Contains(Reflection.Assembly.GetExecutingAssembly.ManifestModule.Name.ToLower) Then
            Dim dlg = New Interactivity
            dlg.ShowDialog()
            End
        End If



        If fileName Is Nothing AndAlso args.IsTrue("clearlog") Then
            End
        End If

        If fileName Is Nothing Then
            log.Error("No file provided")
            err.AppendLine("You must provide a file name(e.g. -file=<fileName>)")
            rc = Errors.ParameterMissing
        End If

        If rc = 0 Then

            'check parameters
            Select Case mode
                Case "d", "delete"
                    log.Info("Mode: Delete Property")
                    If propertyName Is Nothing Then
                        log.Error("Property name missing")
                        err.AppendLine("You must provide a property name (e.g. -prop=<propertyName>)")
                        rc = Errors.ParameterMissing
                    End If
                    If rc = 0 Then
                        log.DebugFormat("Calling DeleteProperty(fileName={0},propertyName={1}", fileName, propertyName)
                        rc = util.DeleteProperty(fileName, propertyName)
                    End If

                Case "e", "export"
                    log.Info("Mode: Export Properties")
                    Dim s As String
                    If rc = 0 Then
                        log.DebugFormat("Calling ExportProperties(fileName={0}", fileName)
                        s = util.ExportProperties(fileName)
                        System.Console.Out.WriteLine(s)
                        rc = System.Environment.ExitCode
                    End If

                Case "g", "getallproperties"
                    log.Info("Mode: List Properties")
                    Dim s As String
                    If delim Is Nothing Then
                        log.DebugFormat("Calling GetAllCusProperties(fileName={0}", fileName)
                        s = util.GetAllCusProperties(fileName)
                    Else
                        log.DebugFormat("Calling GetAllCusProperties(fileName={0},delim={1}", fileName, delim)
                        s = util.GetAllCusProperties(fileName, delim)
                    End If
                    System.Console.Out.WriteLine(s)
                    rc = System.Environment.ExitCode

                Case "i", "import"
                    log.Info("Mode: Import Properties")
                    If xmlFile Is Nothing AndAlso txtFile Is Nothing Then
                        log.Error("Import file missing")
                        err.AppendLine("You must provide an import file name (e.g. -xml=<Path to XML import file> or -txt=<Path to TXT import file>)")
                        rc = Errors.ParameterMissing
                    Else
                        If xmlFile IsNot Nothing Then
                            log.DebugFormat("Calling ImportProperties(xmlFile={0},fileName={1}", xmlFile, fileName)
                            rc = util.ImportProperties(xmlFile, fileName)
                        End If
                        If txtFile IsNot Nothing Then
                            log.DebugFormat("Calling ImportPropertiesFromText(txtFile={0},fileName={1}", txtFile, fileName)
                            If encoding Is Nothing Then
                                rc = util.ImportPropertiesFromText(txtFile, fileName)
                            Else
                                rc = util.ImportPropertiesFromText(txtFile, fileName, encoding)
                            End If
                            If rc <> 0 Then
                                System.Console.WriteLine(rc)
                            End If
                        End If
                    End If

                Case "m", "macro"
                    log.Info("Mode: Execute Macro")
                    If macroName Is Nothing Then
                        log.Error("Macro name missing")
                        err.AppendLine("You must provide a macro to be executed (e.g. -macro=<macroName>)")
                        rc = Errors.ParameterMissing
                    End If
                    If template Is Nothing Then
                        template = String.Empty
                    End If
                    If rc = 0 Then
                        log.DebugFormat("Calling RunMacro(fileName={0},macroName={1},template={2}", fileName, macroName, template)
                        rc = util.RunMacro(fileName, macroName, template, visible)
                    End If

                Case "n", "getpropertynames"
                    log.Info("Mode: Properties to textfile")
                    Dim s As String
                    If delim Is Nothing Then
                        log.DebugFormat("Calling GetAllProperties(fileName={0}", fileName)
                        s = util.GetAllProperties(fileName)
                    Else
                        log.DebugFormat("Calling GetAllProperties(fileName={0},delim={1}", fileName, delim)
                        s = util.GetAllProperties(fileName, delim)
                    End If
                    System.Console.Out.WriteLine(s)
                    rc = System.Environment.ExitCode

                Case "r", "read"
                    log.Info("Mode: Read Property")
                    Dim s As String
                    If propertyName Is Nothing Then
                        log.Error("Property name missing")
                        err.AppendLine("You must provide a property name (e.g. -prop=<propertyName>)")
                        System.Console.Out.WriteLine(util.ReadProperty(fileName, "*"))
                        rc = Errors.ParameterMissing
                    End If
                    If rc = 0 Then
                        log.DebugFormat("Calling ReadProperty(fileName={0},propertyName={1}", fileName, propertyName)
                        s = util.ReadProperty(fileName, propertyName)
                        System.Console.Out.WriteLine(s)
                        rc = System.Environment.ExitCode
                    End If

                Case "w", "write"
                    log.Info("Mode: Write Property")
                    If propertyName Is Nothing Then
                        log.Error("Property name missing")
                        err.AppendLine("You must provide a property name (e.g. -prop=<propertyName>)")
                        rc = Errors.ParameterMissing
                    End If
                    If propertyValue Is Nothing Then
                        log.Error("Property value missing")
                        err.AppendLine("You must provide a property value (e.g. -val=<propertyValue>) !")
                        rc = Errors.ParameterMissing
                    End If
                    log.DebugFormat("Calling WriteProperty(fileName={0},propertyName={1},propertyValue={2}", fileName, propertyName, propertyValue)
                    If rc = 0 Then rc = util.WriteProperty(fileName, propertyName, propertyValue)

                Case Else
                    log.Error("Mode: unknown")
                    err.AppendLine("Feature mode '" & mode & "' unknown")
                    rc = Errors.WrongMethodCode


            End Select
            util = Nothing
        End If
        GC.Collect(1, GCCollectionMode.Forced)
        If err.Length > 0 Then
            System.Console.Error.WriteLine(err.ToString)
        End If
        System.Environment.ExitCode = rc

        log.DebugFormat("Return code={0}", rc)
    End Sub

End Module
