Imports System
Imports System.IO
Imports System.Windows.Forms
Imports dsofile
Imports MT_PropertyExchange.Globals

''' <summary>
''' Property exchange implementation for Office 2003
''' </summary>
'<Microsoft.VisualBasic.ComClass()> _
<ComVisible(False)> _
Public Class FilePropIO2003
    Implements IFilePropIO
    Private m_oDocument As dsofile.OleDocumentPropertiesClass
    Private m_fOpenedReadOnly As Boolean

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger( _
    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)


    ''' <summary>
    ''' <para>Determine if input file is a Word document based on file extension.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to office document.</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file extension is doc or dot,
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsWord(ByVal fileName As String) As Boolean Implements MT_PropertyExchange.IFilePropIO.IsWord
        Dim fi As New FileInfo(fileName)

        Return fi.Exists AndAlso (fi.Extension.EndsWith("doc", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("dot", StringComparison.InvariantCultureIgnoreCase))
    End Function
    ''' <summary>
    ''' <para>Determine if input file is an Excel worksheet based on file
    ''' extension.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to office document</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file extension is xls or xla,
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsExcel(ByVal fileName As String) As Boolean Implements MT_PropertyExchange.IFilePropIO.IsExcel
        Dim fi As New FileInfo(fileName)
        Return fi.Exists AndAlso (fi.Extension.EndsWith("xls", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("xla", StringComparison.InvariantCultureIgnoreCase))
    End Function
    ''' <summary>
    ''' <para>Determine if input file exists.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to file</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file exists,
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsOther(ByVal fileName As String) As Boolean
        Dim fi As New FileInfo(fileName)
        Return fi.Exists AndAlso Not (IsExcel(fileName) OrElse IsWord(fileName))
    End Function

    ''' <summary>
    ''' Write property to Word 2003 document or Excel 2003 workbook
    ''' </summary>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <param name="propertyValue">Property value to write</param>
    ''' <returns>
    ''' 0 if succesful, 
    ''' <para><see cref="T:MT_PropertyExchange.Worker.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Public Function WriteProp(ByVal fileName As String, _
                               ByVal propertyName As String, _
                               ByVal propertyValue As String) As Long Implements IFilePropIO.WriteProp
        log.InfoFormat("WriteProp(fileName={0},propertyName={1},propertyValue={2}", fileName, propertyName, propertyValue)
        Dim rc As Integer = Errors.NoError
        Dim ui As Boolean = False

        'Workaround for issues in SAP Batchimport
        If propertyValue.Equals("@#DEL#@") OrElse String.IsNullOrEmpty(propertyValue) Then
            log.InfoFormat("Got signal for deleting value!")
            DeleteProp(fileName, propertyName)
            propertyValue = String.Empty
        End If

        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing

            propertyName = propertyName.Trim()

            ' Create the OleDocumentProperties object and open the file. The
            ' dsoOptionOpenReadOnlyIfNoWriteAccess allows us to open the file
            ' read/write if we have access, but go ahead and open read-only if
            ' we don't. 
            m_oDocument = New dsofile.OleDocumentProperties
            m_oDocument.Open(fileName, False, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

            If m_oDocument Is Nothing Then : Throw New System.IO.FileNotFoundException() : End If
            ' Here we can tell if file was open read-only...
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            If m_fOpenedReadOnly Then : rc = Errors.FileReadOnly : Else
                rc = WritePropBatchSingle(propertyName, propertyValue)
            End If

            m_oDocument.Save()
            m_oDocument.Close()
            m_oDocument = Nothing

        Catch ex As Exception
            rc = HandleEx(ex)
        End Try

        Return rc

    End Function

    Public Function WritePropBatchSingle(ByVal propertyName As String, _
                            ByVal propertyValue As String) As Long
        log.InfoFormat("WritePropBatchSingle(propertyName='{0}',propertyValue='{1}'", propertyName, propertyValue)

        If String.IsNullOrEmpty(propertyValue) Then
            log.DebugFormat("Got signal for writing empty value!")
            propertyValue = String.Empty
        End If



        Dim rc As Integer = Errors.NoError
        Dim bFound As Boolean
        Dim iProp As dsofile.SummaryProperties 'internal/built-in properties
        Dim oNewProp As dsofile.CustomProperty 'custom property

        Try
            ' Get the SummaryProperties (these are built-in set)...
            iProp = m_oDocument.SummaryProperties

            Select Case propertyName.ToLower()
                Case "creator", "author" : iProp.Author = propertyValue
                Case "subject" : iProp.Subject = propertyValue
                Case "category" : iProp.Category = propertyValue
                Case "comments", "description" : iProp.Comments = propertyValue
                Case "title" : iProp.Title = propertyValue
                Case "keywords" : iProp.Keywords = propertyValue
                Case "manager" : iProp.Manager = propertyValue
                Case "company" : iProp.Company = propertyValue
                Case Else
                    'customproperties
                    bFound = False
                    For Each oNewProp In m_oDocument.CustomProperties
                        If oNewProp.Name IsNot Nothing AndAlso String.Equals(oNewProp.Name.Trim(), propertyName, StringComparison.InvariantCultureIgnoreCase) Then
                            'System.Console.WriteLine(String.Format("{0}: {1} was {2} and now gets {3}", [Enum].GetName(GetType(DSOFile.dsoFilePropertyType), oNewProp.Type), oNewProp.Name, oNewProp.Value, propertyValue))
                            bFound = True
                            log.Debug("Custom property found!")
                            Select Case oNewProp.Type
                                Case 1 ' text
                                    If propertyValue.Equals(String.Empty) Then
                                        oNewProp.Value = "" 'sollte niocht vorkommen - siehe unten CX1
                                    Else
                                        oNewProp.Value = propertyValue.Trim()
                                    End If
                                Case 2 ' number
                                    If IsNumeric(propertyValue) And CLng(propertyValue) = CDec(propertyValue) Then
                                        oNewProp.Value = CLng(propertyValue)
                                    Else
                                        If CLng(propertyValue) = CDec(propertyValue) Then
                                            rc = Errors.InvalidData
                                        Else
                                            rc = Errors.InvalidData
                                        End If
                                    End If
                                Case 5 ' date
                                    If IsDate(propertyValue) Then
                                        oNewProp.Value = CDate(propertyValue)
                                    Else
                                        rc = Errors.InvalidData
                                    End If
                                Case 4 ' boolean
                                    If propertyValue.ToLower() = "true" Or propertyValue.ToLower() = "false" Then
                                        oNewProp.Value = CBool(propertyValue)
                                    Else
                                        rc = Errors.InvalidData
                                    End If
                            End Select
                            Exit For
                        End If
                    Next
                    If Not bFound Then
                        log.Debug("Adding new custom property")
                        'not found -> add (always as string typed property)
                        If propertyValue.Equals(String.Empty) Then
                            m_oDocument.CustomProperties.Add(propertyName, CStr("")) 'CX1: Leerwert soll wirklich leer sein. Prinzipiell wird bei übergabe "" die property sowieso vorher gelöscht.
                        Else
                            m_oDocument.CustomProperties.Add(propertyName, Trim(CStr(propertyValue)))
                        End If

                    End If
            End Select
        Catch ex As Exception
            rc = HandleEx(ex)
        End Try

        Return rc

    End Function

    ''' <summary>
    ''' Delete property from Office document
    ''' </summary>
    ''' <remarks>
    ''' <para>Default properties get set to an empty string, custom property are deleted
    ''' completely. Implements the <see
    ''' cref="M:MT_PropertyExchange.IFilePropIO.DeleteProp(System.String,System.String)">IFilePropIO.DeleteProp</see>
    ''' interface using the DSOFile library.</para>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <returns>
    ''' <para>0 if succesful, </para>
    ''' <para><see cref="T:MT_PropertyExchange.Worker.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Public Function DeleteProp(ByVal fileName As String, _
                              ByVal propertyName As String) As Long Implements IFilePropIO.DeleteProp
        log.InfoFormat("DeleteProp(fileName={0},propertyName={1}", fileName, propertyName)

        Dim rc As Integer = 0 'return code
        Dim iProp As dsofile.SummaryProperties 'internal/built-in properties
        Dim oNewProp As dsofile.CustomProperty 'custom property

        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing

            m_oDocument = New dsofile.OleDocumentPropertiesClass
            m_oDocument.Open(fileName, False, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

            If m_oDocument Is Nothing Then : Throw New System.IO.FileNotFoundException() : End If
            ' Here we can tell if file was open read-only...
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            If m_fOpenedReadOnly Then : rc = Errors.FileReadOnly : Else
                ' Get the SummaryProperties (these are built-in set)...
                iProp = m_oDocument.SummaryProperties
                Select Case propertyName.Trim().ToLower()
                    Case "author", "creator" : iProp.Author = String.Empty
                    Case "subject" : iProp.Subject = String.Empty
                    Case "category" : iProp.Category = String.Empty
                    Case "comments", "description" : iProp.Comments = String.Empty
                    Case "title" : iProp.Title = String.Empty
                    Case "keywords" : iProp.Keywords = String.Empty
                    Case Else
                        If (IsOther(fileName)) Then
                            Select Case propertyName.ToLower()
                                Case "manager" : iProp.Manager = String.Empty
                                Case "company" : iProp.Company = String.Empty
                            End Select
                        End If
                        For Each oNewProp In m_oDocument.CustomProperties
                            If String.Equals(oNewProp.Name.Trim(), propertyName.Trim(), StringComparison.InvariantCultureIgnoreCase) Then
                                oNewProp.Remove()
                                rc = Errors.NoError
                                Exit For
                            End If
                        Next
                End Select
            End If

            m_oDocument.Close(m_oDocument.IsDirty And Not m_fOpenedReadOnly)
            m_oDocument = Nothing

        Catch ex As Exception
            rc = HandleEx(ex)
        End Try

        Return rc
    End Function

    Private Function ExistsProperty(ByVal fileName As String, ByVal propertyName As String) As Boolean
        Dim rc As Integer = 0 'error code
        Dim bFound As Boolean = False
        Dim iProp As dsofile.SummaryProperties 'internal/built-in properties
        Dim oNewProp As dsofile.CustomProperty 'custom property

        Try
            m_oDocument = New dsofile.OleDocumentPropertiesClass
            m_oDocument.Open(fileName, False, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

            If m_oDocument Is Nothing Then : Throw New System.IO.FileNotFoundException() : End If
            ' Here we can tell if file was open read-only...
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            ' Get the SummaryProperties (these are built-in set)...
            iProp = m_oDocument.SummaryProperties
            Select Case propertyName.Trim().ToLower()
                Case "manager", "company"
                    bFound = IsOther(fileName)
                Case "author", "creator", "subject", "category", "description", "comments", "title", "keywords"
                    bFound = True
                Case Else
                    For Each oNewProp In m_oDocument.CustomProperties
                        If String.Equals(oNewProp.Name.Trim(), propertyName.Trim(), StringComparison.InvariantCultureIgnoreCase) Then
                            bFound = True
                        End If
                    Next
            End Select

            m_oDocument.Close(False)
            m_oDocument = Nothing

        Catch fnf As System.IO.FileNotFoundException
            rc = Errors.FileNotFound
        Catch ex As Exception
            rc = Errors.UnknownError
        End Try

        Return bFound
    End Function

    ''' <summary>
    ''' <para>Read property from Word 2003 document or Excel 2003 workbook</para>
    ''' </summary>
    ''' <remarks>
    ''' <para>In case of an error, the <see
    ''' cref="T:MT_PropertyExchange.Worker.Errors">Error Code</see> will be set as
    ''' system exit code.</para>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <returns>
    ''' <para><i>Property value</i> as string, or an empty string, if the propery could
    ''' not be found.</para>
    ''' </returns>
    Public Function ReadProperty(ByVal fileName As String, ByVal propertyName As String) As String Implements IFilePropIO.ReadProperty
        log.InfoFormat("ReadProperty(fileName={0},propertyName={1}", fileName, propertyName)
        Dim rc As Integer = 0 'error code
        Dim rv As String = String.Empty 'return value
        Dim bFound As Boolean = False
        Dim iProp As dsofile.SummaryProperties 'internal/built-in properties
        Dim oNewProp As dsofile.CustomProperty 'custom property

        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            fi = Nothing

            m_oDocument = New dsofile.OleDocumentPropertiesClass
            m_oDocument.Open(fileName, True, dsofile.dsoFileOpenOptions.dsoOptionDefault)
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            ' Get the SummaryProperties (these are built-in set)...
            iProp = m_oDocument.SummaryProperties

            Dim sb As New StringBuilder(String.Empty)
            Dim sl As New SortedList(Of String, String)
            If propertyName.Equals("*") Then
                sb.AppendFormat("Creator:     {0}", iProp.Author).AppendLine()
                sb.AppendFormat("Subject:     {0}", iProp.Subject).AppendLine()
                sb.AppendFormat("Category:    {0}", iProp.Category).AppendLine()
                sb.AppendFormat("Description: {0}", iProp.Comments).AppendLine()
                sb.AppendFormat("Title:       {0}", iProp.Title).AppendLine()
                sb.AppendFormat("Keywords:    {0}", iProp.Keywords).AppendLine()
                If IsOther(fileName) Then
                    sb.AppendFormat("Manager:     {0}", iProp.Manager).AppendLine()
                    sb.AppendFormat("Company:     {0}", iProp.Company).AppendLine()
                End If
                For Each oNewProp In m_oDocument.CustomProperties 'sorting !
                    sl.Add(oNewProp.Name, Convert.ToString(oNewProp.Value))
                Next
                For Each k As String In sl.Keys
                    sb.AppendFormat("{0,-20}: {1}", k, sl(k)).AppendLine()
                Next
                rv = sb.AppendLine().ToString()
            Else
                Select Case propertyName.Trim().ToLower()
                    Case "author", "creator" : rv = iProp.Author
                    Case "subject" : rv = iProp.Subject
                    Case "category" : rv = iProp.Category
                    Case "description", "comments" : rv = iProp.Comments
                    Case "title" : rv = iProp.Title
                    Case "keywords" : rv = iProp.Keywords
                    Case Else
                        If IsOther(fileName) Then
                            Select Case propertyName.Trim().ToLower()
                                Case "manager" : rv = iProp.Manager
                                Case "company" : rv = iProp.Company
                            End Select
                        End If
                        bFound = False
                        For Each oNewProp In m_oDocument.CustomProperties
                            If String.Equals(oNewProp.Name.Trim(), propertyName.Trim(), StringComparison.InvariantCultureIgnoreCase) Then
                                rv = CStr(oNewProp.Value)
                            End If
                        Next
                End Select
            End If

            m_oDocument.Close(False)
            m_oDocument = Nothing

        Catch ex As Exception
            rc = HandleEx(ex)
        End Try

        System.Environment.ExitCode = rc
        log.InfoFormat("Exit code: {0}", rc)
        log.DebugFormat("Return value: {0}", rv)
        Return rv
    End Function

    ''' <summary>
    ''' Execute a macro on an office document
    ''' </summary>
    ''' <remarks>
    ''' If an empty string is given as <i>template, </i>the procedure will sequentially
    ''' try to load:
    ''' <para></para>
    ''' <list type="bullet">
    ''' <item>
    ''' <description><i>%ProgramFiles%</i>\Siemens\MT\Office\plmUpdateServer.xla
    ''' </description></item>
    ''' <item>
    ''' <description>Office library path\plmUpdateServer.xla </description></item>
    ''' <item>
    ''' <description>Office library path\vai\plmUpdateServer.xla </description></item>
    ''' <item>
    ''' <description>Office library path\mt\plmUpdateServer.xla </description></item>
    ''' <item>
    ''' <description>Office startup macro path\plmUpdateServer.xla </description></item>
    ''' <item>
    ''' <description>Office startup macro path\vai\plmUpdateServer.xla
    ''' </description></item>
    ''' <item>
    ''' <description>Office startup macro
    ''' path\mt\plmUpdateServer.xla</description></item></list>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document.</param>
    ''' <param name="macroName">Macro to execute.</param>
    ''' <param name="template">Full path to the office add-in (currently used only with
    ''' Excel).</param>
    ''' <returns>
    ''' 0, if succesful,
    ''' <para><see cref="T:MT_PropertyExchange.Worker.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Public Function RunMacro(ByVal fileName As String, ByVal macroName As String, ByVal template As String, ByVal visible As Boolean) As Long Implements IFilePropIO.RunMacro
        If IsOther(fileName) Then Return Errors.NoError
        Return Globals.RunMacro(fileName, macroName, template, Me, visible)
    End Function

    Public Function ReadAllProperties(ByVal fileName As String) As String Implements IFilePropIO.ReadAllProperties
        log.InfoFormat("ReadAllProperties(fileName={0}", fileName)

        Dim rc As Integer = Errors.NoError 'error code
        Dim props As DocumentProperties = Nothing
        Try
            log.DebugFormat("Getting File Information")
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing


            log.DebugFormat("Initializing target DocumentProperties object")
            props = New DocumentProperties()

            Dim bFound As Boolean = False
            Dim iProp As dsofile.SummaryProperties 'internal/built-in properties
            Dim oNewProp As dsofile.CustomProperty 'custom property


            log.DebugFormat("Initializing input document object")
            m_oDocument = New dsofile.OleDocumentPropertiesClass
            log.DebugFormat("Opening file in read-only mode")
            Try
                m_oDocument.Open(fileName, True, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)
            Catch ex As Exception
                Try
                    log.WarnFormat("File open operation failed once. Trying fallback.", ex)
                    log.Debug("Fallback: Trying to open file read-only in OLE file mode.")
                    m_oDocument.Open(fileName, True, dsofile.dsoFileOpenOptions.dsoOptionOnlyOpenOLEFiles)
                Catch ex2 As Exception
                    Throw ex2
                End Try
            End Try

            If m_oDocument Is Nothing Then : Throw New System.IO.FileNotFoundException() : End If
            ' Here we can tell if file was open read-only...
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            log.Debug("Adding summary properties to target")
            iProp = m_oDocument.SummaryProperties
            props.Add("creator", iProp.Author)
            props.Add("category", iProp.Category)
            props.Add("description", iProp.Comments)
            props.Add("keywords", iProp.Keywords)
            props.Add("subject", iProp.Subject)
            props.Add("title", iProp.Title)
            Try
                If IsOther(fileName) Then
                    log.Debug("Adding company and manager OLE summary properties to target")
                    props.Add("company", iProp.Company)
                    props.Add("manager", iProp.Manager)
                End If
            Catch
                log.Warn("Could not read company and/or manager information. Skipping.")
            End Try

            log.Debug("Adding custom properties to target")
            For Each oNewProp In m_oDocument.CustomProperties
                props.Add(oNewProp.Name, oNewProp.Value)
            Next
            log.Debug("Releasing input document.")
            m_oDocument.Close(False)
            m_oDocument = Nothing
            GC.Collect()

        Catch fnf As System.IO.FileNotFoundException
            rc = Globals.HandleEx(fnf)
        Catch ex As Exception
            rc = Globals.HandleEx(ex)
        End Try
        Dim rs As String = props.Sort().ToXml()
        log.InfoFormat("Returning sorted XML result: {0}...", rs.Substring(0, 35))
        props.Clear()
        props = Nothing
        Return rs
    End Function

    Public Function WriteAllProperties(ByVal inputXML As String, ByVal fileName As String) As Long Implements IFilePropIO.WriteAllProperties
        log.InfoFormat("WriteAllPropertis(inputXLM=<xml>,fileName={0}", fileName)

        Dim irc As Integer = Errors.NoError
        Dim rc As Integer = Errors.NoError
        Dim props As DocumentProperties = DocumentProperties.FromXml(inputXML)

        Try

            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing


            m_oDocument = New dsofile.OleDocumentProperties
            m_oDocument.Open(fileName, False, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)

            If m_oDocument Is Nothing Then : Throw New System.IO.FileNotFoundException() : End If
            ' Here we can tell if file was open read-only...
            m_fOpenedReadOnly = m_oDocument.IsReadOnly

            If m_fOpenedReadOnly Then : rc = Errors.FileReadOnly : Else
                For Each prop As DocumentProperty In props
                    If prop.Value.Equals("@#del#@") Then
                        irc = DeleteProp(fileName, prop.Name)
                        m_oDocument = New dsofile.OleDocumentProperties
                        m_oDocument.Open(fileName, False, dsofile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess)
                    Else
                        irc = WritePropBatchSingle(prop.Name, prop.Value)
                    End If
                    If rc = Errors.NoError Then rc = irc
                Next
            End If

            m_oDocument.Save()
            m_oDocument.Close(True)
            m_oDocument = Nothing


        Catch ex As Exception
            rc = HandleEx(ex)
        End Try

        Return rc
    End Function

End Class

