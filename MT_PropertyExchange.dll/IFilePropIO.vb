<ComVisible(False)> _
Public Interface IFilePropIO

    ''' <summary>
    ''' Write property to Office document
    ''' </summary>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <param name="propertyValue">Property value to write</param>
    ''' <returns>
    ''' 0 if succesful, 
    ''' <para><see cref="T:MT_PropertyExchange.Globals.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Function WriteProp(ByVal fileName As String, _
                       ByVal propertyName As String, _
                       ByVal propertyValue As String) As Long

    ''' <summary>
    ''' Delete property from Office document
    ''' </summary>
    ''' <remarks>
    ''' <para>Default properties get set to an empty string, custom properties are deleted
    ''' completely.</para>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <returns>
    ''' 0 if succesful, 
    ''' <para><see cref="T:MT_PropertyExchange.Globals.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Function DeleteProp(ByVal fileName As String, _
                        ByVal propertyName As String) As Long

    ''' <summary>
    ''' <para>Read property from Office document</para>
    ''' </summary>
    ''' <remarks>
    ''' <para>In case of an error, the <see
    ''' cref="T:MT_PropertyExchange.Globals.Errors">Error Code</see> will be set as
    ''' system exit code.</para>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <returns>
    ''' <para><i>Property value</i> as string, or an empty string, if the propery could
    ''' not be found.</para>
    ''' </returns>
    Function ReadProperty(ByVal fileName As String, _
                          ByVal propertyName As String) As String
    ''' <summary>
    ''' Execute a macro on an Office document
    ''' </summary>
    ''' <param name="fileName">Full path to Office document.</param>
    ''' <param name="macroName">Macro to execute.</param>
    ''' <param name="template">Full path to the Office add-in (currently used only with
    ''' Excel).</param>
    ''' <returns>
    ''' 0, if succesful,
    ''' <para><see cref="T:MT_PropertyExchange.Globals.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Function RunMacro(ByVal fileName As String, _
                      ByVal macroName As String, _
                      ByVal template As String, _
                      ByVal visible As Boolean) As Long

    ''' <summary>
    ''' <para>Determine if input file is a Word document.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to Excel document</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file extension is xlsx, xlsm ,xlam or xlsb
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsWord(ByVal fileName As String) As Boolean

    ''' <summary>
    ''' <para>Determine if input file is an Excel worksheet.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to Excel document</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file extension is xlsx, xlsm ,xlam or xlsb
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsExcel(ByVal fileName As String) As Boolean


    Function ReadAllProperties(ByVal fileName As String) As String
    Function WriteAllProperties(ByVal inputXML As String, ByVal fileName As String) As Long

End Interface
