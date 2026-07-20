Imports System
Imports System.IO
Imports System.Xml
Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports System.IO.Packaging
Imports MT_PropertyExchange.Globals

''' <summary>
''' Property exchange implementation for Office 2007 and 2010
''' </summary>
'<Microsoft.VisualBasic.ComClass()> _
<ComVisible(False)> _
Public Class FilePropIO2007
    Implements IFilePropIO

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger( _
    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    ''' <summary>
    ''' Determine if input file is a Word document based on file extension.
    ''' </summary>
    ''' <param name="fileName">Full path to Word document.</param>
    ''' <returns>
    ''' <u><font color="#0000FF">True </font></u>if file extension is docx, dotx or
    ''' docm, <see langword="false"/> otherwise.
    ''' </returns>
    Function IsWord(ByVal fileName As String) As Boolean Implements MT_PropertyExchange.IFilePropIO.IsWord
        Dim fi As New FileInfo(fileName)

        Return fi.Exists AndAlso (fi.Extension.EndsWith("docx", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("docm", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("dotx", StringComparison.InvariantCultureIgnoreCase))
    End Function

    ''' <summary>
    ''' <para>Determine if input file is an Excel worksheet based on file
    ''' extension.</para>
    ''' </summary>
    ''' <param name="fileName">Full path to Excel document</param>
    ''' <returns>
    ''' <para><u><font color="#0000FF">True </font></u>if file extension is xlsx, xlsm ,xlam or xlsb
    ''' <see langword="false"/> otherwise.</para>
    ''' </returns>
    Function IsExcel(ByVal fileName As String) As Boolean Implements MT_PropertyExchange.IFilePropIO.IsExcel
        Dim fi As New FileInfo(fileName)
        Return fi.Exists AndAlso (fi.Extension.EndsWith("xlsx", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("xlsm", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("xlam", StringComparison.InvariantCultureIgnoreCase) OrElse fi.Extension.EndsWith("xlsb", StringComparison.InvariantCultureIgnoreCase))
    End Function

    Private Enum PropertyTypes
        YesNo
        Text
        DateTime
        NumberInteger
        NumberDouble
    End Enum

    Private Function GetPropertyTypeName(ByVal propertyType As PropertyTypes) As String
        Dim propertyTypeName As String = "vt:lpwstr"

        ' Calculate the correct type:
        Select Case propertyType
            Case PropertyTypes.DateTime
                propertyTypeName = "vt:filetime"
            Case PropertyTypes.NumberInteger
                propertyTypeName = "vt:i4"
            Case PropertyTypes.NumberDouble
                propertyTypeName = "vt:r8"
            Case PropertyTypes.Text
                propertyTypeName = "vt:lpwstr"
            Case PropertyTypes.YesNo
                propertyTypeName = "vt:bool"
        End Select

        Return propertyTypeName
    End Function

    Private Function GetPropertyValueString(ByVal propertyValue As Object, ByVal propertyType As PropertyTypes) As String
        Dim propertyValueString As String = Nothing

        ' Calculate the correct type:
        Select Case propertyType
            Case PropertyTypes.DateTime
                ' Make sure you were passed a real date, 
                ' and if so, format in the correct way. The date/time 
                ' value passed in should represent a UTC date/time.
                If propertyValue.GetType() Is GetType(System.DateTime) Then
                    propertyValueString = String.Format("{0:s}Z", Convert.ToDateTime(propertyValue))
                End If
            Case PropertyTypes.NumberInteger
                If propertyValue.GetType() Is GetType(System.Int32) Then
                    propertyValueString = Convert.ToInt32(propertyValue).ToString()
                End If
            Case PropertyTypes.NumberDouble
                If propertyValue.GetType() Is GetType(System.Double) Then
                    propertyValueString = Convert.ToDouble(propertyValue).ToString()
                End If
            Case PropertyTypes.Text
                If propertyValue Is Nothing OrElse String.IsNullOrEmpty(propertyValue) Then
                    propertyValueString = ""
                Else
                    propertyValueString = Convert.ToString(propertyValue)
                End If
            Case PropertyTypes.YesNo
                If propertyValue.GetType() Is GetType(System.Boolean) Then
                    ' Must be lower case!
                    propertyValueString = Convert.ToBoolean(propertyValue).ToString().ToLower()
                End If
        End Select
        If propertyValueString Is Nothing Then
            ' If the code wasn't able to convert the 
            ' property to a valid value, throw an exception:
            Throw New InvalidDataException("Invalid parameter value.")
        End If
        Return propertyValueString
    End Function

    Private Function WDRetrieveCoreProperty(ByVal docName As String, ByVal propertyName As String) As String
        log.InfoFormat("WDRetrieveCoreProperty(docName={0},propertyName={1}", docName, propertyName)
        ' Given a document name and a core property, retrieve the value of the property.
        ' If your goal is to retrieve a single core property, use the 
        ' Package.PackageProperties property to retrieve just the property
        ' you need. For example:
        ' 
        ' myContentType = wdPackage.PackageProperties.ContentType
        '
        ' To retrieve any of the properties, this procedure uses 
        ' XML manipulation to retrieve the values directly. This requires you 
        ' to examine the contents of the core.xml part to determine the 
        ' correct namespace for each of the properties.

        ' Example:
        ' MessageBox.Show(WDRetrieveCoreProperty("C:\Samples\DocProperties.docx", "cp:keywords"))
        '
        'core.xml example:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        '  <dc:title>MyTitle</dc:title> 
        '  <dc:subject>MySubject</dc:subject> 
        '  <dc:creator>MyAuthor</dc:creator> 
        '  <cp:keywords>MyKeywords</cp:keywords> 
        '  <dc:description>A simple comment</dc:description> 
        '  <cp:lastModifiedBy>Modifier name</cp:lastModifiedBy> 
        '  <cp:revision>1</cp:revision> 
        '  <dcterms:created xsi:type="dcterms:W3CDTF">2011-06-16T14:10:00Z</dcterms:created> 
        '  <dcterms:modified xsi:type="dcterms:W3CDTF">2011-06-16T14:15:00Z</dcterms:modified> 
        '  <cp:category>MyCategory</cp:category> 
        '</cp:coreProperties>

        'custom Properties look like
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Customer">
        ' <vt:lpwstr>Customer</vt:lpwstr> 
        ' </property>
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="CustomerProjectDescription">
        ' <vt:lpwstr>Customer project description</vt:lpwstr> 
        '</property>
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="DepartmentDescription">
        ' <vt:lpwstr>Project leading department description</vt:lpwstr> 
        '</property>
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Title2">
        ' <vt:lpwstr>2012-07-26 16:35:53.79</vt:lpwstr> 
        '</property>



        Const corePropertiesRelationshipType As String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        Const corePropertiesSchema As String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        Const dcPropertiesSchema As String = "http://purl.org/dc/elements/1.1/"
        Const dcTermsPropertiesSchema As String = "http://purl.org/dc/terms/"

        Dim propertyValue As String = String.Empty
        Dim corePropertiesPart As PackagePart = Nothing

        Using wdPackage As Package = Package.Open(docName, FileMode.Open, FileAccess.ReadWrite)
            ' Get the main document part (document.xml).
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(corePropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                corePropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one core properties part.
                Exit For
            Next

            ' Manage namespaces to perform Xml XPath queries.
            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("cp", corePropertiesSchema)
            nsManager.AddNamespace("dc", dcPropertiesSchema)
            nsManager.AddNamespace("dcterms", dcTermsPropertiesSchema)

            ' Get the properties from the package.
            Dim xdoc As XmlDocument = New XmlDocument(nt)

            ' Load the XML in the part into an XmlDocument instance:
            xdoc.Load(corePropertiesPart.GetStream())

            If propertyName.Equals("comments") Then propertyName = "description"
            If propertyName.Equals("author") Then propertyName = "creator"


            Dim searchString As String = String.Format("//cp:coreProperties/dc:{0}|//cp:coreProperties/cp:{0}", propertyName, propertyName)
            Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManager)
            If Not xNode Is Nothing Then
                propertyValue = xNode.InnerText
            End If
        End Using
        Return propertyValue
    End Function

    Private Function WDRetrieveCustomProperty(ByVal docName As String, ByVal propertyName As String) As String
        log.InfoFormat("WDRetrieveCustomProperty(docName={0},propertyName={1}", docName, propertyName)
        ' Given a document name and a core property, retrieve the value of the property.
        '
        'sample custom.xml:
        ' <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
        ' <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        '     <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MT_CustText">
        '        <vt:lpwstr>Custom Text</vt:lpwstr> 
        '     </property>
        '     <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MT_CustDate">
        '        <vt:filetime>2011-06-19T22:00:00Z</vt:filetime> 
        '     </property>
        '     <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="MT_CustNo">
        '        <vt:i4>1234</vt:i4> 
        '     </property>
        '     <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="MT_CustBool">
        '        <vt:bool>false</vt:bool> 
        '     </property>
        '</Properties>

        Const customPropertiesRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        Const customPropertiesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Const customVTypesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        Dim propertyValue As String = String.Empty
        Dim customPropertiesPart As PackagePart = Nothing

        Using wdPackage As Package = Package.Open(docName, FileMode.Open, FileAccess.ReadWrite)
            ' Get the main document part (document.xml).
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(customPropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                customPropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one custom properties part.
                Exit For
            Next

            ' There may not be a custom properties part.
            If customPropertiesPart IsNot Nothing Then
                ' Manage namespaces to perform Xml XPath queries.
                Dim nt As New NameTable()
                Dim nsManager As New XmlNamespaceManager(nt)
                nsManager.AddNamespace("d", customPropertiesSchema)
                nsManager.AddNamespace("vt", customVTypesSchema)

                ' Get the properties from the package.
                Dim xdoc As XmlDocument = New XmlDocument(nt)

                ' Load the XML in the part into an XmlDocument instance:
                xdoc.Load(customPropertiesPart.GetStream())

                Dim searchString As String = String.Format("d:Properties/d:property[translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='{0}']", propertyName.ToLower())
                Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManager)
                If Not xNode Is Nothing Then
                    propertyValue = xNode.InnerText
                End If
            End If
        End Using
        Return propertyValue
    End Function

    Private Function WDSetCoreProperty(ByVal docName As String, ByVal propertyName As String, _
     ByVal propertyValue As String) As String
        log.InfoFormat("WDSetCoreProperty(docName={0},propertyName={1},propertyValue={2}", docName, propertyName, propertyValue)

        ' Given a document name, a property name, and a value, update the document.
        ' Note that you cannot set the value of a property that doesn't already
        ' exist within the core.xml part. If you successfully updated the property
        ' value, return its old value.

        ' Attempting to modify a non-existent property raises an exception.
        ' Note that property names are case-sensitive.

        ' If your goal is to set a single core property, use the 
        ' Package.PackageProperties property to set just the property
        ' you need. For example:
        ' 
        ' wdPackage.PackageProperties.Creator = "Tom Smith"
        '
        ' To set any of the properties, this procedure uses 
        ' XML manipulation to set the values directly. This requires you 
        ' to examine the contents of the core.xml part to determine the 
        ' correct namespace for each of the properties.

        ' Example:
        ' MessageBox.Show(WDSetCoreProperty("C:\Samples\DocProperties.docx", "cp:keywords", "Demo keywords"))


        Const corePropertiesSchema As String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        Const dcPropertiesSchema As String = "http://purl.org/dc/elements/1.1/"
        Const corePropertiesRelationshipType As String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

        Dim corePropertiesPart As PackagePart = Nothing
        Dim retVal As String = Nothing

        Using wdPackage As Package = Package.Open(docName, FileMode.Open, FileAccess.ReadWrite)
            ' Get the main document part (document.xml).
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(corePropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                corePropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one core properties part.
                Exit For
            Next

            ' Manage namespaces to perform Xml XPath queries.
            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("cp", corePropertiesSchema)
            nsManager.AddNamespace("dc", dcPropertiesSchema)

            ' Get the properties from the package.
            Dim xdoc As XmlDocument = New XmlDocument(nt)

            ' Load the XML in the part into an XmlDocument instance:
            xdoc.Load(corePropertiesPart.GetStream())

            Dim searchString As String = String.Format("//cp:coreProperties/dc:{0}|//cp:coreProperties/cp:{0}", propertyName.ToLower, propertyName.ToLower)
            Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManager)
            If xNode Is Nothing Then
                ' Trying to set the value of a property that 
                ' doesn't exist? Throw an exception:
                Throw New ArgumentException("Invalid property name.")
            Else
                ' Get the current value:
                retVal = xNode.InnerText

                ' Now update the value:
                xNode.InnerText = propertyValue

                ' Save the properties XML back to its part.
                xdoc.Save(corePropertiesPart.GetStream(FileMode.Create, FileAccess.Write))
                wdPackage.Close()
            End If
        End Using
        Return retVal
    End Function

    Private Function WDSetCustomProperty(ByVal docName As String, ByVal propertyName As String, _
     ByVal propertyValue As Object, ByVal propertyType As PropertyTypes) As Boolean
        log.InfoFormat("WDSetCustomProperty(docName={0},propertyName={1},propertyValue={2},propertyType={3}", docName, propertyName, propertyValue, GetPropertyTypeName(propertyType))
        ' Given a document name, a property name/value, and the property type, add a custom property 
        ' to a document. Return True if the property was added/updated, or False if the property couldn't be updated.
        ' The function's return value is true if the code could add/update the property,
        ' and false otherwise.

        ' If the custom.xml part isn't already added to the document, add it. 
        ' If the custom.xml part is there, but the property isn't, add it.
        ' If the property exists and is of the same type, replace the value.
        ' If the property exists and is of a different type, update the existing property. 

        Const customPropertiesRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        Const customPropertiesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Const customVTypesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        Dim retVal As Boolean = False
        Dim propertyTypeName As String = GetPropertyTypeName(propertyType)
        Dim propertyValueString As String = GetPropertyValueString(propertyValue, propertyType)

        Using wdPackage As Package = Package.Open(docName, FileMode.Open, FileAccess.ReadWrite)
            ' Work with the custom properties part:
            Dim customPropsPart As PackagePart = Nothing

            ' Get the custom part (custom.xml). It may not exist.
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(customPropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                customPropsPart = wdPackage.GetPart(documentUri)
                ' There is only one custom properties part, if it exists at all.
                Exit For
            Next

            ' Manage namespaces to perform Xml XPath queries.
            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("d", customPropertiesSchema)
            nsManager.AddNamespace("vt", customVTypesSchema)

            Dim customPropsUri As New Uri("/docProps/custom.xml", UriKind.Relative)
            Dim customPropsDoc As XmlDocument = Nothing
            Dim rootNode As XmlNode = Nothing

            ' There may not be a custom properties part.
            If customPropsPart Is Nothing Then
                customPropsDoc = New XmlDocument(nt)

                ' Part doesn't exist. Create it now.
                customPropsPart = wdPackage.CreatePart(customPropsUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml")

                ' Set up the rudimentary custom part:
                rootNode = customPropsDoc.CreateElement("Properties", customPropertiesSchema)
                rootNode.Attributes.Append(customPropsDoc.CreateAttribute("xmlns:vt"))
                rootNode.Attributes("xmlns:vt").Value = customVTypesSchema

                customPropsDoc.AppendChild(rootNode)

                ' Create the document's relationship to the new custom properties part:
                wdPackage.CreateRelationship(customPropsUri, TargetMode.Internal, customPropertiesRelationshipType)
            Else
                ' Load the contents of the custom properties part into an XML document.
                customPropsDoc = New XmlDocument(nt)
                customPropsDoc.Load(customPropsPart.GetStream())

                rootNode = customPropsDoc.DocumentElement
            End If


            ' Now that you've got a reference to an XmlDocument object that 
            ' corresponds to the custom properties part, 
            ' check to see if the required property is already there:
            Dim searchString As String = String.Format("d:Properties/d:property[translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='{0}']", propertyName.ToLower())
            Dim node As XmlNode = customPropsDoc.SelectSingleNode(searchString, nsManager)

            ' If you didn't find the node, add it. If you found it, and the type is 
            ' different than the new property, delete it. If you found it, and the type 
            ' is the same, replace the property value. Otherwise, add a new 
            ' element with the new value and type.

            Dim valueNode As XmlNode = Nothing

            If node IsNot Nothing Then
                ' You found the node. Now check its type:
                If node.HasChildNodes Then
                    valueNode = node.ChildNodes(0)
                    If valueNode IsNot Nothing Then
                        Dim typeName As String = valueNode.Name
                        If propertyTypeName = typeName Then
                            ' The types are the same. Simply 
                            ' replace the value of the node:
                            valueNode.InnerText = propertyValueString
                            ' If the property existed, and its type
                            ' hasn't changed, you're done:
                            retVal = True
                        Else
                            ' Types are different. Delete the node, 
                            ' and clear the node variable:
                            node.ParentNode.RemoveChild(node)
                            node = Nothing
                        End If
                    End If
                End If
            End If

            ' The previous block of code may have cleared the value in the 
            ' variable named node:
            If node Is Nothing Then
                ' Either you didn't find the node, or you 
                ' found it, its type was incorrect, and you deleted it.
                ' Either way, you need to create the new property node now.

                ' Find the highest existing "pid" value.
                ' The default value for the "pid" attribute is "2".
                Dim pidValue As String = "2"

                Dim propertiesNode As XmlNode = customPropsDoc.DocumentElement
                If propertiesNode.HasChildNodes Then
                    Dim lastNode As XmlNode = propertiesNode.LastChild
                    If lastNode IsNot Nothing Then
                        Dim pidAttr As XmlAttribute = lastNode.Attributes("pid")
                        If Not pidAttr Is Nothing Then
                            pidValue = pidAttr.Value
                            ' Increment pidValue, so that the new property
                            ' gets a pid value one higher. This value should be 
                            ' numeric, but it never hurt so to confirm:
                            Dim value As Integer
                            If Integer.TryParse(pidValue, value) Then
                                pidValue = Convert.ToString(value + 1)
                            End If
                        End If
                    End If
                End If

                node = customPropsDoc.CreateElement("property", customPropertiesSchema)
                node.Attributes.Append(customPropsDoc.CreateAttribute("name"))
                node.Attributes("name").Value = propertyName

                node.Attributes.Append(customPropsDoc.CreateAttribute("fmtid"))
                node.Attributes("fmtid").Value = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

                node.Attributes.Append(customPropsDoc.CreateAttribute("pid"))
                node.Attributes("pid").Value = pidValue

                valueNode = customPropsDoc.CreateElement(propertyTypeName, customVTypesSchema)
                valueNode.InnerText = propertyValueString
                node.AppendChild(valueNode)
                rootNode.AppendChild(node)
                retVal = True
            End If

            ' Save the properties XML back to its part.
            Dim settings As New XmlWriterSettings()
            settings.Indent = False
            settings.ConformanceLevel = ConformanceLevel.Document
            settings.IndentChars = ""
            settings.Encoding = New System.Text.UTF8Encoding()
            settings.OmitXmlDeclaration = False
            settings.NewLineOnAttributes = False
            settings.NewLineHandling = NewLineHandling.Entitize
            settings.NewLineChars = ""

            Dim writer As XmlWriter = XmlWriter.Create(customPropsPart.GetStream(FileMode.Create, FileAccess.Write), settings)
            customPropsDoc.Save(writer) 'customPropsPart.GetStream(FileMode.Create, FileAccess.Write))
            wdPackage.Close()
        End Using
        Return retVal
    End Function

    Private Function WDDeleteCustomProperty(ByVal docName As String, ByVal propertyName As String) As Boolean
        log.InfoFormat("WDDeleteCustomProperty(docName={0},propertyName={1}", docName, propertyName)
        ' Given a document name, a property name/value, and the property type, add a custom property 
        ' to a document. Return True if the property was added/updated, or False if the property couldn't be updated.
        ' The function's return value is true if the code could add/update the property,
        ' and false otherwise.

        ' If the custom.xml part isn't already added to the document, add it. 
        ' If the custom.xml part is there, but the property isn't, add it.
        ' If the property exists and is of the same type, replace the value.
        ' If the property exists and is of a different type, update the existing property. 

        'Const documentRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Const customPropertiesRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        Const customPropertiesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Const customVTypesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        Dim retVal As Boolean = False
        'Dim documentPart As PackagePart = Nothing


        Using wdPackage As Package = Package.Open(docName, FileMode.Open, FileAccess.ReadWrite)
            '' Get the main document part (document.xml).
            'For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(documentRelationshipType)
            '    Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
            '    documentPart = wdPackage.GetPart(documentUri)
            '    ' There is only one document.
            '    Exit For
            'Next

            ' Work with the custom properties part:
            Dim customPropsPart As PackagePart = Nothing

            ' Get the custom part (custom.xml). It may not exist.
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(customPropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                customPropsPart = wdPackage.GetPart(documentUri)
                ' There is only one custom properties part, if it exists at all.
                Exit For
            Next

            ' Manage namespaces to perform Xml XPath queries.
            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("d", customPropertiesSchema)
            nsManager.AddNamespace("vt", customVTypesSchema)

            Dim customPropsUri As New Uri("/docProps/custom.xml", UriKind.Relative)
            Dim customPropsDoc As XmlDocument = Nothing
            Dim rootNode As XmlNode = Nothing

            ' There may not be a custom properties part.
            If customPropsPart Is Nothing Then
                customPropsDoc = New XmlDocument(nt)

                ' Part doesn't exist. Create it now.
                customPropsPart = wdPackage.CreatePart(customPropsUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml")

                ' Set up the rudimentary custom part:
                rootNode = customPropsDoc.CreateElement("Properties", customPropertiesSchema)
                rootNode.Attributes.Append(customPropsDoc.CreateAttribute("xmlns:vt"))
                rootNode.Attributes("xmlns:vt").Value = customVTypesSchema

                customPropsDoc.AppendChild(rootNode)

                ' Create the document's relationship to the new custom properties part:
                wdPackage.CreateRelationship(customPropsUri, TargetMode.Internal, customPropertiesRelationshipType)
            Else
                ' Load the contents of the custom properties part into an XML document.
                customPropsDoc = New XmlDocument(nt)
                customPropsDoc.Load(customPropsPart.GetStream())

                rootNode = customPropsDoc.DocumentElement
            End If

            ' Now that you've got a reference to an XmlDocument object that 
            ' corresponds to the custom properties part, 
            ' check to see if the required property is already there:
            Dim searchString As String = String.Format("d:Properties/d:property[translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='{0}']", propertyName.ToLower())
            'Dim searchString As String = String.Format("d:Properties/d:property[@name='{0}']", propertyName)
            Dim node As XmlNode = customPropsDoc.SelectSingleNode(searchString, nsManager)

            ' If you didn't find the node, add it. If you found it, and the type is 
            ' different than the new property, delete it. If you found it, and the type 
            ' is the same, replace the property value. Otherwise, add a new 
            ' element with the new value and type.

            Dim valueNode As XmlNode = Nothing
            If node IsNot Nothing Then
                node.ParentNode.RemoveChild(node)
                node = Nothing
                retVal = True
            End If

            ' Save the properties XML back to its part.
            customPropsDoc.Save(customPropsPart.GetStream(FileMode.Create, FileAccess.Write))
        End Using
        Return retVal
    End Function

    Public Function ReadAllProperties(ByVal fileName As String) As String Implements IFilePropIO.ReadAllProperties
        log.InfoFormat("ReadAllProperties(fileName={0}", fileName)
        Dim fi As New FileInfo(fileName)
        If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
        fi = Nothing

        Dim props As New DocumentProperties()

        Const corePropertiesRelationshipType As String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        Const corePropertiesSchema As String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        Const dcPropertiesSchema As String = "http://purl.org/dc/elements/1.1/"
        Const dcTermsPropertiesSchema As String = "http://purl.org/dc/terms/"
        Const customPropertiesRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        Const customPropertiesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Const customVTypesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        Dim corePropertiesPart As PackagePart = Nothing
        Dim customPropertiesPart As PackagePart = Nothing

        Using wdPackage As Package = Package.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
            ' Get the main document part (document.xml).
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(corePropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                corePropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one core properties part.
                Exit For
            Next
            If corePropertiesPart IsNot Nothing Then

                ' Manage namespaces to perform Xml XPath queries.
                Dim nt As New NameTable()
                Dim nsManager As New XmlNamespaceManager(nt)
                nsManager.AddNamespace("cp", corePropertiesSchema)
                nsManager.AddNamespace("dc", dcPropertiesSchema)
                nsManager.AddNamespace("dcterms", dcTermsPropertiesSchema)

                ' Get the properties from the package.
                Dim xdoc As XmlDocument = New XmlDocument(nt)

                ' Load the XML in the part into an XmlDocument instance:
                xdoc.Load(corePropertiesPart.GetStream())

                Dim searchString As String = "//cp:coreProperties"
                Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManager)
                If Not xNode Is Nothing Then
                    For Each pNode As XmlNode In xNode.ChildNodes
                        Select Case pNode.LocalName.ToLower
                            Case "subject", "category", "description", "title", "keywords", "creator"
                                props.Add(pNode.LocalName, pNode.InnerText)
                        End Select
                    Next
                End If
            End If


            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(customPropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                customPropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one custom properties part.
                Exit For
            Next
            ' There may not be a custom properties part.
            If customPropertiesPart IsNot Nothing Then
                ' Manage namespaces to perform Xml XPath queries.
                Dim nt As New NameTable()
                Dim nsManager As New XmlNamespaceManager(nt)
                nsManager.AddNamespace("d", customPropertiesSchema)
                nsManager.AddNamespace("vt", customVTypesSchema)

                ' Get the properties from the package.
                Dim xdoc As XmlDocument = New XmlDocument(nt)

                ' Load the XML in the part into an XmlDocument instance:
                xdoc.Load(customPropertiesPart.GetStream())

                Dim searchString As String = "d:Properties"
                Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManager)
                If Not xNode Is Nothing Then
                    For Each pNode As XmlNode In xNode.ChildNodes
                        If pNode.Name.StartsWith("property") _
                           AndAlso pNode.Attributes("name") IsNot Nothing _
                           AndAlso pNode.HasChildNodes Then
                            Dim pName As String = pNode.Attributes("name").InnerText
                            Dim pValue As String = pNode.FirstChild.InnerText
                            If props(pName) Is Nothing Then props.Add(pName, pValue)
                        End If
                    Next
                End If
            End If
        End Using
        Dim rs As String = props.Sort().ToXml()
        log.InfoFormat("Returning sorted XML result: {0}...", rs.Substring(0, 40))
        props.Clear()
        props = Nothing
        Return rs
    End Function

    Public Function WriteAllProperties(ByVal inputXML As String, ByVal fileName As String) As Long Implements IFilePropIO.WriteAllProperties
        log.InfoFormat("WriteAllPropertis(inputXLM=<xml>,fileName={0}", fileName)
        Dim irc As Integer = 0
        Dim rc As Integer = 0
        Dim props As DocumentProperties = DocumentProperties.FromXml(inputXML)


        Const corePropertiesSchema As String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        Const dcPropertiesSchema As String = "http://purl.org/dc/elements/1.1/"
        Const corePropertiesRelationshipType As String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        Const customPropertiesRelationshipType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        Const customPropertiesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        Const customVTypesSchema As String = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        Dim retVal As Boolean = False
        Dim corePropertiesPart As PackagePart = Nothing
        Dim customPropsPart As PackagePart = Nothing ' Work with the custom properties part


        Using wdPackage As Package = Package.Open(fileName, FileMode.Open, FileAccess.ReadWrite)

            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(corePropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                corePropertiesPart = wdPackage.GetPart(documentUri)
                ' There is only one core properties part.
                Exit For
            Next

            ' Get the custom part (custom.xml). It may not exist.
            For Each relationship As PackageRelationship In wdPackage.GetRelationshipsByType(customPropertiesRelationshipType)
                Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), relationship.TargetUri)
                customPropsPart = wdPackage.GetPart(documentUri)
                ' There is only one custom properties part, if it exists at all.
                Exit For
            Next


            ' Manage namespaces to perform Xml XPath queries.
            Dim ntCore As New NameTable()
            Dim nsManagerCore As New XmlNamespaceManager(ntCore)
            nsManagerCore.AddNamespace("cp", corePropertiesSchema)
            nsManagerCore.AddNamespace("dc", dcPropertiesSchema)
            ' Get the properties from the package.
            Dim xdoc As XmlDocument = New XmlDocument(ntCore)
            ' Load the XML in the part into an XmlDocument instance:
            xdoc.Load(corePropertiesPart.GetStream())

            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("d", customPropertiesSchema)
            nsManager.AddNamespace("vt", customVTypesSchema)
            Dim customPropsUri As New Uri("/docProps/custom.xml", UriKind.Relative)
            Dim customPropsDoc As XmlDocument = Nothing
            Dim rootNode As XmlNode = Nothing

            ' There may not be a custom properties part.
            If customPropsPart Is Nothing Then
                customPropsDoc = New XmlDocument(nt)
                ' Part doesn't exist. Create it now.
                customPropsPart = wdPackage.CreatePart(customPropsUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml")
                ' Set up the rudimentary custom part:
                rootNode = customPropsDoc.CreateElement("Properties", customPropertiesSchema)
                rootNode.Attributes.Append(customPropsDoc.CreateAttribute("xmlns:vt"))
                rootNode.Attributes("xmlns:vt").Value = customVTypesSchema
                customPropsDoc.AppendChild(rootNode)
                ' Create the document's relationship to the new custom properties part:
                wdPackage.CreateRelationship(customPropsUri, TargetMode.Internal, customPropertiesRelationshipType)
            Else
                ' Load the contents of the custom properties part into an XML document.
                customPropsDoc = New XmlDocument(nt)
                customPropsDoc.Load(customPropsPart.GetStream())
                rootNode = customPropsDoc.DocumentElement
            End If

            For Each prop As DocumentProperty In props
                Dim propertyName As String = prop.Name
                Dim propertyValue As Object = prop.Value
                Dim propertyTypeName As String = GetPropertyTypeName(PropertyTypes.Text)

                'First, try to find a core property with that name
                Dim searchString As String = String.Format("//cp:coreProperties/dc:{0}|//cp:coreProperties/cp:{0}", propertyName.ToLower, propertyName.ToLower)
                Dim xNode As XmlNode = xdoc.SelectSingleNode(searchString, nsManagerCore)
                If xNode IsNot Nothing Then
                    If propertyValue.Equals("#!del!#") Then
                        xNode.InnerText = String.Empty
                    Else
                        xNode.InnerText = propertyValue
                    End If
                    ' Now update the value:
                    Continue For
                End If

                ' Now that you've got a reference to an XmlDocument object that 
                ' corresponds to the custom properties part, 
                ' check to see if the required property is already there:
                searchString = String.Format("d:Properties/d:property[translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='{0}']", propertyName.ToLower())
                xNode = customPropsDoc.SelectSingleNode(searchString, nsManager)

                ' If you didn't find the node, add it. If you found it, and the type is 
                ' different than the new property, delete it. If you found it, and the type 
                ' is the same, replace the property value. Otherwise, add a new 
                ' element with the new value and type.
                Dim valueNode As XmlNode = Nothing
                If xNode IsNot Nothing Then
                    ' You found the node. Now check its type:
                    If xNode.HasChildNodes Then
                        valueNode = xNode.ChildNodes(0)
                        If valueNode IsNot Nothing Then
                            Dim typeName As String = valueNode.Name
                            If propertyTypeName = typeName AndAlso Not propertyValue.Equals("#!del!#") Then
                                ' The types are the same. Simply 
                                ' replace the value of the node:
                                valueNode.InnerText = propertyValue
                            Else
                                ' Types are different. Delete the node, 
                                ' and clear the node variable:
                                xNode.ParentNode.RemoveChild(xNode)
                                xNode = Nothing
                            End If
                        End If
                    End If
                End If

                ' The previous block of code may have cleared the value in the variable named node:
                If xNode Is Nothing AndAlso Not propertyValue.Equals("#!del!#") Then
                    ' Either you didn't find the node, or you 
                    ' found it, its type was incorrect, and you deleted it.
                    ' Either way, you need to create the new property node now.
                    ' Find the highest existing "pid" value.
                    ' The default value for the "pid" attribute is "2".
                    Dim pidValue As String = "2"
                    Dim propertiesNode As XmlNode = customPropsDoc.DocumentElement
                    If propertiesNode.HasChildNodes Then
                        Dim lastNode As XmlNode = propertiesNode.LastChild
                        If lastNode IsNot Nothing Then
                            Dim pidAttr As XmlAttribute = lastNode.Attributes("pid")
                            If Not pidAttr Is Nothing Then
                                pidValue = pidAttr.Value
                                ' Increment pidValue, so that the new property
                                ' gets a pid value one higher. This value should be 
                                ' numeric, but it never hurt so to confirm:
                                Dim value As Integer
                                If Integer.TryParse(pidValue, value) Then
                                    pidValue = Convert.ToString(value + 1)
                                End If
                            End If
                        End If
                    End If
                    xNode = customPropsDoc.CreateElement("property", customPropertiesSchema)
                    xNode.Attributes.Append(customPropsDoc.CreateAttribute("name"))
                    xNode.Attributes("name").Value = propertyName
                    xNode.Attributes.Append(customPropsDoc.CreateAttribute("fmtid"))
                    xNode.Attributes("fmtid").Value = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                    xNode.Attributes.Append(customPropsDoc.CreateAttribute("pid"))
                    xNode.Attributes("pid").Value = pidValue
                    valueNode = customPropsDoc.CreateElement(propertyTypeName, customVTypesSchema)
                    valueNode.InnerText = propertyValue
                    xNode.AppendChild(valueNode)
                    rootNode.AppendChild(xNode)
                End If
            Next



            ' Save the properties XML back to its part.
            Dim settings As New XmlWriterSettings()
            settings.Indent = False
            'settings.ConformanceLevel = ConformanceLevel.Document
            settings.Encoding = New System.Text.UTF8Encoding()
            settings.NewLineHandling = NewLineHandling.Entitize

            Dim writer As XmlWriter = XmlWriter.Create(customPropsPart.GetStream(FileMode.Create, FileAccess.Write), settings)
            customPropsDoc.Save(writer) 'customPropsPart.GetStream(FileMode.Create, FileAccess.Write))

            xdoc.Save(corePropertiesPart.GetStream(FileMode.Create, FileAccess.Write))

            wdPackage.Close()
        End Using

        Return rc
    End Function

    ''' <summary>
    ''' Write property to Word document or Excel workbook
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
        Dim rc As Integer = 0

        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing


            Select Case propertyName.ToLower()
                Case "author" : WDSetCoreProperty(fileName, "creator", propertyValue)
                Case "comments" : WDSetCoreProperty(fileName, "description", propertyValue)
                Case "creator", "subject", "category", "description", "title", "keywords" : WDSetCoreProperty(fileName, propertyName, propertyValue)
                Case Else : WDSetCustomProperty(fileName, propertyName, propertyValue, PropertyTypes.Text)
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
    ''' <para>Default properties get set to an empty string, custom properties are deleted
    ''' completely.</para>
    ''' </remarks>
    ''' <param name="fileName">Full path to office document</param>
    ''' <param name="propertyName">Name of the property to write to</param>
    ''' <returns>
    ''' 0 if succesful, 
    ''' <para><see cref="T:MT_PropertyExchange.Worker.Errors">Error Code</see>
    ''' otherwise</para>
    ''' </returns>
    Public Function DeleteProp(ByVal fileName As String, _
                           ByVal propertyName As String) As Long Implements IFilePropIO.DeleteProp
        log.InfoFormat("DeleteProp(fileName={0},propertyName={1}", fileName, propertyName)
        Dim rc As Integer = Errors.NoError
        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            If fi.IsReadOnly Then : Throw New FileReadOnlyException("File " & fileName & " is read only.") : End If
            fi = Nothing

            WriteProp(fileName, propertyName, String.Empty)
            WDDeleteCustomProperty(fileName, propertyName)

        Catch ex As Exception
            rc = HandleEx(ex)
        End Try
        log.InfoFormat("Return code: {0}", rc)
        Return rc
    End Function

    ''' <summary>
    ''' <para>Read property from Word document or Excel workbook</para>
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
        Dim rc As Integer = 0
        Dim rv As String = String.Empty
        Try
            Dim fi As New FileInfo(fileName)
            If Not fi.Exists Then : Throw New FileNotFoundException("File " & fileName & " not found.") : End If
            fi = Nothing

            Dim sb As New StringBuilder(String.Empty)
            Dim sl As New SortedList(Of String, String)
            If propertyName.Equals("*") Then
                Dim p As DocumentProperties = DocumentProperties.FromXml(ReadAllProperties(fileName))
                For Each pr As DocumentProperty In p.Sort()
                    sb.AppendFormat("{0,-20}: {1}", pr.Name, pr.Value).AppendLine()
                Next
                rv = sb.ToString()
            Else
                rv = WDRetrieveCoreProperty(fileName, propertyName)
                If String.IsNullOrEmpty(rv) Then rv = WDRetrieveCustomProperty(fileName, propertyName)
            End If

        Catch invDataEx As InvalidDataException
            System.Console.Error.WriteLine(invDataEx.Message)
            rc = Errors.InvalidData
        Catch ex As Exception
            rc = HandleEx(ex)
        End Try
        Return rv
    End Function

    ''' <summary>
    ''' Execute a macro on an office document
    ''' </summary>
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
        Return Globals.RunMacro(fileName, macroName, template, Me, visible)
    End Function


End Class

