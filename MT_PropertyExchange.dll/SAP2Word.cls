'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "SAP2Word"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
'Option Explicit
' References
' 1) DSO File (DS: OLE Document Property)
' 2) Microsoft Word
' 3) Microsoft Excel

'input
' Filename          filename including complete path
' PropertyName      name of property to set
' IsCustomProperty  true if property to set is a custom property, false for built in properties
' PropertyValue     new value for property
'return codes
' 0              success
' -2147220987    custom property not found (custom properties, that do not exist are not added)
' 9001           built in property not found (custom error code)
' 432            file not found
' 9002           wrong datatype (custom error code)
' 9003           numeric values may not have decimals (custom error code)
' 9004           custom property not found -> not used any more ( only with DeleteProp)
Public Function WriteProp(ByVal FileName As String, _
                          ByVal PropertyName As String, _
                          ByVal IsCustomProperty As Boolean, _
                          ByVal PropertyValue As String) As Long


    On Error GoTo Err_WriteProp

    Dim oFilePropReader As DSOleFile.PropertyReader
    Dim iProp As DSOleFile.DocumentProperties
    Dim oNewProp As DSOleFile.CustomProperty
    Dim iResult As Long
    Dim bFound As Boolean

    iResult = 0
    PropertyName = Trim(PropertyName)

    oFilePropReader = New DSOleFile.PropertyReader
    iProp = oFilePropReader.GetDocumentProperties(FileName)

    '  If Not IsCustomProperty Then
    'built in properties
    Select Case LCase(PropertyName)
        Case "author" : iProp.Author = PropertyValue
        Case "subject" : iProp.Subject = PropertyValue
        Case "category" : iProp.Category = PropertyValue
        Case "comments" : iProp.Comments = PropertyValue
        Case "title" : iProp.Title = PropertyValue
        Case "manager" : iProp.Manager = PropertyValue
        Case "company" : iProp.Company = PropertyValue
        Case "keywords" : iProp.Keywords = PropertyValue
        Case Else

            'customproperties
            bFound = False
            For Each oNewProp In iProp.CustomProperties
                If LCase(oNewProp.Name) = LCase(PropertyName) Then
                    bFound = True
                    Select Case oNewProp.Type
                        Case 1 ' text
                            oNewProp.Value = Trim(PropertyValue)
                        Case 2 ' number
                            If IsNumeric(PropertyValue) And CLng(PropertyValue) = CDec(PropertyValue) Then
                                oNewProp.Value = CLng(PropertyValue)
                            Else
                                If CLng(PropertyValue) = CDec(PropertyValue) Then
                                    iResult = 9003
                                Else
                                    iResult = 9002
                                End If
                            End If
                        Case 5 ' date
                            If IsDate(PropertyValue) Then
                                oNewProp.Value = CDate(PropertyValue)
                            Else
                                iResult = 9002
                            End If
                        Case 4 ' boolean
                            If LCase(PropertyValue) = "true" Or LCase(PropertyValue) = "false" Then
                                oNewProp.Value = CBool(PropertyValue)
                            Else
                                iResult = 9002
                            End If
                    End Select
                End If
            Next
            If Not bFound Then
                'not found -> add (always as string typed property)
                iProp.CustomProperties.Add(PropertyName, Trim(CStr(PropertyValue)))
            End If
    End Select



    oFilePropReader = Nothing
    iProp = Nothing

    WriteProp = iResult

    Exit Function

Err_WriteProp:

    oFilePropReader = Nothing
    iProp = Nothing
    iResult = Err.Number
    WriteProp = iResult

End Function


Public Function DeleteProp(ByVal FileName As String, _
                          ByVal PropertyName As String) As Long

    ' Error Codes
    ' 0   ....... Success
    ' 9004 ...... Custom Document Property not found
    On Error GoTo Err_DeleteProp

    Dim oFilePropReader As DSOleFile.PropertyReader
    Dim iProp As DSOleFile.DocumentProperties
    Dim oNewProp As DSOleFile.CustomProperty
    Dim iResult As Long
    Dim bFound As Boolean

    iResult = 0
    PropertyName = Trim(PropertyName)

    oFilePropReader = New DSOleFile.PropertyReader
    iProp = oFilePropReader.GetDocumentProperties(FileName)
    ' Custom Document Properties are deleted
    ' Built In Document Properties are set to ""
    Select Case LCase(PropertyName)
        Case "author" : iProp.Author = ""
        Case "subject" : iProp.Subject = ""
        Case "category" : iProp.Category = ""
        Case "comments" : iProp.Comments = ""
        Case "title" : iProp.Title = ""
        Case "manager" : iProp.Manager = ""
        Case "company" : iProp.Company = ""
        Case "keywords" : iProp.Keywords = ""
        Case Else : iResult = 9001
            bFound = False
            For Each oNewProp In iProp.CustomProperties
                If LCase(oNewProp.Name) = LCase(PropertyName) Then
                    bFound = True
                    oNewProp.Remove()
                    iResult = 0
                End If
            Next
            If Not bFound Then
                iResult = 9004
            End If
    End Select

    oFilePropReader = Nothing
    iProp = Nothing

    DeleteProp = iResult

    Exit Function

Err_DeleteProp:

    oFilePropReader = Nothing
    iProp = Nothing
    iResult = Err.Number
    DeleteProp = iResult

End Function

Public Function existsProperty(ByVal FileName As String, ByVal PropertyName As String) As Boolean
    Dim oFilePropReader As DSOleFile.PropertyReader
    Dim iProp As DSOleFile.DocumentProperties
    Dim oNewProp As DSOleFile.CustomProperty
    Dim bFound As Boolean

    PropertyName = Trim(PropertyName)

    oFilePropReader = New DSOleFile.PropertyReader
    iProp = oFilePropReader.GetDocumentProperties(FileName)

    bFound = False
    For Each oNewProp In iProp.CustomProperties
        If LCase(oNewProp.Name) = LCase(PropertyName) Then
            bFound = True
        End If
    Next
    existsProperty = bFound
End Function

Public Function ReadProperty(ByVal FileName As String, ByVal PropertyName As String)
    Dim oFilePropReader As DSOleFile.PropertyReader
    Dim iProp As DSOleFile.DocumentProperties
    Dim oNewProp As DSOleFile.CustomProperty
    Dim strHelp As String

    PropertyName = Trim(PropertyName)
    strHelp = ""
    oFilePropReader = New DSOleFile.PropertyReader
    iProp = oFilePropReader.GetDocumentProperties(FileName)

    For Each oNewProp In iProp.CustomProperties
        If LCase(oNewProp.Name) = LCase(PropertyName) Then
            strHelp = oNewProp.Value
        End If
    Next
    ReadProperty = strHelp
End Function

Public Function deleteOldProperties(ByVal FileName As String, ByVal MapOld2New As Boolean) As Boolean
    Dim iResult
    Dim arrayProperties(1 To 38, 1 To 2)
    Dim i
    Dim strHelp
    arrayProperties(1, 1) = "vaiAreaCode"
    arrayProperties(1, 2) = "Area"
    arrayProperties(2, 1) = "vaiSeqNo"
    arrayProperties(2, 2) = "CountNo"
    arrayProperties(3, 1) = "vaiCustomerDocNo"
    arrayProperties(3, 2) = "CustomerDrawingNo"
    arrayProperties(4, 1) = "vaiDepartment"
    arrayProperties(4, 2) = "projectDepartment"
    arrayProperties(5, 1) = "vaiSCSNumber"
    arrayProperties(5, 2) = "DocumentNo"
    arrayProperties(6, 1) = "vaiDocTypeCode"
    arrayProperties(6, 2) = "DocType"
    arrayProperties(7, 1) = "vaiDocType"
    arrayProperties(7, 2) = "DocumentTypeText"
    arrayProperties(8, 1) = "vaiEngPackage"
    arrayProperties(8, 2) = "EngineeringPackage"
    arrayProperties(9, 1) = "vaiEngPackTxt"
    arrayProperties(9, 2) = "EngPackText"
    arrayProperties(10, 1) = "vaiEquipCode"
    arrayProperties(10, 2) = "Equipment"
    arrayProperties(11, 1) = "vaiPartNo"
    arrayProperties(11, 2) = "PartNo"
    arrayProperties(12, 1) = "vaiPlantCode"
    arrayProperties(12, 2) = "PlantCode"
    arrayProperties(13, 1) = "vaiProjectCode"
    arrayProperties(13, 2) = "ProjectCode"
    arrayProperties(14, 1) = "vaiProjectProperty1"
    arrayProperties(14, 2) = "ProjectProperty1"
    arrayProperties(15, 1) = "vaiProjectProperty2"
    arrayProperties(15, 2) = "ProjectProperty2"
    arrayProperties(16, 1) = "vaiProjectProperty3"
    arrayProperties(16, 2) = "ProjectProperty3"
    arrayProperties(17, 1) = "vaiProjectProperty4"
    arrayProperties(17, 2) = "ProjectProperty4"
    arrayProperties(18, 1) = "vaiProjectProperty5"
    arrayProperties(18, 2) = "ProjectProperty5"
    arrayProperties(19, 1) = "vaiProjectProperty6"
    arrayProperties(19, 2) = "ProjectProperty6"
    arrayProperties(20, 1) = "vaiProjectProperty7"
    arrayProperties(20, 2) = "ProjectProperty7"
    arrayProperties(21, 1) = "vaiProjectProperty8"
    arrayProperties(21, 2) = "ProjectProperty8"
    arrayProperties(22, 1) = "vaiProjectProperty9"
    arrayProperties(22, 2) = "ProjectProperty9"
    arrayProperties(23, 1) = "vaiRevision"
    arrayProperties(23, 2) = "RevisionNo"
    arrayProperties(24, 1) = "vaiSecLevelTxt"
    arrayProperties(24, 2) = "SecLevText"
    arrayProperties(25, 1) = "vaiSecLevel"
    arrayProperties(25, 2) = "SecurityLevel"
    arrayProperties(26, 1) = "vaiStatus"
    arrayProperties(26, 2) = "Status"
    arrayProperties(27, 1) = "vaiSubGroup"
    arrayProperties(27, 2) = "SubGroup"
    arrayProperties(28, 1) = "vaiSupplier"
    arrayProperties(28, 2) = "Supplier"
    arrayProperties(29, 1) = "vaiSupplierDocNo"
    arrayProperties(29, 2) = "SupplierDocNo"
    arrayProperties(30, 1) = "vaiSGCCode"
    arrayProperties(30, 2) = "SupplyGroup"
    arrayProperties(31, 1) = "vaiSGC"
    arrayProperties(31, 2) = "SupplyGroupText"
    arrayProperties(32, 1) = "vaiTitle1"
    arrayProperties(32, 2) = "Title1"
    arrayProperties(33, 1) = "vaiTitle2"
    arrayProperties(33, 2) = "Title2"
    arrayProperties(34, 1) = "vaiTitle3"
    arrayProperties(34, 2) = "Title3"
    arrayProperties(35, 1) = "vaiTitle4"
    arrayProperties(35, 2) = "Title4"
    arrayProperties(36, 1) = "vaiDescription"
    arrayProperties(36, 2) = "Title5"
    arrayProperties(37, 1) = "vaiAsBuilt"
    arrayProperties(37, 2) = "AsBuilt"
    arrayProperties(38, 1) = "vaiCorrespNo"
    arrayProperties(38, 2) = "CorrespNo"

    If MapOld2New Then
        For i = 1 To 38
            If existsProperty(FileName, arrayProperties(i, 1)) Then
                strHelp = ReadProperty(FileName, arrayProperties(i, 1))
                WriteProp(FileName, arrayProperties(i, 2), False, strHelp)
            End If
        Next
    End If
    For i = 1 To 38
        DeleteProp(FileName, arrayProperties(i, 1))
    Next


    ' Delete all old properties


End Function


'input
' filename       filename including complete path
' macroname      macroname PROJECTNAME.MODULENAME.MACRONAME
' Excel and Word has to be linked to DLL
'return codes
' 0              success
' 5174           file not found
' -2147352573    macro not found
' 1              write protected
' 5941           application error
' Other          Word or Excel errors
' 9100           error Writing EXCEL Properties
' 9101           error Writing EXCEL BuiltInProperties
' 9102           error Writing EXCEL CustomProperties
' 9200           error Writing EXCEL Document
' 9201           error Writing EXCEL Document from BuiltInProperties
' 9201           error Writing EXCEL Document from CustomProperties
Public Function RunMacro(ByVal FileName As String, ByVal MacroName As String) As Long

    'Dim iResult As Long
    Dim iResult As Long
    Dim sExtension As String
    Dim oWord As Word.Application
    Dim oExcel As Excel.Application
    Dim wbXL As Excel.Workbook
    'Dim oDoc As Word.Document
    Dim bStarted As Boolean
    Dim sHelp As String

    bStarted = True

    iResult = 0
    MacroName = Trim(MacroName)
    On Error Resume Next
    sHelp = UCase(Left(MacroName, Len("DLLSAP2Document")))

    If (sHelp = UCase("DLLSAP2Document")) Then
        ' Word conversion macros
        ' we have to check wheather the old properties still exist
        ' vaiTitle1 was present at all old project templates
        ' therefore within those files the old properties have to be deleted.
        ' Property Harmonisation

        If existsProperty(FileName, "vaiTitle1") Then
            deleteOldProperties(FileName, False)
        End If
    End If
    If (sHelp = UCase("DLLDocument2SAP")) Then
        ' Word conversion macros
        ' we have to check wheather the old properties still exist
        ' vaiTitle1 was present at all old project templates
        ' therefore within those files the old properties have to be deleted.
        ' Property Harmonisation
        ' Batchimport
        If existsProperty(FileName, "vaiTitle1") Then
            deleteOldProperties(FileName, True)
        End If
    End If

    sExtension = UCase(Right(FileName, 3))
    ' Creating Documents out of a xml file or xml files of a directory
    ' This only works with Word
    ' therefore the extension is not relevant
    If (UCase(MacroName) = UCase("DLLCreateDocument")) Or (UCase(MacroName) = UCase("DLLCreateDocuments")) Then
        sExtension = "DOC"
    End If

    Select Case sExtension
        Case "DOC"
            On Error GoTo Err_RunMacro
            oWord = CreateObject("Word.Application")
            oWord.DisplayAlerts = wdAlertsNone
            oWord.NormalTemplate.Saved = True
            iResult = oWord.Run(MacroName, FileName)
            oWord.NormalTemplate.Saved = True
            oWord.DisplayAlerts = wdAlertsAll
            If bStarted Then
                oWord.Quit()
            End If
            oWord = Nothing
        Case "XLS"
            On Error GoTo Err_RunMacro

            oExcel = CreateObject("Excel.Application")
            oExcel.DisplayAlerts = False

            ' Addin has to be loaded first of all
            wbXL = oExcel.Workbooks.Open(oExcel.LibraryPath & "\vai\plmupdateServer.XLA")
            'wbXL.RunAutoMacros xlAutoDeactivate
            iResult = oExcel.Run(MacroName, FileName)
            wbXL.Close(False)
            wbXL = Nothing

            oExcel.DisplayAlerts = True
            If bStarted Then
                oExcel.Quit()
            End If
            oExcel = Nothing
    End Select

    RunMacro = iResult

    Exit Function

Err_RunMacro:

    Select Case sExtension
        Case "DOC"

            If Not oWord Is Nothing Then
                oWord.DisplayAlerts = wdAlertsAll
            End If

            If bStarted Then
                If Not oWord Is Nothing Then
                    oWord.Quit()
                End If
            End If
            oWord = Nothing
        Case "XLS"
            If Not oExcel Is Nothing Then
                oExcel.DisplayAlerts = True
            End If

            If bStarted Then
                If Not oExcel Is Nothing Then
                    oExcel.Quit()
                End If
            End If
            oExcel = Nothing
    End Select
    iResult = Err.Number
    RunMacro = iResult

End Function



