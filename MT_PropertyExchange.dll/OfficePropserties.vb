Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Runtime.Serialization.Formatters.Binary



<Serializable()> _
Public Class OfficeProperties
    Implements ICloneable
    Public Sub New()
        projektBearbeiter = New Bearbeiter()

        eingabeSpannung = New ListValue(3, "Drehstrom", 400)
        eingabeLeiter = New ListValue(2, "Kupfer", 57)
        eingabeIsolierung = New AddValue("PVC-Isolierte Kabel", 1)
        eingabeKennlinie = New AddValue("LSS B", 1)

        eingabeStromkreis = New AddValue("Endstromkreis", 1)
        eingabeVerlegeart = New AddValue("in Gebäuden (in/auf Wand)", 1)
        eingabeVerlegung = New AddValue("A - Aderleitungen in Wärmedämmung", 1)
        eingabeVerlegeAnordnung = New AddValue("gebündelt", 1)
        eingabeAnordnung = New AddValue("in Rohr/Kanal in Wand", 2)
        eingabeParaKabel = New Double(4) {0, 0, 0, 0, 0}
    End Sub

    Private m_leistung As Double = 0
    Private m_bStrom As Double = 0
    Public showMessage As Boolean = False
    Public message As String = [String].Empty

#Region "Projektangaben"
    Public projektId As Long = -1
    Public projektBezeichnung As String = [String].Empty
    Public projektBearbeiter As Bearbeiter
    Public projektNormId As Long = 2009
    Public projektBerechnungsArt As String = CalculationType.Kabel.ToString()
#End Region

#Region "Einstellungen"
    Public eingabeSpannung As ListValue
    Public eingabeLeiter As ListValue
    Public eingabeIsolierung As AddValue
    Public eingabeSpannAbf As Double = 1.5
#End Region


    Private Shared Function checkVersion(ByVal fileVersion As String) As Boolean
        Dim rc As Boolean = True
        'Projekt obj = new Projekt();
        Dim ver As Version = System.Reflection.Assembly.GetEntryAssembly().GetName().Version
        ' (obj.GetType()).GetName().Version;
        Dim parts As String() = fileVersion.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
        If Int32.Parse(parts(0).ToString()) = ver.Major Then
            If Int32.Parse(parts(1).ToString()) > ver.Minor Then
                rc = False
            End If
        Else
            rc = False
        End If
        Return rc
    End Function

#Region "Clone function"
    Public Function Clone() As Object Implements ICloneable.Clone
        Dim ms As New MemoryStream()
        Dim bf As New BinaryFormatter()
        bf.Serialize(ms, Me)
        ms.Position = 0
        Dim obj As Object = bf.Deserialize(ms)
        ms.Close()
        Return obj
    End Function
#End Region

#Region "Override of Equals function"
    Private Shared inEquals As Boolean = False
    Public Overrides Function Equals(ByVal obj As System.Object) As Boolean
        If inEquals Then
            Return (obj Is Me)
        End If
        ' If parameter is null return false.
        If obj Is Nothing Then
            Return False
        End If

        ' If parameter cannot be cast to Point return false.
        Dim p As Projekt = TryCast(obj, Projekt)
        If DirectCast(p, System.Object) Is Nothing Then
            Return False
        End If

        ' Return true if the fields match:
        Return Equals(p)
    End Function
    Public Overloads Function Equals(ByVal p As Projekt) As Boolean
        If inEquals Then
            Return True
        End If
        If DirectCast(p, Object) Is Nothing Then
            Return False
        End If

        inEquals = True
        'XmlSerializer mySerializer = new XmlSerializer(typeof(Projekt));
        'StringWriter xml1 = new StringWriter();
        'StringWriter xml2 = new StringWriter();
        'mySerializer.Serialize(xml1, this);
        'mySerializer.Serialize(xml2, p);

        Dim rc As Boolean = Me.ToXml().Equals(p.ToXml())
        inEquals = False
        Return rc
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
#End Region

#Region "Serialization"
    Public Function ToXml() As String
        Dim mySerializer As New XmlSerializer(GetType(Projekt))
        Dim memStream As New MemoryStream()
        Dim myWriter As New StreamWriter(memStream)
        Dim utf8e As New UTF8Encoding()
        Try
            inEquals = True
            mySerializer.Serialize(myWriter, Me)
            inEquals = False
            myWriter.Close()
        Catch e As Exception
            MessageBox.Show("Fehler beim Prüfen des Projekts." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return [String].Empty
        End Try
        Return utf8e.GetString(memStream.ToArray())
    End Function
    Public Shared Function FromXml(ByVal xml As String) As Projekt
        Dim loaded As Projekt
        Dim mySerializer As New XmlSerializer(GetType(Projekt))
        Dim myReader As New StringReader(xml)
        Try
            'if (loaded.version.Equals(String.Empty)) loaded.SetActVersion();
            loaded = DirectCast(mySerializer.Deserialize(myReader), Projekt)
        Catch e As Exception
            MessageBox.Show("Fehler beim Lesen des Projekts." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            loaded = Nothing
        Finally
            myReader.Close()
        End Try
        Return loaded
    End Function
    Public Function SaveToFile(ByVal fileName As String) As Boolean
        Try
            Dim xml As String = ToXml()
            If [String].IsNullOrEmpty(xml) Then
                Return False
            End If
            File.WriteAllText(fileName, xml)
        Catch e As Exception
            MessageBox.Show("Fehler beim Speichern des Projekts." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return False
        End Try
        Return True
    End Function
    Public Shared Function LoadFromFile(ByVal fileName As String) As Projekt
        Dim xml As String
        Try
            xml = File.ReadAllText(fileName)
        Catch
            MessageBox.Show([String].Format("Fehler beim Laden der Datei '{0}'", fileName), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return New Projekt()
        End Try
        Return FromXml(xml)
    End Function
    Public Shared Function LoadFromZip(ByVal fileName As String) As Projekt
        Dim zip As ZipStorer = Nothing
        Dim tempDir As String = [String].Empty
        Dim tempXmlFile As String = [String].Empty
        Dim tempZipFile As String = [String].Empty
        Dim versionString As String = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3)
        Dim dir As List(Of ZipStorer.ZipFileEntry)
        Const fmtInvalidVer As String = "Inkompatible Dateiversion." & vbCr & vbLf & vbCr & vbLf & "Die Datei wurde mit der Version {0} erstellt und kann nicht mit der aktuellen Version {1} gelesen werden."

        Try
            Try
                tempDir = Path.GetTempFileName()
                File.Delete(tempDir)
                Directory.CreateDirectory(tempDir)
                tempXmlFile = Path.GetTempFileName()
                File.Delete(tempXmlFile)
                tempZipFile = Path.Combine(tempDir, Path.GetFileName(tempDir))
                tempXmlFile = Path.Combine(tempDir, Path.GetFileName(tempXmlFile))
                File.Copy(fileName, tempZipFile, True)
                zip = ZipStorer.Open(tempZipFile, FileAccess.ReadWrite)
                dir = zip.ReadCentralDir()
                If checkVersion(dir(0).ExtraField) Then
                    zip.ExtractStoredFile(dir(0), tempXmlFile)
                Else
                    MessageBox.Show(String.Format(fmtInvalidVer, dir(0).ExtraField.ToString(), versionString), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    Return Nothing
                End If
                Return LoadFromFile(tempXmlFile)
            Catch generatedExceptionName As System.Exception
                Throw New Exception(String.Format("Datei konnte nicht geöffnet werden."))

                'try
                '{
                '    return LoadFromFile(tempXmlFile);
                '}
                'catch (System.Exception)
                '{
                '    throw new Exception(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString));
                '}
            End Try
        Catch ex As System.Exception
            MessageBox.Show(ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
        Finally
            If zip IsNot Nothing Then
                zip.Close()
            End If
            Directory.Delete(tempDir, True)
        End Try
        Return Nothing
    End Function
    Public Function SaveAsZip(ByVal fileName As String) As Boolean
        If SaveToFile(fileName) Then
            Return ZipFile(fileName)
        Else
            Return False
        End If
    End Function
    Public Function ZipFile(ByVal fileName As String) As Boolean
        Dim zip As ZipStorer = Nothing
        Dim tempPath As String = [String].Empty
        Dim tempDir As String = [String].Empty
        Dim versionString As String = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3)


        Try
            tempDir = Path.GetTempFileName()
            File.Delete(tempDir)
            Directory.CreateDirectory(tempDir)
            tempPath = Path.Combine(tempDir, Path.GetFileName(fileName))
            zip = New ZipStorer()
            zip = ZipStorer.Create(tempPath, "", versionString)
            zip.AddFile(fileName, Path.GetFileName(fileName), "", versionString)
            zip.Close()
            File.Copy(tempPath, fileName, True)
            Return True
        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Datei" & vbLf & "'" & fileName.ToString() & "':" & ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return False
        Finally
            Try
                zip.Close()
            Catch
            End Try
            Directory.Delete(tempDir, True)
        End Try
    End Function
#End Region
End Class

#Region "Bearbeiter"
<Serializable()> _
Public Class Bearbeiter
    Public Sub New()
        Me.New(-1, [String].Empty)
    End Sub
    Public Sub New(ByVal id As Long, ByVal name As String)
        Me.id = id
        Me.name = name
    End Sub
    Public id As Long
    Public name As String

#Region "Clone function"
    Public Function Clone() As Object
        Dim ms As New MemoryStream()
        Dim bf As New BinaryFormatter()
        bf.Serialize(ms, Me)
        ms.Position = 0
        Dim obj As Object = bf.Deserialize(ms)
        ms.Close()
        Return obj
    End Function
#End Region

#Region "Override of Equals function"
    Private Shared inEquals As Boolean = False
    Public Overrides Function Equals(ByVal obj As System.Object) As Boolean
        If inEquals Then
            Return True
        End If
        ' If parameter is null return false.
        If obj Is Nothing Then
            Return False
        End If

        ' If parameter cannot be cast to Point return false.
        Dim p As Bearbeiter = TryCast(obj, Bearbeiter)
        If DirectCast(p, System.Object) Is Nothing Then
            Return False
        End If

        ' Return true if the fields match:
        Return Equals(p)
    End Function
    Public Overloads Function Equals(ByVal p As Bearbeiter) As Boolean
        If inEquals Then
            Return True
        End If
        If DirectCast(p, Object) Is Nothing Then
            Return False
        End If

        inEquals = True
        Dim mySerializer As New XmlSerializer(GetType(Bearbeiter))
        Dim xml1 As New StringWriter()
        Dim xml2 As New StringWriter()
        mySerializer.Serialize(xml1, Me)
        mySerializer.Serialize(xml2, p)

        inEquals = False
        Return (xml1.ToString().Equals(xml2.ToString()))
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
#End Region
End Class

<Serializable()> _
Public Class BearbeiterCollection
    Inherits CollectionBase
#Region "Constructors"
    Public Sub New()
    End Sub
#End Region

#Region "Implementation of CollectionBase"
    Public Function Add(ByVal value As Bearbeiter) As Integer
        Return Me.List.Add(value)
    End Function

    Public Sub AddRange(ByVal value As Bearbeiter())
        For i As Integer = 0 To value.Length - 1
            Me.Add(value(i))
        Next
    End Sub

    Public Sub AddRange(ByVal value As BearbeiterCollection)
        For i As Integer = 0 To value.Count - 1
            Me.Add(value(i))
        Next
    End Sub

    Public Function Contains(ByVal value As Bearbeiter) As Boolean
        Return Me.List.Contains(value)
    End Function

    Public Sub CopyTo(ByVal layArray As Bearbeiter(), ByVal intIndex As Integer)
        Me.List.CopyTo(layArray, intIndex)
    End Sub

    'public new BearbeiterEnumerator GetEnumerator()
    '{
    '    return new BearbeiterEnumerator(this);
    '}

    Public Function IndexOf(ByVal value As Bearbeiter) As Integer
        Return Me.List.IndexOf(value)
    End Function

    Public Function IndexOf(ByVal id As Long) As Integer
        Dim num3 As Integer = Me.List.Count - 1
        For i As Integer = 0 To num3
            If DirectCast(Me.List(i), Bearbeiter).id.Equals(id) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Function IndexOf(ByVal bezeichnung As String) As Integer
        For i As Integer = 0 To Me.List.Count - 1
            If DirectCast(Me.List(i), Bearbeiter).name.Equals(bezeichnung.Trim()) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Sub Insert(ByVal intIndex As Integer, ByVal value As Bearbeiter)
        Me.List.Insert(intIndex, value)
    End Sub

    Protected Overrides Sub OnClear()
        While Me.List.Count > 0
            Dim value As Bearbeiter = DirectCast(Me.List(0), Bearbeiter)
            Me.Remove(value)
            value = Nothing
        End While
    End Sub

    Protected Overrides Sub OnInsert(ByVal intIndex As Integer, ByVal objValue As Object)
    End Sub

    Protected Overrides Sub OnRemove(ByVal intIndex As Integer, ByVal objValue As Object)
    End Sub

    Protected Overrides Sub OnSet(ByVal intIndex As Integer, ByVal objOldValue As Object, ByVal objNewValue As Object)
    End Sub

    Protected Overrides Sub OnValidate(ByVal objValue As Object)
    End Sub

    Public Sub Remove(ByVal value As Bearbeiter)
        Me.List.Remove(value)
    End Sub
#End Region

    Public Function GetSorted() As SortedList(Of String, Long)
        Dim slist As New SortedList(Of String, Long)()
        For Each v As Bearbeiter In Me.List
            slist.Add(v.name, v.id)
        Next
        Return slist
    End Function

#Region "User Properties"
    Public Function GetById(ByVal id As Long) As Bearbeiter
        For Each v As Bearbeiter In Me.List
            If v.id.Equals(id) Then
                Return v
            End If
        Next
        Return Nothing
    End Function
    Public ReadOnly Property Latest() As Bearbeiter
        Get
            Dim rc As Bearbeiter = Nothing
            Dim currId As Long = Long.MinValue

            For Each v As Bearbeiter In Me.List
                If v.id > currId Then
                    currId = v.id
                    rc = v
                End If
            Next
            Return rc
        End Get
    End Property
#End Region

#Region "Default accessor"
    Default Public Property Item(ByVal intIndex As Integer) As Bearbeiter
        Get
            Return DirectCast(Me.List(intIndex), Bearbeiter)
        End Get
        Set(ByVal value As Bearbeiter)
            Me.List(intIndex) = value
        End Set
    End Property
#End Region

#Region "Clone function"
    Public Function Clone() As Object
        Dim ms As New MemoryStream()
        Dim bf As New BinaryFormatter()
        bf.Serialize(ms, Me)
        ms.Position = 0
        Dim obj As Object = bf.Deserialize(ms)
        ms.Close()
        Return obj
    End Function
#End Region

#Region "Override of Equals function"
    Private Shared inEquals As Boolean = False
    Public Overrides Function Equals(ByVal obj As System.Object) As Boolean
        If inEquals Then
            Return True
        End If
        ' If parameter is null return false.
        If obj Is Nothing Then
            Return False
        End If

        ' If parameter cannot be cast to Point return false.
        Dim p As BearbeiterCollection = TryCast(obj, BearbeiterCollection)
        If DirectCast(p, System.Object) Is Nothing Then
            Return False
        End If

        ' Return true if the fields match:
        Return Equals(p)
    End Function
    Public Overloads Function Equals(ByVal p As BearbeiterCollection) As Boolean
        If inEquals Then
            Return True
        End If
        If DirectCast(p, Object) Is Nothing Then
            Return False
        End If

        inEquals = True
        Dim mySerializer As New XmlSerializer(GetType(BearbeiterCollection))
        Dim xml1 As New StringWriter()
        Dim xml2 As New StringWriter()
        mySerializer.Serialize(xml1, Me)
        mySerializer.Serialize(xml2, p)

        inEquals = False
        Return (xml1.ToString().Equals(xml2.ToString()))
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
#End Region

#Region "Serialization"
    Public Shared Function FromXml(ByVal xml As String) As BearbeiterCollection
        Dim loaded As BearbeiterCollection
        Dim mySerializer As New XmlSerializer(GetType(BearbeiterCollection))
        Dim myReader As New StringReader(xml)
        Try
            loaded = DirectCast(mySerializer.Deserialize(myReader), BearbeiterCollection)
            myReader.Close()
        Catch e As Exception
            MessageBox.Show("Fehler beim Lesen der Bearbeiterdaten." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            loaded = Nothing
        End Try
        Return loaded
    End Function
    Public Function ToXml() As String
        Dim mySerializer As New XmlSerializer(GetType(BearbeiterCollection))
        Dim memStream As New MemoryStream()
        Dim myWriter As New StreamWriter(memStream)
        Dim enc As New UTF8Encoding()
        Try
            mySerializer.Serialize(myWriter, Me)
            myWriter.Close()
        Catch e As Exception
            MessageBox.Show("Fehler beim Prüfen der Bearbeiterdaten." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return [String].Empty
        End Try
        Return enc.GetString(memStream.ToArray())
    End Function
    Public Shared Function ZipFile(ByVal fileName As String) As Boolean
        Dim zip As ZipStorer = Nothing
        Dim tempPath As String = [String].Empty
        Dim tempDir As String = [String].Empty
        Dim versionString As String = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3)

        Try
            tempDir = Path.GetTempFileName()
            File.Delete(tempDir)
            Directory.CreateDirectory(tempDir)
            tempPath = Path.Combine(tempDir, Path.GetFileName(fileName))
            zip = ZipStorer.Create(tempPath, "", versionString)
            zip.AddFile(fileName, Path.GetFileName(fileName), "", versionString)
            zip.Close()
            File.Copy(tempPath, fileName, True)
            Return True
        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Datei" & vbLf & "'" & fileName.ToString() & "':" & ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return False
        Finally
            Try
                zip.Close()
            Catch
            End Try
            Directory.Delete(tempDir, True)
        End Try
    End Function
    Public Shared Function LoadFromFile(ByVal fileName As String) As BearbeiterCollection
        Dim xml As String
        Try
            xml = File.ReadAllText(fileName)
        Catch
            MessageBox.Show([String].Format("Fehler beim Laden der Datei '{0}'", fileName), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return Nothing
        End Try
        Return FromXml(xml)
    End Function
    Public Shared Function LoadFromZip(ByVal fileName As String) As BearbeiterCollection
        Dim zip As ZipStorer = Nothing
        Dim tempDir As String = [String].Empty
        Dim tempXmlFile As String = [String].Empty
        Dim tempZipFile As String = [String].Empty
        Dim versionString As String = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3)
        Dim dir As List(Of ZipStorer.ZipFileEntry)
        Const fmtInvalidVer As String = "Inkompatible Dateiversion." & vbCr & vbLf & vbCr & vbLf & "Die Datei wurde mit der Version {0} erstellt und kann nicht mit der aktuellen Version {1} gelesen werden."

        Try
            Try
                tempDir = Path.GetTempFileName()
                File.Delete(tempDir)
                Directory.CreateDirectory(tempDir)
                tempXmlFile = Path.GetTempFileName()
                File.Delete(tempXmlFile)
                tempZipFile = Path.Combine(tempDir, Path.GetFileName(tempDir))
                tempXmlFile = Path.Combine(tempDir, Path.GetFileName(tempXmlFile))
                File.Copy(fileName, tempZipFile, True)
                zip = ZipStorer.Open(tempZipFile, FileAccess.ReadWrite)
                dir = zip.ReadCentralDir()
                If checkVersion(dir(0).ExtraField) Then
                    zip.ExtractStoredFile(dir(0), tempXmlFile)
                Else
                    MessageBox.Show(String.Format(fmtInvalidVer, dir(0).ExtraField.ToString(), versionString), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    Return New BearbeiterCollection()
                End If
                Return LoadFromFile(tempXmlFile)
            Catch generatedExceptionName As System.Exception
                Throw New Exception(String.Format("Datei konnte nicht geöffnet werden."))

                'try
                '{
                '    return LoadFromFile(tempXmlFile);
                '}
                'catch (System.Exception)
                '{
                '    throw new Exception(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString));
                '}
            End Try
        Catch ex As System.Exception
            MessageBox.Show(ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
        Finally
            If zip IsNot Nothing Then
                zip.Close()
            End If
            Directory.Delete(tempDir, True)
        End Try
        Return New BearbeiterCollection()
    End Function
    Public Function SaveAsZip(ByVal fileName As String) As Boolean
        If SaveToFile(fileName) Then
            Return ZipFile(fileName)
        Else
            Return False
        End If
    End Function
    Public Function SaveToFile(ByVal fileName As String) As Boolean
        Dim rc As Boolean = False
        Try
            Dim xml As String = ToXml()
            If [String].IsNullOrEmpty(xml) Then
                Return rc
            End If
            File.WriteAllText(fileName, xml)
            rc = True
        Catch e As Exception
            MessageBox.Show("Fehler beim Speichern der Normdaten." & vbLf & vbLf & e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.[Error])
            Return rc
        End Try
        Return rc
    End Function

#End Region
End Class

Public Class BearbeiterEnumerator
    Implements IEnumerator
#Region "Fields"
    Private iEnBaseEnumerator As IEnumerator
#End Region

#Region "Implementation of IEnumerator Methods"
    Public Sub New(ByVal mappings As BearbeiterCollection)
        Me.iEnBaseEnumerator = mappings.GetEnumerator()
    End Sub

    Public Function IEnumerator_MoveNext() As Boolean
        Return Me.iEnBaseEnumerator.MoveNext()
    End Function

    Public Sub IEnumerator_Reset()
        Me.iEnBaseEnumerator.Reset()
    End Sub

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Return Me.iEnBaseEnumerator.MoveNext()
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me.iEnBaseEnumerator.Reset()
    End Sub
#End Region

#Region "Implementation of IEnumerator Properties"
    Public ReadOnly Property Current() As Bearbeiter
        Get
            Return DirectCast(Me.iEnBaseEnumerator.Current, Bearbeiter)
        End Get
    End Property

    Public ReadOnly Property IEnumerator_Current() As Object
        Get
            Return Me.iEnBaseEnumerator.Current
        End Get
    End Property

    Private ReadOnly Property System_Collections_IEnumerator_Current() As Object Implements System.Collections.IEnumerator.Current
        Get
            Return Me.iEnBaseEnumerator.Current
        End Get
    End Property
#End Region

End Class
#End Region
