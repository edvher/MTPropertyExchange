Imports System.Collections
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Runtime.Serialization.Formatters.Binary

<Serializable(), XmlType("Property"), ComVisible(False)> _
Public Class DocumentProperty
    <XmlAttribute("Name")> Public Name As String
    <XmlAttribute("Value")> Public Value As String
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger( _
        System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region "    Constructors"
    Public Sub New()
        Me.New(String.Empty, String.Empty)
    End Sub
    Public Sub New(ByVal name As String, ByVal value As String)
        Me.Value = value
        Me.Name = name
    End Sub
#End Region

#Region "    Clone function"
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

#Region "    Override of Equals function"
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
        Dim p As DocumentProperty = TryCast(obj, DocumentProperty)
        If DirectCast(p, System.Object) Is Nothing Then
            Return False
        End If

        ' Return true if the fields match:
        Return Equals(p)
    End Function
    Public Overloads Function Equals(ByVal p As DocumentProperty) As Boolean
        If inEquals Then
            Return True
        End If
        If DirectCast(p, Object) Is Nothing Then
            Return False
        End If

        inEquals = True
        Dim mySerializer As New XmlSerializer(GetType(DocumentProperty))
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

<Serializable(), XmlType("Properties"), ComVisible(False)> _
Public Class DocumentProperties
    Inherits CollectionBase
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger( _
        System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region "    Constructors"
    Public Sub New()
    End Sub
#End Region

#Region "    Implementation of CollectionBase"
    Public Function Add(ByVal name As String, ByVal value As String) As Integer
        Return Me.List.Add(New DocumentProperty(name, value))
    End Function

    Public Function Add(ByVal value As DocumentProperty) As Integer
        Return Me.List.Add(value)
    End Function

    Public Sub AddRange(ByVal value() As DocumentProperty)
        For i As Integer = 0 To value.Length - 1
            Me.Add(value(i))
        Next
    End Sub

    Public Sub AddRange(ByVal value As DocumentProperties)
        For i As Integer = 0 To value.Count - 1
            Me.Add(value(i))
        Next
    End Sub

    Public Function Contains(ByVal value As DocumentProperty) As Boolean
        Return Me.List.Contains(value)
    End Function

    Public Sub CopyTo(ByVal layArray() As DocumentProperty, ByVal intIndex As Integer)
        Me.List.CopyTo(layArray, intIndex)
    End Sub

    Public Overloads Function GetEnumerator() As IEnumerator 'DocumentPropertyEnumerator
        Return DirectCast(List, IEnumerable).GetEnumerator 'New DocumentPropertyEnumerator(Me)
    End Function

    Public Function IndexOf(ByVal value As DocumentProperty) As Integer
        Return Me.List.IndexOf(value)
    End Function

    Public Function IndexOf(ByVal name As String) As Integer
        Dim num3 As Integer = Me.List.Count - 1
        For i As Integer = 0 To num3
            If DirectCast(Me.List(i), DocumentProperty).Name.Equals(name) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Sub Insert(ByVal intIndex As Integer, ByVal value As DocumentProperty)
        Me.List.Insert(intIndex, value)
    End Sub

    Protected Overrides Sub OnClear()
        While Me.List.Count > 0
            Dim value As DocumentProperty = DirectCast(Me.List(0), DocumentProperty)
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

    Public Sub Remove(ByVal value As DocumentProperty)
        Me.List.Remove(value)
    End Sub
#End Region

#Region "    Sorting"
    Public Function GetSortedList() As SortedList(Of String, Long)
        Dim slist As New SortedList(Of String, Long)()
        For Each v As DocumentProperty In Me.List
            slist.Add(v.Name, v.Value)
        Next
        Return slist
    End Function

    Public Function Sort() As DocumentProperties
        Dim slist As New SortedList(Of String, String)()
        For Each v As DocumentProperty In Me.List
            slist.Add(v.Name, v.Value)
        Next
        Me.List.Clear()

        For i As Integer = 0 To slist.Count - 1
            Me.List.Add(New DocumentProperty(slist.Keys(i), slist.Values(i)))
        Next
        Return Me
    End Function
#End Region

#Region "    Item accessors"
    Default Public Property Item(ByVal intIndex As Integer) As DocumentProperty
        Get
            Return DirectCast(Me.List(intIndex), DocumentProperty)
        End Get
        Set(ByVal value As DocumentProperty)
            Me.List(intIndex) = value
        End Set
    End Property

    Default Public Property Item(ByVal name As String) As DocumentProperty
        Get
            For Each p As DocumentProperty In Me.List
                If p.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase) Then Return p
            Next
            Return Nothing
        End Get
        Set(ByVal value As DocumentProperty)
            For Each p As DocumentProperty In Me.List
                If p.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase) Then p.Value = value.Value
            Next
        End Set
    End Property

#End Region

#Region "    Clone function"
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

#Region "    Override of Equals function"
    'Private Shared inEquals As Boolean = False
    Public Overrides Function Equals(ByVal obj As System.Object) As Boolean
        'If inEquals Then
        '    Return True
        'End If
        ' If parameter is null return false.
        If obj Is Nothing Then
            Return False
        End If

        ' If parameter cannot be cast to Point return false.
        Dim p As DocumentProperties = TryCast(obj, DocumentProperties)
        If DirectCast(p, System.Object) Is Nothing Then
            Return False
        End If

        ' Return true if the fields match:
        Return Equals(p)
    End Function
    Public Overloads Function Equals(ByVal p As DocumentProperties) As Boolean
        'If inEquals Then
        '    Return True
        'End If
        If DirectCast(p, Object) Is Nothing Then
            Return False
        End If

        'inEquals = True
        Dim mySerializer As New XmlSerializer(GetType(DocumentProperties))
        Dim xml1 As New StringWriter()
        Dim xml2 As New StringWriter()
        mySerializer.Serialize(xml1, Me)
        mySerializer.Serialize(xml2, p)

        'inEquals = False
        Return (xml1.ToString().Equals(xml2.ToString()))
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return 0
    End Function
#End Region

#Region "    Serialization"
#Region "        Unicode XML"
    Public Function ToXml() As String
        Return ToXml("utf-8").TrimStart()
    End Function

    Public Function ToXml(ByVal encodingName As String) As String
        Dim rs As String

        Dim mySerializer As New XmlSerializer(GetType(DocumentProperties))
        Dim memStream As New MemoryStream()
        Dim enc As Encoding = System.Text.Encoding.GetEncoding(encodingName)
        Dim myWriter As New StreamWriter(memStream, enc)

        Try
            mySerializer.Serialize(myWriter, Me)
            myWriter.Close()
        Catch e As Exception
            rs = String.Empty
        End Try
        rs = enc.GetString(memStream.ToArray())

        Return rs
    End Function
    Public Shared Function FromXml(ByVal xml As String) As DocumentProperties
        Dim loaded As DocumentProperties
        Dim mySerializer As New XmlSerializer(GetType(DocumentProperties))
        Dim myReader As New StringReader(xml)

        Try
            loaded = DirectCast(mySerializer.Deserialize(myReader), DocumentProperties)
            myReader.Close()
        Catch e As Exception
            loaded = Nothing
        Finally
            If myReader IsNot Nothing Then myReader.Close()
            'If stream IsNot Nothing Then stream.Close()
        End Try
        Return loaded
    End Function
#End Region

#Region "        To File"
    Public Shared Function LoadFromTextFile(ByVal fileName As String) As DocumentProperties
        Return LoadFromTextFile(fileName, "utf-8")
    End Function
    Public Shared Function LoadFromTextFile(ByVal fileName As String, ByVal encodingName As String) As DocumentProperties
        Dim props As New DocumentProperties()
        Dim line As String
        Dim enc = System.Text.Encoding.GetEncoding(encodingName)
        Dim rdr As New StreamReader(fileName, enc)
        line = rdr.ReadLine()
        While Not line Is Nothing
            Dim nameval() As String = line.Split("=".ToCharArray(), StringSplitOptions.None)
            If Not nameval(0).StartsWith("[", StringComparison.InvariantCultureIgnoreCase) Then
                If nameval.Length > 1 Then
                    If nameval.Length = 2 Then
                        log.DebugFormat("Reading Property {0} = {1}", nameval(0), nameval(1))
                        props.Add(nameval(0), nameval(1))
                    Else
                        Dim val2 As String = nameval(1)
                        For i As Integer = 2 To nameval.Length - 1
                            val2 = val2 + "=" + nameval(i)
                        Next
                        log.DebugFormat("Reading Property {0} = {1}", nameval(0), val2)
                        props.Add(nameval(0), val2)
                    End If
                Else
                    System.Environment.ExitCode = Globals.Errors.EncodingNotSupported
                    Exit While
                End If
            End If
            line = rdr.ReadLine()
        End While
        rdr.Close()

        Return props
    End Function

    Public Shared Function LoadFromFile(ByVal fileName As String) As DocumentProperties
        Dim xml As String
        Try
            xml = File.ReadAllText(fileName).Trim
        Catch
            Return Nothing
        End Try
        Return FromXml(xml)
    End Function
    Public Function SaveToFile(ByVal fileName As String, ByVal encodingName As String) As Boolean
        Dim enc = System.Text.Encoding.GetEncoding(encodingName)
        Dim rc As Boolean = False
        Dim xml As String = ToXml(encodingName).Trim
        Try
            File.WriteAllText(fileName, xml, enc)
            rc = True
        Catch
        End Try
        Return rc
    End Function
    Public Function SaveToFile(ByVal fileName As String) As Boolean
        Return SaveToFile(fileName, "utf-8")
    End Function
#End Region
#End Region
End Class

<ComVisible(False)> _
Public Class DocumentPropertyEnumerator
    Implements IEnumerator
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger( _
        System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
#Region "Fields"
    Private iEnBaseEnumerator As IEnumerator
#End Region

#Region "Implementation of IEnumerator Methods"
    Public Sub New(ByVal mappings As DocumentProperties)
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

#Region "Implementation of IEnumerator DocumentProperties"
    Public ReadOnly Property Current() As DocumentProperty
        Get
            Return DirectCast(Me.iEnBaseEnumerator.Current, DocumentProperty)
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