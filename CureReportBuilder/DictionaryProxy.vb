Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Xml.Serialization

Public Class DictionaryProxy(Of K, V)
    Public Sub New(ByVal original As Dictionary(Of K, V))
        original = original
    End Sub

    Public Sub New()
    End Sub

    <XmlIgnore>
    Public Property Original As Dictionary(Of K, V)

    Public Class KeyAndValue
        Public Property Key As K
        Public Property Value As V
    End Class

    Private _list As Collection(Of KeyAndValue)

    <XmlElement>
    Public ReadOnly Property KeysAndValues As Collection(Of KeyAndValue)
        Get

            If _list Is Nothing Then
                _list = New Collection(Of KeyAndValue)()
            End If

            If Original Is Nothing Then
                Return _list
            End If

            _list.Clear()

            For Each pair In Original
                _list.Add(New KeyAndValue With {
                    .Key = pair.Key,
                    .Value = pair.Value
                })
            Next

            Return _list
        End Get
    End Property

    Public Function ToDictionary() As Dictionary(Of K, V)
        Return KeysAndValues.ToDictionary(Function(key) key.Key, Function(value) value.Value)
    End Function
End Class