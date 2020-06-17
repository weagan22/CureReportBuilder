Imports System.Runtime.CompilerServices

Public Class MainForm

    Public loadedDataSet(,) As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadCSVin("C:\Users\Will.Eagan\Desktop\DA-18-20.csv")
    End Sub

    Sub loadCSVin(inFile As String)
        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(inFile)
            'ReDim loadedDataSet(0, 0)

            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()

            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    If UBound(currentRow) > 7 Then
                        loadedDataSet.Add(currentRow)
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using
    End Sub

End Class

Module ArrayExtensions
    <Extension()>
    Public Sub Add(Of T)(ByRef arr As T(,), item() As T)

        Dim i As Integer

        If arr IsNot Nothing Then
            ReDim Preserve arr(UBound(arr, 1), UBound(arr, 2) + 1)

            For i = 0 To UBound(arr, 1)
                arr(i, UBound(arr, 2)) = item(i)
            Next

        Else
            ReDim arr(UBound(item), 0)
            For i = 0 To UBound(arr, 1)
                arr(i, UBound(arr, 2)) = item(i)
            Next
        End If
    End Sub
End Module

