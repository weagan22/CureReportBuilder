Imports System.Runtime.CompilerServices
Imports System.Xml.Serialization

Public Class MainForm

    Public loadedDataSet(,) As String
    Public TC_only As Boolean = False
    Public cureProfiles() As CureProfile

    Public partValues As New Dictionary(Of String, String) From {{"JobNum", ""}, {"PONum", ""}, {"PartNum", ""}, {"PartRev", ""}, {"PartNom", ""}, {"ProgramNum", ""}, {"PartQty", ""}, {"DataPath", ""}}
    Dim dateValues As New Dictionary(Of String, DateTime) From {{"startTime", Nothing}, {"endTime", Nothing}}

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call errorReset()

        Call loadCureProfiles("C:\Users\Will Eagan\Source\Repos\CureReportBuilder\CureReportBuilder\Sample Files\test.cprof")

        'loadCSVin("C:\Users\Will.Eagan\source\repos\CureReportBuilder\CureReportBuilder\Sample Files\DA-18-20.csv")
        'loadCSVin("C:\Users\Will.Eagan\source\repos\CureReportBuilder\CureReportBuilder\Sample Files\BATCH 38 JOB 101573, 101574 1-23-20.CSV")









    End Sub

    Sub errorReset()
        'Clear out the current loaded values if they exist
        If loadedDataSet IsNot Nothing Then
            Array.Clear(loadedDataSet, 0, loadedDataSet.Length)
            loadedDataSet = Nothing
        End If

        'Reset TC_only to false
        TC_only = False

        'Reset partValues to nothing
        partValues("JobNum") = String.Empty
        partValues("PONum") = String.Empty
        partValues("PartNum") = String.Empty
        partValues("PartRev") = String.Empty
        partValues("PartNom") = String.Empty
        partValues("ProgramNum") = String.Empty
        partValues("PartQty") = String.Empty
        partValues("DataPath") = String.Empty

        'Reset dateValues to null values
        dateValues("startTime") = Nothing
        dateValues("endTime") = Nothing

    End Sub

    Sub loadCSVin(inFile As String)

        Call errorReset()

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(inFile)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()

            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    If UBound(currentRow) > 7 Then
                        loadedDataSet.AddArr(currentRow)
                    End If

                    'Finds date in header of Omega TC reader style files
                    If InStr(currentRow(0), "Omega", 0) <> 0 Then
                        TC_only = True
                    End If

                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using

    End Sub

    Sub outputCureProfiles(inPath As String)
        Dim outputWriter As IO.StreamWriter = New System.IO.StreamWriter(inPath)

        Dim i As Integer
        For i = 0 To UBound(cureProfiles)
            outputWriter.Write(cureProfiles(i).serializeCure)
            If i < UBound(cureProfiles) Then
                outputWriter.Write("~&&&~" & vbNewLine)
            End If
        Next

        outputWriter.Close()
    End Sub

    Sub loadCureProfiles(inPath As String)
        Dim cureDef() As String = Split(IO.File.ReadAllText(inPath), "~&&&~")

        For i = 0 To UBound(cureDef)
            If cureProfiles Is Nothing Then
                ReDim cureProfiles(0)
                cureProfiles(i) = New CureProfile()
            Else
                ReDim Preserve cureProfiles(UBound(cureProfiles) + 1)
                cureProfiles(i) = New CureProfile()
            End If

            cureProfiles(i).deserializeCure(cureDef(i))
        Next

    End Sub

End Class

Module ArrayExtensions

    ''' <summary>
    ''' Addes values of a 1D array to a 2D array. 1D array length must match 1st dimension of 2D array (if they do not match, values will not be addded).
    ''' </summary>

    <Extension()>
    Public Sub AddArr(Of T)(ByRef arr As T(,), addArr() As T)
        Dim i As Integer

        If arr IsNot Nothing Then
            If UBound(addArr) <= UBound(arr, 1) Then
                ReDim Preserve arr(UBound(arr, 1), UBound(arr, 2) + 1)

                For i = 0 To UBound(addArr)
                    arr(i, UBound(arr, 2)) = addArr(i)
                Next
            End If
        Else
            ReDim arr(UBound(addArr), 0)
            For i = 0 To UBound(arr, 1)
                arr(i, UBound(arr, 2)) = addArr(i)
            Next
        End If
    End Sub

End Module

