Imports System.Xml.Serialization

<Serializable()>
Public Class CureProfile
    Private iName As String = ""
    Private icureDoc As String = ""
    Private icureDocRev As String = ""

    Public inCureSteps() As CureStep


    '<XmlAttribute("Name")>
    Public Property Name As String
        Get
            Return iName
        End Get
        Set(value As String)
            iName = value
        End Set
    End Property

    '<XmlAttribute("cureDoc")>
    Public Property cureDoc As String
        Get
            Return icureDoc
        End Get
        Set(value As String)
            icureDoc = value
        End Set
    End Property

    '<XmlAttribute("cureDocRev")>
    Public Property cureDocRev As String
        Get
            Return icureDocRev
        End Get
        Set(value As String)
            icureDocRev = value
        End Set
    End Property

    ''<XmlAttribute("cureSteps")>
    'Public Property CureSteps() As CureStep()
    '    Get
    '        Return inCureSteps
    '    End Get
    '    Set(value() As CureStep)
    '        inCureSteps = value
    '    End Set
    'End Property


    '<XmlAttribute("cureSteps")>
    Public Property CureStep(ByVal i As Integer) As Object
        Get
            If inCureSteps Is Nothing Then
                ReDim inCureSteps(0)
            End If

            If i <= UBound(inCureSteps) Then
                Return inCureSteps(i)
            Else Return Nothing
            End If
        End Get

        Set(ByVal value As Object)
            If inCureSteps Is Nothing Then
                ReDim inCureSteps(0)
            End If

            If UBound(inCureSteps) <= i Then
                ReDim Preserve inCureSteps(i)
            End If

            inCureSteps(i) = value
        End Set
    End Property




    Sub addCureStep()
        If inCureSteps Is Nothing Then
            ReDim inCureSteps(0)
        Else
            ReDim Preserve inCureSteps(UBound(inCureSteps) + 1)
        End If

        inCureSteps(UBound(inCureSteps)) = New CureStep()
    End Sub

    Public Sub New(inName As String, inCureDoc As String, inCureDocRev As String)
        iName = inName
        icureDoc = inCureDoc
        icureDocRev = inCureDocRev

        ReDim inCureSteps(0)
        inCureSteps(0) = New CureStep()

        'cureSteps(0).pressureSet("negTol") = 5

    End Sub
End Class

<Serializable()>
Public Class CureStep
    Public stepName As String = ""

    'Public pressureSetPoint As Double = 0
    'Public pressurePosTol As Double = 0
    'Public pressureNegTol As Double = 0
    'Public pressureRampRate As Double = 0
    'Public pressureRampPosTol As Double = 0
    'Public pressureRampNegTol As Double = 0

    'Public tempSetPoint As Double = 0
    'Public tempPosTol As Double = 0
    'Public tempNegTol As Double = 0
    'Public tempRampRate As Double = 0
    'Public tempRampPosTol As Double = 0
    'Public tempRampNegTol As Double = 0

    'Public vacSetPoint As Double = 0
    'Public vacPosTol As Double = 0
    'Public vacNegTol As Double = 0
    'Public vacRampRate As Double = 0
    'Public vacRampPosTol As Double = 0
    'Public vacRampNegTol As Double = 0

    Public pressureSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}
    Public tempSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}
    Public vacSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}

    Public termCond1 As New Dictionary(Of String, Object) From {{"Type", ""}, {"Condition", ""}, {"Goal", 0.0}, {"TCNum", 0}}
    Public termCond2 As New Dictionary(Of String, Object) From {{"Type", ""}, {"Condition", ""}, {"Goal", 0.0}, {"TCNum", 0}}

    Sub setPressure(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        pressureSet("SetPoint") = inSetPoint
        pressureSet("PosTol") = inPosTol
        pressureSet("NegTol") = inNegTol
        pressureSet("RampRate") = inRampRate
        pressureSet("RampPosTol") = inRampPosTol
        pressureSet("RampNegTol") = inRampNegTol
    End Sub

    Sub setTemp(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        tempSet("SetPoint") = inSetPoint
        tempSet("PosTol") = inPosTol
        tempSet("NegTol") = inNegTol
        tempSet("RampRate") = inRampRate
        tempSet("RampPosTol") = inRampPosTol
        tempSet("RampNegTol") = inRampNegTol
    End Sub

End Class
