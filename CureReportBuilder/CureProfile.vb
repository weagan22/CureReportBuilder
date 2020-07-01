Imports System.Xml.Serialization

<Serializable()>
Public Class CureProfile
    Private iName As String = ""
    Private icureDoc As String = ""
    Private icureDocRev As String = ""

    'Public cureSteps() As CureStep

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


    'Public Sub New(inName As String, inCureDoc As String, inCureDocRev As String)
    '    iName = inName
    '    icureDoc = inCureDoc
    '    icureDocRev = inCureDocRev

    '    ReDim cureSteps(0)
    '    cureSteps(0) = New CureStep()

    '    cureSteps(0).pressureSet("negTol") = 5

    'End Sub
End Class

Public Class CureStep
    Public stepName As String = ""
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
