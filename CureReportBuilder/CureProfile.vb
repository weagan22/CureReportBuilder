Option Explicit On

Imports System.Text.RegularExpressions

Public Class CureProfile
    Private iName As String = ""
    Private icureDoc As String = ""
    Private icureDocRev As String = ""

    Public curePass As Boolean = True

    Public checkTemp As Boolean = True
    Public checkPressure As Boolean = True
    Public checkVac As Boolean = True

    Public fileEditDate As Date

    Public CureSteps() As CureStep

    Public Property Name As String
        Get
            Return iName
        End Get
        Set(value As String)
            iName = value
        End Set
    End Property

    Public Property cureDoc As String
        Get
            Return icureDoc
        End Get
        Set(value As String)
            icureDoc = value
        End Set
    End Property

    Public Property cureDocRev As String
        Get
            Return icureDocRev
        End Get
        Set(value As String)
            icureDocRev = value
        End Set
    End Property

    Sub addCureStep()
        If CureSteps Is Nothing Then
            ReDim CureSteps(0)
        Else
            ReDim Preserve CureSteps(UBound(CureSteps) + 1)
        End If

        CureSteps(UBound(CureSteps)) = New CureStep()
    End Sub

    Public Sub New(Optional inName As String = "",
                   Optional inCureDoc As String = "",
                   Optional inCureDocRev As String = "")

        iName = inName
        icureDoc = inCureDoc
        icureDocRev = inCureDocRev

        ReDim CureSteps(0)
        CureSteps(0) = New CureStep()
    End Sub


    Public Sub DeserializeCure(defString As String)

        defString = Replace(Trim(defString), vbTab, "")

        Dim cureTxt As String = xmlStyleValRet("Cure", defString)
        Name = xmlStyleValRet("Name", cureTxt)
        cureDoc = xmlStyleValRet("cureDoc", cureTxt)
        cureDocRev = xmlStyleValRet("cureDocRev", cureTxt)
        checkTemp = xmlStyleValRet("checkTemp", cureTxt)
        checkPressure = xmlStyleValRet("checkPressure", cureTxt)
        checkVac = xmlStyleValRet("checkVac", cureTxt)

        Dim stepsTxt As String = xmlStyleValRet("Steps", cureTxt)

        Dim stepCount As Integer = 0
        stepCount = xmlStyleValRet("stepCount", stepsTxt)

        Dim i As Integer
        For i = 1 To stepCount
            addCureStep()
        Next

        For i = 0 To stepCount
            Dim exStep As CureStep = CureSteps(i)

            Dim thisStepTxt As String = xmlStyleValRet("Step" & i, stepsTxt)
            exStep.stepName = xmlStyleValRet("stepName", thisStepTxt)
            exStep.stepDuration = xmlStyleValRet("stepDuration", thisStepTxt)

            Dim pressureSetTxt As String = xmlStyleValRet("pressureSet", thisStepTxt)
            exStep.pressureSet.SetPoint = xmlStyleValRet("setPoint", pressureSetTxt)
            exStep.pressureSet.PosTol = xmlStyleValRet("PosTol", pressureSetTxt)
            exStep.pressureSet.NegTol = xmlStyleValRet("NegTol", pressureSetTxt)
            exStep.pressureSet.RampRate = xmlStyleValRet("RampRate", pressureSetTxt)
            exStep.pressureSet.RampPosTol = xmlStyleValRet("RampPosTol", pressureSetTxt)
            exStep.pressureSet.RampNegTol = xmlStyleValRet("RampNegTol", pressureSetTxt)

            Dim tempSetTxt As String = xmlStyleValRet("tempSet", thisStepTxt)
            exStep.tempSet.SetPoint = xmlStyleValRet("setPoint", tempSetTxt)
            exStep.tempSet.PosTol = xmlStyleValRet("PosTol", tempSetTxt)
            exStep.tempSet.NegTol = xmlStyleValRet("NegTol", tempSetTxt)
            exStep.tempSet.RampRate = xmlStyleValRet("RampRate", tempSetTxt)
            exStep.tempSet.RampPosTol = xmlStyleValRet("RampPosTol", tempSetTxt)
            exStep.tempSet.RampNegTol = xmlStyleValRet("RampNegTol", tempSetTxt)

            Dim vacSetTxt As String = xmlStyleValRet("vacSet", thisStepTxt)
            exStep.vacSet.SetPoint = xmlStyleValRet("SetPoint", vacSetTxt)
            exStep.vacSet.PosTol = xmlStyleValRet("PosTol", vacSetTxt)
            exStep.vacSet.NegTol = xmlStyleValRet("NegTol", vacSetTxt)
            exStep.vacSet.RampRate = xmlStyleValRet("RampRate", vacSetTxt)
            exStep.vacSet.RampPosTol = xmlStyleValRet("RampPosTol", vacSetTxt)
            exStep.vacSet.RampNegTol = xmlStyleValRet("RampNegTol", vacSetTxt)

            Dim termCond1Txt As String = xmlStyleValRet("termCond1", thisStepTxt)
            exStep.termCond1Type = xmlStyleValRet("Type", termCond1Txt)
            exStep.termCond1Condition = xmlStyleValRet("Condition", termCond1Txt)
            exStep.termCond1Goal = xmlStyleValRet("Goal", termCond1Txt)
            exStep.termCond1Modifier = xmlStyleValRet("Modifier", termCond1Txt)

            Dim termCond2Txt As String = xmlStyleValRet("termCond2", thisStepTxt)
            exStep.termCond2Type = xmlStyleValRet("Type", termCond2Txt)
            exStep.termCond2Condition = xmlStyleValRet("Condition", termCond2Txt)
            exStep.termCond2Goal = xmlStyleValRet("Goal", termCond2Txt)
            exStep.termCond2Modifier = xmlStyleValRet("Modifier", termCond2Txt)

            exStep.termCondOper = xmlStyleValRet("termCondOper", thisStepTxt)

        Next

    End Sub

    Function xmlStyleValRet(inSearch As String, inString As String) As String
        Try
            Dim pattern As String = "(?i)(?<=<" & inSearch & ">)(.*)(?=</" & inSearch & ">)"
            Dim retVals As Match = Regex.Match(inString, pattern, RegexOptions.Singleline)

            If retVals.Success = False Then
                Throw New Exception("Failed to get value out of string pattern")
            End If

            If retVals.Captures.Count > 1 Then
                Throw New Exception("Found more than one value in string")
            End If

            Return retVals.Captures.Item(0).Value

            Dim test = 0
        Catch ex As Exception
            Throw New Exception("Failed to get value out of string pattern")
        End Try
    End Function


    Public Function SerializeCure() As String
        Dim retSer As String = ""

        addToSer("<Cure>", retSer)

        addToSer("<Name>" & Name & "</Name>", retSer, 1)
        addToSer("<cureDoc>" & cureDoc & "</cureDoc>", retSer, 1)
        addToSer("<cureDocRev>" & cureDocRev & "</cureDocRev>", retSer, 1)

        addToSer("<checkTemp>" & checkTemp & "</checkTemp>", retSer, 1)
        addToSer("<checkPressure>" & checkPressure & "</checkPressure>", retSer, 1)
        addToSer("<checkVac>" & checkVac & "</checkVac>", retSer, 1)


        addToSer("<Steps>", retSer, 1)
        addToSer("<stepCount>" & UBound(CureSteps) & "</stepCount>", retSer, 2)

        Dim i As Integer
        For i = 0 To UBound(CureSteps)

            Dim exStep As CureStep = CureSteps(i)

            addToSer("<Step" & i & ">", retSer, 2)

            addToSer("<stepName>" & exStep.stepName & "</stepName>", retSer, 3)

            addToSer("<stepDuration>" & exStep.stepDuration & "</stepDuration>", retSer, 3)

            addToSer("<" & "pressureSet" & ">", retSer, 3)
            addToSer("<setPoint>" & exStep.pressureSet.SetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.pressureSet.PosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.pressureSet.NegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.pressureSet.RampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.pressureSet.RampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.pressureSet.RampNegTol & "</RampNegTol>", retSer, 4)
            addToSer("</" & "pressureSet" & ">", retSer, 3)

            addToSer("<" & "tempSet" & ">", retSer, 3)
            addToSer("<setPoint>" & exStep.tempSet.SetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.tempSet.PosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.tempSet.NegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.tempSet.RampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.tempSet.RampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.tempSet.RampNegTol & "</RampNegTol>", retSer, 4)
            addToSer("</" & "tempSet" & ">", retSer, 3)

            addToSer("<" & "vacSet" & ">", retSer, 3)
            addToSer("<setPoint>" & exStep.vacSet.SetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.vacSet.PosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.vacSet.NegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.vacSet.RampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.vacSet.RampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.vacSet.RampNegTol & "</RampNegTol>", retSer, 4)
            addToSer("</" & "vacSet" & ">", retSer, 3)

            addToSer("<" & "termCond1" & ">", retSer, 3)
            addToSer("<Type>" & exStep.termCond1Type & "</Type>", retSer, 4)
            addToSer("<Condition>" & exStep.termCond1Condition & "</Condition>", retSer, 4)
            addToSer("<Goal>" & exStep.termCond1Goal & "</Goal>", retSer, 4)
            addToSer("<Modifier>" & exStep.termCond1Modifier & "</Modifier>", retSer, 4)
            addToSer("</" & "termCond1" & ">", retSer, 3)

            addToSer("<" & "termCond2" & ">", retSer, 3)
            addToSer("<Type>" & exStep.termCond2Type & "</Type>", retSer, 4)
            addToSer("<Condition>" & exStep.termCond2Condition & "</Condition>", retSer, 4)
            addToSer("<Goal>" & exStep.termCond2Goal & "</Goal>", retSer, 4)
            addToSer("<Modifier>" & exStep.termCond2Modifier & "</Modifier>", retSer, 4)
            addToSer("</" & "termCond2" & ">", retSer, 3)

            addToSer("<termCondOper>" & exStep.termCondOper & "</termCondOper>", retSer, 3)

            addToSer("</Step" & i & ">", retSer, 2)
        Next

        addToSer("</Steps>", retSer, 1)

        addToSer("</Cure>", retSer)

        Return retSer
    End Function

    Private Sub addToSer(addMe As String, ByRef toSerial As String, Optional tabNum As Integer = 0)
        Dim tabSpace As String = ""

        Dim i As Integer
        For i = 1 To tabNum
            tabSpace = tabSpace & vbTab
        Next
        toSerial = toSerial & tabSpace & addMe & vbNewLine
    End Sub


End Class



Public Class CureStep
    Public stepName As String = ""

    Public stepDuration As Double = -1

    Public pressureSet As SetValues = New SetValues
    Public tempSet As SetValues = New SetValues
    Public vacSet As SetValues = New SetValues

    Public termCond1Type As String = ""
    Public termCond1Condition As String = ""
    Public termCond1Goal As Double = 0
    Public termCond1Modifier As Object = Nothing

    Public termCond2Type As String = ""
    Public termCond2Condition As String = ""
    Public termCond2Goal As Double = 0
    Public termCond2Modifier As Object = Nothing

    Public termCondOper As String

    Public pressureResult As ResultValues = New ResultValues
    Public tempResult As ResultValues = New ResultValues
    Public vacResult As ResultValues = New ResultValues

    Public pressurePass As Boolean = True
    Public tempPass As Boolean = True
    Public vacPass As Boolean = True
    Public pressureRampPass As Boolean = True
    Public tempRampPass As Boolean = True
    Public timeLimitPass As Boolean = True
    Public soakStepPass As Boolean = False

    Public stepPass As Boolean = False
    Public hardFail As Boolean = False
    Public stepTerminate As Boolean = True

    Public stepStart As Integer = -1
    Public stepEnd As Integer = 0


    Sub setPressure(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        pressureSet.SetPoint = inSetPoint
        pressureSet.PosTol = inPosTol
        pressureSet.NegTol = inNegTol
        pressureSet.RampRate = inRampRate
        pressureSet.RampPosTol = inRampPosTol
        pressureSet.RampNegTol = inRampNegTol
    End Sub

    Sub setTemp(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        tempSet.SetPoint = inSetPoint
        tempSet.PosTol = inPosTol
        tempSet.NegTol = inNegTol
        tempSet.RampRate = inRampRate
        tempSet.RampPosTol = inRampPosTol
        tempSet.RampNegTol = inRampNegTol
    End Sub

End Class

Public Class SetValues
    Public SetPoint As Double = -1
    Public PosTol As Double = -1
    Public NegTol As Double = -1
    Public RampRate As Double = -1
    Public RampPosTol As Double = -1
    Public RampNegTol As Double = -1
End Class

Public Class ResultValues
    Public Max As Double = -1
    Public Min As Double = -1
    Public Avg As Double = -1
    Public MaxRamp As Double = -1
    Public MinRamp As Double = -1
    Public AvgRamp As Double = -1
End Class
