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

    Public Sub New(Optional inName As String = "", Optional inCureDoc As String = "", Optional inCureDocRev As String = "")
        iName = inName
        icureDoc = inCureDoc
        icureDocRev = inCureDocRev

        ReDim CureSteps(0)
        CureSteps(0) = New CureStep()
    End Sub

    Public Sub deserializeCure(defString As String)

        defString = Replace(Trim(defString), vbTab, "")

        Dim readtext As New IO.StringReader(defString)

        If Left(defString, 1) = vbCr Then
            readtext.ReadLine()
        End If

        If readtext.ReadLine <> "<Cure>" Then Throw New Exception("deserializeCure unrecognized data type")

        inputValue(Name, readtext.ReadLine)
        inputValue(cureDoc, readtext.ReadLine)
        inputValue(cureDocRev, readtext.ReadLine)

        inputValue(checkTemp, readtext.ReadLine)
        inputValue(checkPressure, readtext.ReadLine)
        inputValue(checkVac, readtext.ReadLine)

        If Trim(readtext.ReadLine) <> "<Steps>" Then Throw New Exception("deserializeCure unrecognized data type")

        Dim stepCount As Integer = 0
        inputValue(stepCount, readtext.ReadLine)

        Dim i As Integer
        For i = 1 To stepCount
            addCureStep()
        Next

        For i = 0 To stepCount
            Dim exStep As CureStep = CureSteps(i)

            If Trim(readtext.ReadLine) <> "<Step" & i & ">" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.stepName, readtext.ReadLine)
            inputValue(exStep.stepDuration, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "<pressureSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.pressureSetPoint, readtext.ReadLine)
            inputValue(exStep.pressurePosTol, readtext.ReadLine)
            inputValue(exStep.pressureNegTol, readtext.ReadLine)
            inputValue(exStep.pressureRampRate, readtext.ReadLine)
            inputValue(exStep.pressureRampPosTol, readtext.ReadLine)
            inputValue(exStep.pressureRampNegTol, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</pressureSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<tempSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.tempSetPoint, readtext.ReadLine)
            inputValue(exStep.tempPosTol, readtext.ReadLine)
            inputValue(exStep.tempNegTol, readtext.ReadLine)
            inputValue(exStep.tempRampRate, readtext.ReadLine)
            inputValue(exStep.tempRampPosTol, readtext.ReadLine)
            inputValue(exStep.tempRampNegTol, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</tempSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<vacSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.vacSetPoint, readtext.ReadLine)
            inputValue(exStep.vacPosTol, readtext.ReadLine)
            inputValue(exStep.vacNegTol, readtext.ReadLine)
            inputValue(exStep.vacRampRate, readtext.ReadLine)
            inputValue(exStep.vacRampPosTol, readtext.ReadLine)
            inputValue(exStep.vacRampNegTol, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</vacSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<termCond1>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.termCond1Type, readtext.ReadLine)
            inputValue(exStep.termCond1Condition, readtext.ReadLine)
            inputValue(exStep.termCond1Goal, readtext.ReadLine)
            inputValue(exStep.termCond1Modifier, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</termCond1>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<termCond2>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.termCond2Type, readtext.ReadLine)
            inputValue(exStep.termCond2Condition, readtext.ReadLine)
            inputValue(exStep.termCond2Goal, readtext.ReadLine)
            inputValue(exStep.termCond2Modifier, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</termCond2>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<termCondOper>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.termCondOper, readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</termCondOper>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "</Step" & i & ">" Then Throw New Exception("deserializeCure unrecognized data type")
        Next

        If Trim(readtext.ReadLine) <> "</Steps>" Then Throw New Exception("deserializeCure unrecognized data type")
        If Trim(readtext.ReadLine) <> "</Cure>" Then Throw New Exception("deserializeCure unrecognized data type")

    End Sub

    Sub inputValue(ByRef inVariable As Object, inStr As String)
        Dim values() As String
        values = Split(inStr, ",")
        If UBound(values) = 1 Then
            inVariable = values(1)
        End If
    End Sub

    Public Sub newDeserializeCure(defString As String)

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
            exStep.pressureSetPoint = xmlStyleValRet("pressureSetPoint", pressureSetTxt)
            exStep.pressurePosTol = xmlStyleValRet("pressurePosTol", pressureSetTxt)
            exStep.pressureNegTol = xmlStyleValRet("pressureNegTol", pressureSetTxt)
            exStep.pressureRampRate = xmlStyleValRet("pressureRampRate", pressureSetTxt)
            exStep.pressureRampPosTol = xmlStyleValRet("pressureRampPosTol", pressureSetTxt)
            exStep.pressureRampNegTol = xmlStyleValRet("pressureRampNegTol", pressureSetTxt)

            Dim tempSetTxt As String = xmlStyleValRet("tempSet", thisStepTxt)
            exStep.tempSetPoint = xmlStyleValRet("tempSetPoint", tempSetTxt)
            exStep.tempPosTol = xmlStyleValRet("tempPosTol", tempSetTxt)
            exStep.tempNegTol = xmlStyleValRet("tempNegTol", tempSetTxt)
            exStep.tempRampRate = xmlStyleValRet("tempRampRate", tempSetTxt)
            exStep.tempRampPosTol = xmlStyleValRet("tempRampPosTol", tempSetTxt)
            exStep.tempRampNegTol = xmlStyleValRet("tempRampNegTol", tempSetTxt)

            Dim vacSetTxt As String = xmlStyleValRet("vacSet", thisStepTxt)
            exStep.vacSetPoint = xmlStyleValRet("vacSetPoint", vacSetTxt)
            exStep.vacPosTol = xmlStyleValRet("vacPosTol", vacSetTxt)
            exStep.vacNegTol = xmlStyleValRet("vacNegTol", vacSetTxt)
            exStep.vacRampRate = xmlStyleValRet("vacRampRate", vacSetTxt)
            exStep.vacRampPosTol = xmlStyleValRet("vacRampPosTol", vacSetTxt)
            exStep.vacRampNegTol = xmlStyleValRet("vacRampNegTol", vacSetTxt)

            Dim termCond1Txt As String = xmlStyleValRet("termCond1", thisStepTxt)
            exStep.termCond1Type = xmlStyleValRet("termCond1Type", termCond1Txt)
            exStep.termCond1Condition = xmlStyleValRet("termCond1Condition", termCond1Txt)
            exStep.termCond1Goal = xmlStyleValRet("termCond1Goal", termCond1Txt)
            exStep.termCond1Modifier = xmlStyleValRet("termCond1Modifier", termCond1Txt)

            Dim termCond2Txt As String = xmlStyleValRet("termCond2", thisStepTxt)
            exStep.termCond2Type = xmlStyleValRet("termCond2Type", termCond2Txt)
            exStep.termCond2Condition = xmlStyleValRet("termCond2Condition", termCond2Txt)
            exStep.termCond2Goal = xmlStyleValRet("termCond2Goal", termCond2Txt)
            exStep.termCond2Modifier = xmlStyleValRet("termCond2Modifier", termCond2Txt)

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

    Public Function serializeCure() As String
        Dim retSer As String = ""

        addToSer("<Cure>", retSer)

        addToSer("Name," & Name, retSer, 1)
        addToSer("cureDoc," & cureDoc, retSer, 1)
        addToSer("cureDocRev," & cureDocRev, retSer, 1)

        addToSer("checkTemp," & checkTemp, retSer, 1)
        addToSer("checkPressure," & checkPressure, retSer, 1)
        addToSer("checkVac," & checkVac, retSer, 1)


        addToSer("<Steps>", retSer, 1)
        addToSer("stepCount," & UBound(CureSteps), retSer, 2)

        Dim i As Integer
        For i = 0 To UBound(CureSteps)

            Dim exStep As CureStep = CureSteps(i)

            addToSer("<Step" & i & ">", retSer, 2)

            addToSer("stepName, " & exStep.stepName, retSer, 3)

            addToSer("stepDuration, " & exStep.stepDuration, retSer, 3)

            addToSer("<" & "pressureSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.pressureSetPoint, retSer, 4)
            addToSer("PosTol," & exStep.pressurePosTol, retSer, 4)
            addToSer("NegTol," & exStep.pressureNegTol, retSer, 4)
            addToSer("RampRate," & exStep.pressureRampRate, retSer, 4)
            addToSer("RampPosTol," & exStep.pressureRampPosTol, retSer, 4)
            addToSer("RampNegTol," & exStep.pressureRampNegTol, retSer, 4)
            addToSer("</" & "pressureSet" & ">", retSer, 3)

            addToSer("<" & "tempSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.tempSetPoint, retSer, 4)
            addToSer("PosTol," & exStep.tempPosTol, retSer, 4)
            addToSer("NegTol," & exStep.tempNegTol, retSer, 4)
            addToSer("RampRate," & exStep.tempRampRate, retSer, 4)
            addToSer("RampPosTol," & exStep.tempRampPosTol, retSer, 4)
            addToSer("RampNegTol," & exStep.tempRampNegTol, retSer, 4)
            addToSer("</" & "tempSet" & ">", retSer, 3)

            addToSer("<" & "vacSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.vacSetPoint, retSer, 4)
            addToSer("PosTol," & exStep.vacPosTol, retSer, 4)
            addToSer("NegTol," & exStep.vacNegTol, retSer, 4)
            addToSer("RampRate," & exStep.vacRampRate, retSer, 4)
            addToSer("RampPosTol," & exStep.vacRampPosTol, retSer, 4)
            addToSer("RampNegTol," & exStep.vacRampNegTol, retSer, 4)
            addToSer("</" & "vacSet" & ">", retSer, 3)

            addToSer("<" & "termCond1" & ">", retSer, 3)
            addToSer("Type," & exStep.termCond1Type, retSer, 4)
            addToSer("Condition," & exStep.termCond1Condition, retSer, 4)
            addToSer("Goal," & exStep.termCond1Goal, retSer, 4)
            addToSer("Modifier," & exStep.termCond1Modifier, retSer, 4)
            addToSer("</" & "termCond1" & ">", retSer, 3)

            addToSer("<" & "termCond2" & ">", retSer, 3)
            addToSer("Type," & exStep.termCond2Type, retSer, 4)
            addToSer("Condition," & exStep.termCond2Condition, retSer, 4)
            addToSer("Goal," & exStep.termCond2Goal, retSer, 4)
            addToSer("Modifier," & exStep.termCond2Modifier, retSer, 4)
            addToSer("</" & "termCond2" & ">", retSer, 3)

            addToSer("<" & "termCondOper" & ">", retSer, 3)
            addToSer("termCondOper," & exStep.termCondOper, retSer, 4)
            addToSer("</" & "termCondOper" & ">", retSer, 3)

            addToSer("</Step" & i & ">", retSer, 2)
        Next

        addToSer("</Steps>", retSer, 1)

        addToSer("</Cure>", retSer)

        Return retSer
    End Function

    Public Function newSerializeCure() As String
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
            addToSer("<setPoint>" & exStep.pressureSetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.pressurePosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.pressureNegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.pressureRampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.pressureRampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.pressureRampNegTol & "</RampNegTol>", retSer, 4)
            addToSer("</" & "pressureSet" & ">", retSer, 3)

            addToSer("<" & "tempSet" & ">", retSer, 3)
            addToSer("<setPoint>" & exStep.tempSetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.tempPosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.tempNegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.tempRampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.tempRampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.tempRampNegTol & "</RampNegTol>", retSer, 4)
            addToSer("</" & "tempSet" & ">", retSer, 3)

            addToSer("<" & "vacSet" & ">", retSer, 3)
            addToSer("<setPoint>" & exStep.vacSetPoint & "</setPoint>", retSer, 4)
            addToSer("<PosTol>" & exStep.vacPosTol & "</PosTol>", retSer, 4)
            addToSer("<NegTol>" & exStep.vacNegTol & "</NegTol>", retSer, 4)
            addToSer("<RampRate>" & exStep.vacRampRate & "</RampRate>", retSer, 4)
            addToSer("<RampPosTol>" & exStep.vacRampPosTol & "</RampPosTol>", retSer, 4)
            addToSer("<RampNegTol>" & exStep.vacRampNegTol & "</RampNegTol>", retSer, 4)
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

    Public pressureSetPoint As Double = -1
    Public pressurePosTol As Double = -1
    Public pressureNegTol As Double = -1
    Public pressureRampRate As Double = -1
    Public pressureRampPosTol As Double = -1
    Public pressureRampNegTol As Double = -1

    Public tempSetPoint As Double = -1
    Public tempPosTol As Double = -1
    Public tempNegTol As Double = -1
    Public tempRampRate As Double = -1
    Public tempRampPosTol As Double = -1
    Public tempRampNegTol As Double = -1

    Public vacSetPoint As Double = -1
    Public vacPosTol As Double = -1
    Public vacNegTol As Double = -1
    Public vacRampRate As Double = -1
    Public vacRampPosTol As Double = -1
    Public vacRampNegTol As Double = -1

    Public termCond1Type As String = ""
    Public termCond1Condition As String = ""
    Public termCond1Goal As Double = 0
    Public termCond1Modifier As Object = Nothing

    Public termCond2Type As String = ""
    Public termCond2Condition As String = ""
    Public termCond2Goal As Double = 0
    Public termCond2Modifier As Object = Nothing

    Public termCondOper As String


    Public pressureResultMax As Double = -1
    Public pressureResultMin As Double = -1
    Public pressureResultAvg As Double = -1
    Public pressureResultMaxRamp As Double = -1
    Public pressureResultMinRamp As Double = -1
    Public pressureResultAvgRamp As Double = -1

    Public tempResultMax As Double = -1
    Public tempResultMin As Double = -1
    Public tempResultAvg As Double = -1
    Public tempResultMaxRamp As Double = -1
    Public tempResultMinRamp As Double = -1
    Public tempResultAvgRamp As Double = -1

    Public vacResultMax As Double = -1
    Public vacResultMin As Double = -1
    Public vacResultAvg As Double = -1
    Public vacResultMaxRamp As Double = -1
    Public vacResultMinRamp As Double = -1
    Public vacResultAvgRamp As Double = -1

    Public pressurePass As Boolean = True
    Public tempPass As Boolean = True
    Public vacPass As Boolean = True
    Public pressureRampPass As Boolean = True
    Public tempRampPass As Boolean = True
    Public timePass As Boolean = True

    Public stepPass As Boolean = False
    Public hardFail As Boolean = False
    Public stepTerminate As Boolean = True

    Public stepStart As Integer = -1
    Public stepEnd As Integer = 0


    Sub setPressure(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        pressureSetPoint = inSetPoint
        pressurePosTol = inPosTol
        pressureNegTol = inNegTol
        pressureRampRate = inRampRate
        pressureRampPosTol = inRampPosTol
        pressureRampNegTol = inRampNegTol
    End Sub

    Sub setTemp(inSetPoint As Double, inPosTol As Double, inNegTol As Double, inRampRate As Double, inRampPosTol As Double, inRampNegTol As Double)
        tempSetPoint = inSetPoint
        tempPosTol = inPosTol
        tempNegTol = inNegTol
        tempRampRate = inRampRate
        tempRampPosTol = inRampPosTol
        tempRampNegTol = inRampNegTol
    End Sub

End Class
