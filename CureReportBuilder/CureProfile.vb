
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
            If Trim(readtext.ReadLine) <> "<pressureSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.pressureSet("SetPoint"), readtext.ReadLine)
            inputValue(exStep.pressureSet("PosTol"), readtext.ReadLine)
            inputValue(exStep.pressureSet("NegTol"), readtext.ReadLine)
            inputValue(exStep.pressureSet("RampRate"), readtext.ReadLine)
            inputValue(exStep.pressureSet("RampPosTol"), readtext.ReadLine)
            inputValue(exStep.pressureSet("RampNegTol"), readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</pressureSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<tempSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.tempSet("SetPoint"), readtext.ReadLine)
            inputValue(exStep.tempSet("PosTol"), readtext.ReadLine)
            inputValue(exStep.tempSet("NegTol"), readtext.ReadLine)
            inputValue(exStep.tempSet("RampRate"), readtext.ReadLine)
            inputValue(exStep.tempSet("RampPosTol"), readtext.ReadLine)
            inputValue(exStep.tempSet("RampNegTol"), readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</tempSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<vacSet>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.vacSet("SetPoint"), readtext.ReadLine)
            inputValue(exStep.vacSet("PosTol"), readtext.ReadLine)
            inputValue(exStep.vacSet("NegTol"), readtext.ReadLine)
            inputValue(exStep.vacSet("RampRate"), readtext.ReadLine)
            inputValue(exStep.vacSet("RampPosTol"), readtext.ReadLine)
            inputValue(exStep.vacSet("RampNegTol"), readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</vacSet>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<termCond1>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.termCond1("Type"), readtext.ReadLine)
            inputValue(exStep.termCond1("Condition"), readtext.ReadLine)
            inputValue(exStep.termCond1("Goal"), readtext.ReadLine)
            inputValue(exStep.termCond1("TCNum"), readtext.ReadLine)
            If Trim(readtext.ReadLine) <> "</termCond1>" Then Throw New Exception("deserializeCure unrecognized data type")

            If Trim(readtext.ReadLine) <> "<termCond2>" Then Throw New Exception("deserializeCure unrecognized data type")
            inputValue(exStep.termCond2("Type"), readtext.ReadLine)
            inputValue(exStep.termCond2("Condition"), readtext.ReadLine)
            inputValue(exStep.termCond2("Goal"), readtext.ReadLine)
            inputValue(exStep.termCond2("TCNum"), readtext.ReadLine)
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

            addToSer("<" & "pressureSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.pressureSet("SetPoint"), retSer, 4)
            addToSer("PosTol," & exStep.pressureSet("PosTol"), retSer, 4)
            addToSer("NegTol," & exStep.pressureSet("NegTol"), retSer, 4)
            addToSer("RampRate," & exStep.pressureSet("RampRate"), retSer, 4)
            addToSer("RampPosTol," & exStep.pressureSet("RampPosTol"), retSer, 4)
            addToSer("RampNegTol," & exStep.pressureSet("RampNegTol"), retSer, 4)
            addToSer("</" & "pressureSet" & ">", retSer, 3)

            addToSer("<" & "tempSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.tempSet("SetPoint"), retSer, 4)
            addToSer("PosTol," & exStep.tempSet("PosTol"), retSer, 4)
            addToSer("NegTol," & exStep.tempSet("NegTol"), retSer, 4)
            addToSer("RampRate," & exStep.tempSet("RampRate"), retSer, 4)
            addToSer("RampPosTol," & exStep.tempSet("RampPosTol"), retSer, 4)
            addToSer("RampNegTol," & exStep.tempSet("RampNegTol"), retSer, 4)
            addToSer("</" & "tempSet" & ">", retSer, 3)

            addToSer("<" & "vacSet" & ">", retSer, 3)
            addToSer("setPoint," & exStep.vacSet("SetPoint"), retSer, 4)
            addToSer("PosTol," & exStep.vacSet("PosTol"), retSer, 4)
            addToSer("NegTol," & exStep.vacSet("NegTol"), retSer, 4)
            addToSer("RampRate," & exStep.vacSet("RampRate"), retSer, 4)
            addToSer("RampPosTol," & exStep.vacSet("RampPosTol"), retSer, 4)
            addToSer("RampNegTol," & exStep.vacSet("RampNegTol"), retSer, 4)
            addToSer("</" & "vacSet" & ">", retSer, 3)

            addToSer("<" & "termCond1" & ">", retSer, 3)
            addToSer("Type," & exStep.termCond1("Type"), retSer, 4)
            addToSer("Condition," & exStep.termCond1("Condition"), retSer, 4)
            addToSer("Goal," & exStep.termCond1("Goal"), retSer, 4)
            addToSer("TCNum," & exStep.termCond1("TCNum"), retSer, 4)
            addToSer("</" & "termCond1" & ">", retSer, 3)

            addToSer("<" & "termCond2" & ">", retSer, 3)
            addToSer("Type," & exStep.termCond2("Type"), retSer, 4)
            addToSer("Condition," & exStep.termCond2("Condition"), retSer, 4)
            addToSer("Goal," & exStep.termCond2("Goal"), retSer, 4)
            addToSer("TCNum," & exStep.termCond2("TCNum"), retSer, 4)
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

    Public pressureSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}
    Public tempSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}
    Public vacSet As New Dictionary(Of String, Double) From {{"SetPoint", 0}, {"PosTol", 0}, {"NegTol", 0}, {"RampRate", 0}, {"RampPosTol", 0}, {"RampNegTol", 0}}

    Public termCond1 As New Dictionary(Of String, Object) From {{"Type", ""}, {"Condition", ""}, {"Goal", 0.0}, {"TCNum", 0}}
    Public termCond2 As New Dictionary(Of String, Object) From {{"Type", ""}, {"Condition", ""}, {"Goal", 0.0}, {"TCNum", 0}}
    Public termCondOper As String


    Public pressureResult As New Dictionary(Of String, Double) From {{"Max", 0}, {"Min", 0}, {"Avg", 0}, {"MaxRamp", 0}, {"MinRamp", 0}, {"AvgRamp", 0}}
    Public tempResult As New Dictionary(Of String, Double) From {{"Max", 0}, {"Min", 0}, {"Avg", 0}, {"MaxRamp", 0}, {"MinRamp", 0}, {"AvgRamp", 0}}
    Public vacResult As New Dictionary(Of String, Double) From {{"Max", 0}, {"Min", 0}, {"Avg", 0}, {"MaxRamp", 0}, {"MinRamp", 0}, {"AvgRamp", 0}}

    Public pressurePass As Boolean = True
    Public tempPass As Boolean = True
    Public vacPass As Boolean = True
    Public pressureRampPass As Boolean = True
    Public tempRampPass As Boolean = True

    Public stepPass As Boolean = False
    Public hardFail As Boolean = False
    Public stepTerminate As Boolean = True

    Public stepStart As Integer = -1
    Public stepEnd As Integer = 0


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
