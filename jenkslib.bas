' Jenks Natural Breaks Macro
' https://github.com/gomininggo/Jenks-Natural-Breaks
' Translated by: Nick Easton
' From Source: Geostats Javascript library by Simon Geoget and Doug Curl
' Available: https://github.com/simogeo/geostats

Function RunJenks(theRange As Range, numClasses As Integer) As Double()
    'sort data
    theRange.Sort key1:=theRange, order1:=xlAscending, Header:=xlGuess

    'put dataset into an array
    Dim tempArray As Variant
    Dim theData() As Double
    Dim realCount As Integer
    
    tempArray = theRange.Value
    ReDim theData(0 To UBound(tempArray))
    
    For a = 0 To UBound(tempArray) - 1
        If IsNumeric(tempArray(a + 1, 1)) Then 'filter out header/non-numbers
            theData(a) = tempArray(a + 1, 1)
            realCount = realCount + 1
        End If
    Next
    ReDim Preserve theData(0 To realCount)
    
    'Build computation matrices
    Dim ClassLowerLimits()  As Double
    ReDim ClassLowerLimits(0 To UBound(theData), 0 To numClasses) As Double
    Dim VarianceCombos() As Double
    ReDim VarianceCombos(0 To UBound(theData), 0 To numClasses) As Double
    Dim variance As Double

    'Fills variance combo matrix with impossibly high value
    For a = 1 To numClasses
        ClassLowerLimits(1, a) = 1      '?? not sure why
        For b = 2 To UBound(theData)
            VarianceCombos(b, a) = 1E+16    'should be infinity
        Next
    Next
    
    'Calculation loop
    For DataLoop = 2 To UBound(theData)
        Dim GroupSum As Double   'sum of entire dataset
        GroupSum = 0
        Dim GroupSumSqaures As Double    'sum of each value of dataset squared
        GroupSumSqaures = 0
        Dim MemberCount As Integer  'keeps the number of values in current group
        MemberCount = 0
        Dim CurIndex As Double     'Iteration counter
        CurIndex = 0

        For PreviousLoop = 1 To DataLoop   'cycle through all previous data values
            Dim LowerLimIndex As Integer
            LowerLimIndex = DataLoop - PreviousLoop + 1
            Dim curValue As Double
            curValue = theData(LowerLimIndex - 1) 'pull value from dataset
            
            MemberCount = MemberCount + 1
            GroupSum = GroupSum + curValue 'add current value to total sum
            GroupSumSqaures = GroupSumSqaures + curValue * curValue '[curValue] x [curValue] is faster than [curValue]^2
            variance = sum_squares - (GroupSum * GroupSum) / MemberCount    'calc variance of current group
            CurIndex = LowerLimIndex - 1

            If CurIndex <> 0 Then
                For ClassLoop = 2 To numClasses
                    If VarianceCombos(DataLoop, ClassLoop) >= variance + VarianceCombos(CurIndex, ClassLoop - 1) Then
                        ClassLowerLimits(DataLoop, ClassLoop) = LowerLimIndex
                        VarianceCombos(DataLoop, ClassLoop) = variance + VarianceCombos(CurIndex, ClassLoop - 1)
                    End If
                Next
            End If
        Next
        ClassLowerLimits(DataLoop, 1) = l
        VarianceCombos(DataLoop, 1) = variance
    Next

    ' extract classes out of the computed matrices
    Dim k As Integer    'temp variable
    k = UBound(theData)
    Dim countNum As Integer 'count down loop variable
    countNum = numClasses   'output loop to total desired classes
    Dim BreakValues() As Double
    ReDim BreakValues(countNum) As Double
    
    'loop missed first and last values, set them ahead of time
    BreakValues(numClasses) = theData(UBound(theData))
    BreakValues(0) = theData(0)
    
    Do While countNum > 1
        BreakValues(countNum - 1) = theData(ClassLowerLimits(k, countNum) - 2)
        k = ClassLowerLimits(k, countNum) - 1
        countNum = countNum - 1
    Loop
    
    RunJenks = BreakValues
End Function
