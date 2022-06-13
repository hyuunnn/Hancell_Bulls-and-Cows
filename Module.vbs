Public Declare Function GetTickCount64 Lib "kernel32.dll" () As Long

Sub send()
    inputData = Array(Range("B2:B2").Value, Range("C2:C2").Value, Range("D2:D2").Value, Range("E2:E2").Value)
    comData = Range("H4:H4").Value
    sendCount = Range("B8:B8").Value

    strike = 0
    ball = 0

    If Not(IsEmpty(inputData(0)) Or IsEmpty(inputData(1)) Or IsEmpty(inputData(2)) Or IsEmpty(inputData(3)) Or IsEmpty(comData)) Then
        ' 한셀 문제인지 For Each로 접근해야 값 비교가 가능하다..?
        comData = Split(Range("H4:H4").Value, " ")
        i_idx = 0
        For Each i in inputData
            j_idx = 0
            For Each j in comData
                If (i_idx = j_idx) And (i = j) Then
                    strike = strike + 1
                
                ElseIf (i = j) Then
                    ball = ball + 1

                End If
                j_idx = j_idx + 1
            Next
            i_idx = i_idx + 1
        Next

        startIdx = 9
        sendCount = sendCount + 1
        With Cells(startIdx + sendCount, 6)
            .Value = sendCount
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Cells(startIdx + sendCount, 7)
            .Value = Join(inputData)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With Cells(startIdx + sendCount, 8)
            .Value = strike & "S" & ball & "B"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        Range("B8:B8").Value = sendCount
    Else
        MsgBox("값에 문제가 있습니다.")
    End If
End Sub

Sub makeNumber()
    Dim checkNum(9)

    With Range("H4:H4")
        .Value = result
        .Interior.Color = RGB(0, 0, 0)
    End With
    
    maxNum = 9
    For i = 1 To 4
        Do While 1
            ' 자정으로부터 지난 초를 의미하는 Timer 값으로 seed 값 지정 (소수점 2번째 자리로 계속 변경하는 값)
            ' 직접 바꾸지 않으면 seed 값이 고정되어 있어 똑같은 결과가 나오므로 주기적으로 변경
            Randomize GetTickCount64() + Timer

            ' Int( ( upperbound - lowerbound + 1) * Rnd + lowerbound )
            randomNum = Str(Int(maxNum * Rnd() + 1))
            ' Array를 사용하여 중복 사용 유무 체크
            If checkNum(randomNum) <> 1 Then
                result = result + randomNum
                checkNum(randomNum) = 1
                Exit Do
            End If
        Loop
    Next
    Range("H4:H4").Value = Trim(result)
    Range("B8:B8").Value = 0
    Range("F10:H99999").ClearContents
End Sub

Sub showNumber()
    With Range("H4:L4")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
End Sub