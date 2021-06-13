VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'day가 없어서 시험이 치러질 수 없게 되는지를 확인해주는 변수
Dim can_start As Boolean
'시험이 시작되었는지를 구분할 변수
Dim test_start As Boolean
'범위의 시작점과 끝점
Dim range__(2) As Integer
'문제 시트 변수
Dim Days() As Worksheet
'문제들 위치
Dim problem_pos() As Integer
'문제들
Dim problems() As String
'답들
Dim answers() As String
'문제의 갯수
Dim num_of_test
'day의 갯수
Dim num_of_days As Integer
'인덱스값
Dim index_i As Integer
'정답 번호
Dim answer_number As Integer
'겹치는지 체크해주는 변수
Dim checking As Boolean

'선택된 번호
Dim selected_number As Integer
'선택되어졌는지 확인해주는 변수
Dim selecting As Boolean

Dim 틀린개수 As Integer

Dim check_up As Boolean
Dim check_right As Boolean
Dim where_it_is(2) As Integer





'시작버튼과 함께 시작될 매크로
Sub start()
    If can_start = False Then
        MsgBox "단어들을 먼저 입력하고 시작해주세요"
        Exit Sub
    End If
    
    '범위가 시작이 가능한지를 먼저 체크한다.
    If Cells(5, 3).Value > Cells(5, 5).Value Then
        MsgBox "범위는 시작점이 종점보다 작아야 합니다."
        Exit Sub
    End If
    
    If Cells(8, 3).Value < 1 Then
        MsgBox "문제의 수는 0보다 많아야 합니다."
        Exit Sub
    End If
    
    range__(1) = Cells(5, 3).Value
    range__(2) = Cells(5, 5).Value
    
    '해당 범위에서 나올 수 있는 문제의 최대 갯수
    num_of_test = 0
    For i = range__(1) To range__(2)
        num_of_test = Days(i).Cells(3, 4).Value + num_of_test
    Next
    
    If Cells(8, 3).Value = "" Then
        MsgBox "문제의 갯수는 숫자를 입력하셔야 합니다."
        Exit Sub
    End If
    
    If IsNumeric(Cells(8, 3)) = False Then
        MsgBox "문제의 갯수는 숫자를 입력하셔야 합니다."
        Exit Sub
    ElseIf Cells(8, 3).Value > num_of_test Then
        MsgBox "문제의 갯수가 해당 범위의 단어보다 많습니다. " & Format(num_of_test) & "보다 적게 적어주세요"
        Exit Sub
    End If
    
    num_of_test = Cells(8, 3).Value
    
    '이제 여기서 문제를 미리 만들어버리자
    '문제들이랑 그에 해당하는 답의 사이즈를 동적으로 바꾸어주어야 한다.
    ReDim problems(1 To num_of_test)
    ReDim answers(1 To num_of_test, 1 To 4)
    ReDim problem_pos(1 To num_of_test, 1 To 2)
    
    index_i = 1
    Do While True
        Randomize
        '일단 시트값부터 랜덤값을 주자
        Dim value_of_distance As Integer
        Dim temp_sheet_value As Integer
        
        value_of_distance = range__(2) - range__(1)
        temp_sheet_value = Int(Rnd * (value_of_distance + 1)) + range__(1)
        
        '그 다음 셀 값을 주자
        Dim temp_cell_value As Integer
        temp_cell_value = Int(Rnd * (Days(temp_sheet_value).Cells(3, 4).Value)) + 1
        
        checking = True
        For i = 1 To index_i - 1
            If (temp_sheet_value = problem_pos(i, 1)) And (temp_cell_value = problem_pos(i, 2)) Then
                checking = False
                Exit For
            End If
        Next
        If checking Then
            problem_pos(index_i, 1) = temp_sheet_value
            problem_pos(index_i, 2) = temp_cell_value
            
            problems(index_i) = Days(temp_sheet_value).Cells(5 + temp_cell_value, 3).Value
            
            For i = 1 To 4
                answers(index_i, i) = Days(temp_sheet_value).Cells(5 + temp_cell_value, 3 + i).Value
            Next
            
            index_i = index_i + 1
        End If
        
        
        'index_i가 num_of_test가 되면 loop에서 탈출
        If index_i = num_of_test + 1 Then
            Exit Do
        End If
    Loop
    
    
    test_start = True
    Cells(8, 6).Value = "다시 시작"
    
    index_i = 1
    
    Cells(2, 10).Value = Format(index_i) & " / " & Format(num_of_test)
    
    Call change_problem
End Sub

'테스트 중에 시작버튼이 눌릴 경우 초기화를 해주는 매크로
Sub init()
    '혹시 사용자가 day1개라도 삭제를 했을 때를 대비하는 부분
    num_of_days = ActiveWorkbook.Worksheets.Count
    num_of_days = num_of_days - 5
    
    If num_of_days <> Sheets("단어입력").Cells(1, 16).Value Then
        Sheets("단어입력").Cells(1, 16).Value = num_of_days
    End If
    For i = 1 To num_of_days
        Dim checking_name As String
        checking_name = Format(i)
        checking_name = "Day" & checking_name
        If Sheets(5 + i).Name <> checking_name Then
            Sheets(5 + i).Name = checking_name
        End If
    Next
    
    
    '존재하는 day를 보고 범위를 만들어주어야함
    '일단 초기상황에는 단어장이 아무것도 입력이 안 되어 있을테니 0개인 상황부터 생각해야 한다
    If num_of_days = 0 Then
        '0개일 때에는 범위에 0만 나오게 해야함
        With Cells(5, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='단어입력'!$Q$1:$Q$1"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        Cells(5, 3).Value = 0
        
        With Cells(5, 5).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='단어입력'!$Q$1:$Q$1"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        Cells(5, 5).Value = 0
        
        '0개이니 test를 시작할 수 없는 상황이다
        can_start = False
        
    '0개가 아닐 경우
    Else
        Dim range_day As String
        range_day = Format(num_of_days + 1)
        range_day = "='단어입력'!$Q$2:$Q$" & range_day
        
        With Cells(5, 3).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=range_day
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        '친절하게 초기값까지 설정해주자
        Cells(5, 3).Value = 1
        
        With Cells(5, 5).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=range_day
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        '친절하게 초기값까지 설정해주자
        Cells(5, 5).Value = 1
        
        '시작할 수 있는 상황이다
        can_start = True
        
    End If

    'worksheet도 여기서 초기화를 하자
    ReDim Days(num_of_days)
    
    For i = 1 To num_of_days
        Dim days_name As String
        days_name = Format(i)
        days_name = "Day" & days_name
        
        Set Days(i) = Worksheets(days_name)
    Next
    check_up = True
    check_right = True
    Cells(5, 9).Value = ""
    Cells(10, 9).Value = ""
    Cells(12, 9).Value = ""
    Cells(14, 9).Value = ""
    Cells(16, 9).Value = ""
    Cells(18, 9).Value = ""
    
    test_start = False
    Cells(8, 6).Value = "시작"
    
    Cells(2, 10).Value = ""
    Cells(10, 2).Value = ""
    Cells(8, 3).Value = ""
    틀린개수 = 0
End Sub


'셀 움직여질 때 생기는 이벤트
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    selecting = False
    
    If test_start Then
        '다시 시작 버튼이 눌린 상황
        If (Target.Row = 8 And Target.Column = 6) And (num_of_test + 1 = index_i) Then
            MsgBox "초기화 됩니다."
            Call init
            Exit Sub
        ElseIf Target.Row = 8 And Target.Column = 6 Then
            반환값 = MsgBox("시험을 멈추고 초기화하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2, "시험 초기화")
   
            Select Case 반환값
                Case Is = vbYes
                    MsgBox "초기화 됩니다."
                    Call init
                    Exit Sub
                Case Is = vbNo
                    MsgBox "초기화하지 않습니다."
                End Select
        End If
        
        '시험이 끝난 상황
        If num_of_test + 1 = index_i Then
            MsgBox " 시험이 끝났습니다. " & Format(num_of_test) & "개의 문제들 중 " & Format(틀린개수) & "개가 틀렸습니다." & vbCrLf & "다시 시작 버튼을 눌러주시고 초기화를 해주세요."
        '시험중
        Else
            '버튼이 선택되어을 때만 선택되게 해야 함
            '1번 선택
            If Target.Row = 10 And Target.Column = 9 Then
                selecting = True
                selected_number = 1
            '2번 선택
            ElseIf Target.Row = 12 And Target.Column = 9 Then
                selecting = True
                selected_number = 2
            '3번 선택
            ElseIf Target.Row = 14 And Target.Column = 9 Then
                selecting = True
                selected_number = 3
            '4번 선택
            ElseIf Target.Row = 16 And Target.Column = 9 Then
                selecting = True
                selected_number = 4
            '5번 선택
            ElseIf Target.Row = 18 And Target.Column = 9 Then
                selecting = True
                selected_number = 5
            End If
            
            If selecting Then
                If selected_number = answer_number Then
                    If Days(problem_pos(index_i, 1)).Cells(5 + problem_pos(index_i, 2), 8).Value > 0 Then
                        Days(problem_pos(index_i, 1)).Cells(5 + problem_pos(index_i, 2), 8).Value = Days(problem_pos(index_i, 1)).Cells(5 + problem_pos(index_i, 2), 8).Value - 1
                    End If
                Else
                    Days(problem_pos(index_i, 1)).Cells(5 + problem_pos(index_i, 2), 8).Value = Days(problem_pos(index_i, 1)).Cells(5 + problem_pos(index_i, 2), 8).Value + 1
                    틀린개수 = 틀린개수 + 1
                End If
                    
                Cells(2, 10).Value = Format(index_i) & " / " & Format(num_of_test)
                    
                index_i = index_i + 1
                Call change_problem
                selecting = False
            End If
        End If
        '1,16에 고정
        Cells(1, 16).Select
        Selection.ClearContents
     
    '시험이 시작되기 전
    Else
        '시작 버튼이 눌린 상황
        If Target.Row = 8 And Target.Column = 6 Then
            Call start
            Cells(1, 16).Select
        End If

        
        '위로 갈수 없고, 오른쪽으로 갈수 있음
        If Target.Row = 5 And Target.Column = 3 Then
            check_up = False
            check_right = True
            where_it_is(1) = 5
            where_it_is(2) = 3
        
        '위로 갈수 없고, 오른쪽으로 갈수 없음
        ElseIf Target.Row = 5 And Target.Column = 5 Then
            check_up = False
            check_right = False
            where_it_is(1) = 5
            where_it_is(2) = 5
        '위로 갈수 있고, 오른쪽으로 갈수 없음
        ElseIf Target.Row = 8 And Target.Column = 3 Then
            check_up = True
            check_right = False
            where_it_is(1) = 8
            where_it_is(2) = 3
            '여기서 최대값을 알려주자
            If Cells(5, 3).Value > Cells(5, 5).Value Then
                Cells(10, 2).Value = "시작점이 종점보다 큽니다."
            
            Else
                range__(1) = Cells(5, 3).Value
                range__(2) = Cells(5, 5).Value
    
                '해당 범위에서 나올 수 있는 문제의 최대 갯수
                num_of_test = 0
                If num_of_days <> 0 Then
                    For i = range__(1) To range__(2)
                        num_of_test = Days(i).Cells(3, 4).Value + num_of_test
                    Next
                End If
            
                Cells(10, 2).Value = Format(num_of_test) & "보다 적게 적어주세요"
            End If


        Else
            '옆으로 갈 때
            If Target.Row = 5 And (Target.Column = 4) Then

                If where_it_is(1) = 5 And where_it_is(2) = 3 Then
                    Cells(5, 5).Select
                    where_it_is(1) = 5
                    where_it_is(2) = 5
                Else
                    Cells(5, 3).Select
                    where_it_is(1) = 5
                    where_it_is(2) = 3
                End If
            
            '상하로 움직일 때
            ElseIf (Target.Row = 7) And (Target.Column > 2 Or Target.Column < 6) Then
                If where_it_is(1) = 8 And where_it_is(2) = 3 Then
                    Cells(5, 3).Select
                    where_it_is(1) = 5
                    where_it_is(2) = 3
                Else
                    Cells(8, 3).Select
                    where_it_is(1) = 8
                    where_it_is(2) = 3
                End If
            
            '나머지 상황은 가까운 곳으로 가게
            ElseIf Target.Row < 7 Then
                If Target.Column < 5 Then
                    Cells(5, 3).Select
                Else
                    Cells(5, 5).Select
                End If
            Else
                Cells(8, 3).Select
            End If
        End If
        
        
    End If
    
    
    

End Sub


'sheet가 처음 열릴 때 이벤트
Private Sub Worksheet_Activate()
    'test는 시작된 상황이 아니니
    Call init
End Sub

'여기서 문제를 바꾸도록 하자 (문제를 출력)
Sub change_problem()
    If index_i > num_of_test Then
        Exit Sub
    End If

    Dim wrong_answer(1 To 4) As String
    Dim wrong_answer_pos(1 To 4, 1 To 2) As Integer
    Dim answer_str As String
    answer_str = ""
    wrong_answer(1) = "1. "
    wrong_answer(2) = "2. "
    wrong_answer(3) = "3. "
    wrong_answer(4) = "4. "
    '여기서 바꿔도 되는지를 체크하는게 아닌 바꾸는 것만 생각하자
    '일단 정답 번호를 랜덤으로 배치시켜야 한다
    Randomize
    answer_number = Int(Rnd * 4) + 1
    answer_str = Format(answer_number) & ". "
    
    answer_str = answer_str & answers(index_i, 1)
    For i = 2 To 4
        If answers(index_i, i) <> "x" Then
            answer_str = answer_str & ", " & answers(index_i, i)
        End If
    Next
    
    '오답 생성
    For i = 1 To 4
        If i <> answer_number Then
            Do While True
                checking = True
                Randomize
                Dim temp_sheet_pos As Integer
                temp_sheet_pos = Int(Rnd * num_of_days) + 1
                
                Randomize
                Dim temp_cell_pos As Integer
                temp_cell_pos = Int(Rnd * (Days(temp_sheet_pos).Cells(3, 4).Value)) + 1
                
                For j = 1 To i - 1
                    If (temp_sheet_pos = wrong_answer_pos(j, 1)) And (temp_cell_pos = wrong_answer_pos(j, 2)) Then
                        checking = False
                        Exit For
                    End If
                Next
                
                If (temp_sheet_pos = problem_pos(index_i, 1)) And (temp_cell_pos = problem_pos(index_i, 2)) Then
                    checking = False
                End If
                
                If checking Then
                    wrong_answer(i) = wrong_answer(i) & Days(temp_sheet_pos).Cells(5 + temp_cell_pos, 3 + 1).Value
                    For j = 2 To 4
                        If Days(temp_sheet_pos).Cells(5 + temp_cell_pos, 3 + j).Value <> "x" Then
                            wrong_answer(i) = wrong_answer(i) & ", " & Days(temp_sheet_pos).Cells(5 + temp_cell_pos, 3 + j).Value
                        End If
                    Next
                    Exit Do
                End If
            Loop
        End If
    Next
    
    '전부 출력
    Cells(5, 9).Value = problems(index_i)
    For i = 1 To 4
        If i = answer_number Then
            Cells(8 + 2 * i, 9).Value = answer_str
        Else
            Cells(8 + 2 * i, 9).Value = wrong_answer(i)
        End If
    Next
    
    Cells(18, 9).Value = "5. 잘 모르는 단어"
End Sub

'엑셀 파일이 열릴 때 실행될 매크로
Sub excel_open()
'test는 시작된 상황이 아니니
    Call init
End Sub
