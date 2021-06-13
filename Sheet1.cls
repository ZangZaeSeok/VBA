VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
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

Dim no_days As Boolean

Dim scoring_signal As Boolean
Dim rescoring_signal As Boolean

Dim before_cell(2) As Integer


'실제 뜻이 몇개인지 미리 알려주는 변수
Dim real_many() As Integer

'선택된 번호
Dim selected_number As Integer
'선택되어졌는지 확인해주는 변수
Dim selecting As Boolean

Dim 맞춘개수 As Integer

Dim checking_color As Boolean

Sub init()
    '혹시 사용자가 day1개라도 삭제를 했을 때를 대비하는 부분
    num_of_days = ActiveWorkbook.Worksheets.Count
    num_of_days = num_of_days - 5
    
    Cells(8, 4).Value = 0
    Cells(10, 3).Value = ""

    
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
        With Cells(5, 4).Validation
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
        
        Cells(5, 4).Value = 0
        
        With Cells(5, 6).Validation
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
        
        Cells(5, 6).Value = 0
        
        '0개이니 test를 시작할 수 없는 상황이다
        no_days = True
        
    '0개가 아닐 경우
    Else
        Dim range_day As String
        range_day = Format(num_of_days + 1)
        range_day = "='단어입력'!$Q$2:$Q$" & range_day
        
        With Cells(5, 4).Validation
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
        Cells(5, 4).Value = 1
        
        With Cells(5, 6).Validation
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
        Cells(5, 6).Value = 1
        
        '시작할 수 있는 상황이다
        no_days = False
        
    End If
    
    'worksheet도 여기서 초기화를 하자
    ReDim Days(num_of_days)
    
    For i = 1 To num_of_days
        Dim days_name As String
        days_name = Format(i)
        days_name = "Day" & days_name
        
        Set Days(i) = Worksheets(days_name)
    Next
    
    '기존의 시험지를 날려야 함
    Range("B14:I1048576").Clear
    Range("B14:I1048576").Interior.Color = RGB(221, 235, 249)
    
    Cells(8, 5).Value = "시작"
    
    Cells(10, 8).Value = ""
    
    Cells(5, 4).Select
    before_cell(1) = 5
    before_cell(2) = 4
    
    scoring_signal = False
    rescoring_signal = False
    can_start = True
    
End Sub

'당연히 시작되는 매크로
Sub start()
    If no_days Then
        MsgBox "단어들을 먼저 입력하고 시작해주세요"
        Exit Sub
    End If
    
    '범위가 시작이 가능한지를 먼저 체크한다.
    If Cells(5, 4).Value > Cells(5, 6).Value Then
        MsgBox "범위는 시작점이 종점보다 작아야 합니다."
        Exit Sub
    End If
    
    If Cells(8, 4).Value < 1 Then
        MsgBox "문제의 수는 0보다 많아야 합니다."
        Exit Sub
    End If
    
    range__(1) = Cells(5, 4).Value
    range__(2) = Cells(5, 6).Value
    
    '해당 범위에서 나올 수 있는 문제의 최대 갯수
    num_of_test = 0
    For i = range__(1) To range__(2)
        num_of_test = Days(i).Cells(3, 4).Value + num_of_test
    Next
    
    If Cells(8, 4).Value = "" Then
        MsgBox "문제의 갯수는 숫자를 입력하셔야 합니다."
        Exit Sub
    End If
    
    If IsNumeric(Cells(8, 4)) = False Then
        MsgBox "문제의 갯수는 숫자를 입력하셔야 합니다."
        Exit Sub
    ElseIf Cells(8, 4).Value > num_of_test Then
        MsgBox "문제의 갯수가 해당 범위의 단어보다 많습니다. " & Format(num_of_test) & "보다 적게 적어주세요"
        Exit Sub
    End If
    
    num_of_test = Cells(8, 4).Value
    
    '이제 여기서 문제를 미리 만들어버리자
    '문제들이랑 그에 해당하는 답의 사이즈를 동적으로 바꾸어주어야 한다.
    ReDim problems(1 To num_of_test)
    ReDim answers(1 To num_of_test, 1 To 4)
    ReDim problem_pos(1 To num_of_test, 1 To 2)
    ReDim real_many(1 To num_of_test)
    
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
    
    For i = 1 To num_of_test
        Cells(13 + i, 2).Value = i
        Cells(13 + i, 2).Interior.Color = RGB(242, 242, 242)
        Cells(13 + i, 3).Value = problems(i)
        Cells(13 + i, 3).Interior.Color = RGB(242, 242, 242)
        real_many(i) = 0
        For j = 1 To 4
            If answers(i, j) = "x" Then
                Cells(13 + i, 3 + j).Interior.Color = RGB(252, 220, 214)
            Else
                Cells(13 + i, 3 + j).Interior.Color = RGB(255, 255, 255)
                real_many(i) = real_many(i) + 1
            End If
        Next
        Cells(13 + i, 8).Interior.Color = RGB(242, 242, 242)
        Cells(13 + i, 9).Interior.Color = RGB(255, 255, 255)
    Next
    
    Set thick_border = Range(Cells(13, 2), Cells(13 + num_of_test, 9))

    thick_border.Borders.Weight = -4138
    
    can_start = False
    test_start = True
    Cells(8, 5).Value = "채점"
    Cells(14, 4).Select
End Sub

'채점하는 매크로
'진짜 답이 나오고 쓴 거랑 비교해서 일단 채점이 이루어짐
Sub scoring()
    '전체를 그냥 채점하면 됨
    Dim is_it_wrong As Boolean
    Dim is_it_true As Boolean
    
    Dim real_answer_str As String
    
    is_it_wrong = False
    
    For i = 1 To num_of_test
        For j = 1 To real_many(i)
            is_it_wrong = False
            For k = 1 To real_many(i)
                If Cells(13 + i, 3 + j) = answers(i, k) Then
                    is_it_wrong = True
                    Exit For
                End If
            Next
            If is_it_wrong = False Then
                Exit For
            End If
        Next
        
        '맞을 때
        If is_it_wrong Then
            Cells(13 + i, 9).Value = "o"
        '틀릴 때
        Else
            Cells(13 + i, 9).Value = "x"
        End If
        
        real_answer_str = answers(i, 1)
        For j = 2 To real_many(i)
            real_answer_str = real_answer_str & ", " & answers(i, j)
        Next
        
        Cells(13 + i, 8).Value = real_answer_str
    Next
    
    맞춘개수 = 0
    For i = 1 To num_of_test
        If Cells(13 + i, 9).Value = "o" Then
            맞춘개수 = 맞춘개수 + 1
        End If
    Next
    
    Cells(10, 8).Value = Format(맞춘개수) & " / " & Format(num_of_test)
    
    Cells(8, 5).Value = "재채점"
    'test_start=false
    '그리고 채점 시그널을 true로
    test_start = fasle
    scoring_signal = True
End Sub

'재채점하는 매크로 그냥 check 부분만 다시보고 풀자
'그리고 여기서 틀린 단어를 해당 단어의 시트의 셀에 틀렸다고 표기를 해주어야 함
Sub rescoring()
    맞춘개수 = 0
    For i = 1 To num_of_test
        If Cells(13 + i, 9).Value = "o" Then
            맞춘개수 = 맞춘개수 + 1
        End If
    Next
    
    Cells(10, 8).Value = Format(맞춘개수) & " / " & Format(num_of_test)
    
    '여기서부터 틀린횟수를 올려주어야 함(그리고 맞은거는 줄여주어야 함)
    For i = 1 To num_of_test
        '맞으면 틀린 횟수를 줄여주기
        If Cells(13 + i, 9).Value = "o" Then
            If 0 < Days(problem_pos(i, 1)).Cells(5 + problem_pos(i, 2), 8).Value Then
                Days(problem_pos(i, 1)).Cells(5 + problem_pos(i, 2), 8).Value = Days(problem_pos(i, 1)).Cells(5 + problem_pos(i, 2), 8).Value - 1
            End If
        '틀린 경우
        Else
            Days(problem_pos(i, 1)).Cells(5 + problem_pos(i, 2), 8).Value = Days(problem_pos(i, 1)).Cells(5 + problem_pos(i, 2), 8).Value + 1
        End If
    Next
    
    '마지막으로 재채점까지 끝나면 초기화 버튼으로 바뀜
    Cells(8, 5).Value = "초기화"
    scoring_signal = False
    rescoring_signal = True
End Sub
'셀 움직여질 때 생기는 이벤트
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    '시험이 시작되기 전
    If can_start Then
        '이동 범위도 제한을 주어야 함
        If Target.Row = 5 And Target.Column = 4 Then
            before_cell(1) = 5
            before_cell(2) = 4
        ElseIf Target.Row = 5 And Target.Column = 6 Then
            before_cell(1) = 5
            before_cell(2) = 6
        '여기서 최대 문제 개수를 알려주어야 함
        ElseIf Target.Row = 8 And Target.Column = 4 Then
            '여기서 최대값을 알려주자
            If Cells(5, 4).Value > Cells(5, 6).Value Then
                Cells(10, 3).Value = "시작점이 종점보다 큽니다."
            
            Else
                range__(1) = Cells(5, 4).Value
                range__(2) = Cells(5, 6).Value
    
                '해당 범위에서 나올 수 있는 문제의 최대 갯수
                num_of_test = 0
                
                If num_of_days <> 0 Then
                    For i = range__(1) To range__(2)
                        num_of_test = Days(i).Cells(3, 4).Value + num_of_test
                    Next
                End If
            
                Cells(10, 3).Value = Format(num_of_test) & "보다 적게 적어주세요"
            End If
        
            before_cell(1) = 8
            before_cell(2) = 4
        '양 옆으로 움직일 때
        ElseIf Target.Row = 5 And Target.Column = 5 Then
            If before_cell(1) = 5 And before_cell(2) = 4 Then
                Cells(5, 6).Select
            Else
                Cells(5, 4).Select
            End If
        '위 아래로 움직일 때
        ElseIf Target.Row = 7 And (Target.Column = 4 Or Target.Column = 6) Then
            If before_cell(1) = 8 Then
                Cells(5, 4).Select
            Else
                Cells(8, 4).Select
            End If
        '범위의 시작점 셀로 이동
        ElseIf Target.Row < 8 And Target.Column < 6 Then
            Cells(5, 4).Select
        '범위의 종점 셀로 이동
        ElseIf Target.Row < 8 Then
            Cells(5, 6).Select
        '문제수 셀로 이동
        Else
            Cells(8, 4).Select
        End If
        
        
        '시작 버튼이 눌릴 경우
        If Target.Row = 8 And Target.Column = 5 Then
            Call start
        End If
        
    '시험이 시작된 후
    ElseIf test_start Then
        '시험지에만 범위를 잡아야 함
        checking_color = (Target.Interior.Color <> RGB(255, 255, 255))
        checking_color = checking_color Or (Target.Row = Cells(5, 3).Row And Cells(5, 3).Column = Target.Column) Or (Target.Row = Cells(5, 5).Row And Cells(5, 5).Column = Target.Column)
        
        checking_color = checking_color Or (Target.Column = 9)
        checking_color = checking_color Or (Target.Row < 13)
        
        If checking_color Then
            If Target.Row < 15 Then
                Cells(14, 4).Select
            ElseIf Target.Row > 12 + num_of_test Then
                Cells(13 + num_of_test, 4).Select
            Else
                Cells(Target.Row, 4).Select
            End If
        End If
        
        If Target.Row = 8 And Target.Column = 5 Then
            Call scoring
        End If
        
    '시험이 끝나고 1차 채점이 이루어진 후
    ElseIf scoring_signal Then
        'check만 선택 가능하게 하기
        checking_color = (Target.Interior.Color <> RGB(255, 255, 255))
        checking_color = checking_color Or (Target.Row = Cells(5, 3).Row And Cells(5, 3).Column = Target.Column) Or (Target.Row = Cells(5, 5).Row And Cells(5, 5).Column = Target.Column)
        
        checking_color = checking_color Or (Target.Column <> 9)
        
        If checking_color Then
            If Target.Row < 15 Then
                Cells(14, 9).Select
            ElseIf Target.Row > 12 + num_of_test Then
                Cells(13 + num_of_test, 9).Select
            Else
                Cells(Target.Row, 9).Select
            End If
        End If
        
        If Target.Row = 8 And Target.Column = 5 Then
            Call rescoring
        End If
        
    '시험이 끝나고 재채점까지 끝난 후
    ElseIf rescoring_signal Then
        Cells(1, 1).Select
        Cells(1, 1).ClearContents
        If Target.Row = 8 And Target.Column = 5 Then
            Cells(1, 1).ClearContents
            Call init
        End If
    End If
    
End Sub


'sheet가 처음 열릴 때 이벤트
Private Sub Worksheet_Activate()
    Call init
End Sub

'엑셀 파일이 열릴 때 실행될 매크로
Sub excel_open()
 Call init
End Sub
