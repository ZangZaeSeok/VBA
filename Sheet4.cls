VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim num_of_days As Integer
Dim Days() As Worksheet

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
    
    ReDim Days(num_of_days)
    
    For i = 1 To num_of_days
        Dim days_name As String
        days_name = Format(i)
        days_name = "Day" & days_name
        
        Set Days(i) = Worksheets(days_name)
    Next
    
    '범위 지정
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
        
        
        '시작할 수 있는 상황이다
        no_days = False
        
    End If
    
    '기존의 시험지를 날려야 함
    For i = 1 To Cells(1, 13).Value
        Cells(9, 1).EntireRow.Delete Shift:=xlUp
    Next
    
    Cells(1, 13).Value = 0
End Sub

Sub output()
    Dim start_point As Integer, end_point As Integer
    Dim index_i As Integer
    
    '이제 범위를 받아온다
    start_point = Cells(5, 4).Value
    end_point = Cells(5, 6).Value
    
    If start_point > end_point Then
        MsgBox "종점보다 시작점이 더 큽니다."
        Cells(5, 4).Value = 1
        Exit Sub
    End If
    
    index_i = 0
    
        
    '일단 표를 초기화한다 (범위를 바꾸어서 구할 수 있으니
    Call init
    

    
    
    For i = start_point To end_point
        For j = 1 To Days(i).Cells(3, 4).Value
            If Days(i).Cells(5 + j, 8).Value > 0 Then
               index_i = index_i + 1
               
               Cells(8 + index_i, 3).Value = i
               Cells(8 + index_i, 4).Value = Days(i).Cells(5 + j, 3).Value
               Cells(8 + index_i, 5).Value = Days(i).Cells(5 + j, 4).Value
               Cells(8 + index_i, 6).Value = Days(i).Cells(5 + j, 5).Value
               Cells(8 + index_i, 7).Value = Days(i).Cells(5 + j, 6).Value
               Cells(8 + index_i, 8).Value = Days(i).Cells(5 + j, 7).Value
               Cells(8 + index_i, 9).Value = Days(i).Cells(5 + j, 8).Value
            End If
        Next
    Next
    
    Cells(1, 13).Value = index_i
End Sub



'sheet가 처음 열릴 때 이벤트
Private Sub Worksheet_Activate()
    '친절하게 초기값까지 설정해주자
    Cells(5, 4).Value = 1
    '친절하게 초기값까지 설정해주자
    Cells(5, 6).Value = 1
    Call init
End Sub

'엑셀 파일이 열릴 때 실행될 매크로
Sub excel_open()
    '친절하게 초기값까지 설정해주자
    Cells(5, 4).Value = 1
    '친절하게 초기값까지 설정해주자
    Cells(5, 6).Value = 1
    Call init
End Sub

