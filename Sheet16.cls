VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'뜻들을 받아서 Day를 만들 매크로
Sub AddDay()
    '제작될 DaySheet에 단어를 넣기 위해 필요한 시트
    Dim day As Worksheet
    'sheet 변수의 이름
    Dim sheet_name As String
    '현재 Day의 갯수
    Dim num_of_days As Integer
    '단어의 기준점이 저장될 변수
    Dim word_standard_point(1 To 2) As Integer
    '뜻의 기준점이 저장될 변수
    Dim meaning_standard_point(1 To 4, 1 To 2) As Integer
    '반복문에서 사용될 인덱스
    Dim index_i As Integer
    '단어가 저장될 변수
    Dim word As String
    '뜻이 저장될 변수
    Dim meaning(1 To 4) As String
    
    'day의 갯수 미리 받기
    num_of_days = Cells(1, 16).Value
    'day의 갯수는 이제 기존의 갯수+1이 됨. 때문에 셀 값을 미리 수정해야 함
    Cells(1, 16).Value = num_of_days + 1

    '새로운 시트의 이름
    sheet_name = Format(num_of_days + 1)
    sheet_name = "Day" & sheet_name
    
    '이제 받아온 정보로 새로운 day 시트를 생성함
    Sheets.Add after:=Sheets(5 + num_of_days)
    '새롭게 만들어진 시트의 이름을 수정
    Sheets(5 + num_of_days + 1).Name = sheet_name
    
    
    '새로운 시트를 이제 변수에 넣자
    Set day = Sheets(sheet_name)
    
    
    '새로운 시트 위해 먼저 크기를 설정하자
    day.Columns("C:G").ColumnWidth = 20
    'day.Cells(1, 1).Value = 10
    day.Cells(5, 3).Value = "단어"
    day.Cells(5, 4).Value = "뜻1"
    day.Cells(5, 5).Value = "뜻2"
    day.Cells(5, 6).Value = "뜻3"
    day.Cells(5, 7).Value = "뜻4"
    day.Cells(5, 8).Value = "틀린 횟수"
    
    
    '이제 새로운 시트에 입력된 단어들을 입력하자
    '인덱스 값을 0으로 초기화
    index_i = 0
    '기준점들을 초기화하자
    word_standard_point(1) = 8
    word_standard_point(2) = 4
    
    For i = 1 To 4
        meaning_standard_point(i, 1) = 8
        meaning_standard_point(i, 2) = 4 + i
    Next
    
    Do While True
        '빈 셀이 나올 때까지 반복하고 반복문에서 나감
        If IsEmpty(Cells(word_standard_point(1) + index_i, word_standard_point(2))) Then
            Exit Do
        End If
        
        '이제 여기서부터 출력하자
        '일단 받아와야 함
        '출력시킬 단어를 받음
        word = Cells(word_standard_point(1) + index_i, word_standard_point(2)).Value
        
        For i = 1 To 4
            If IsEmpty(Cells(meaning_standard_point(i, 1) + index_i, meaning_standard_point(i, 2))) Then
                meaning(i) = "x"
            Else
                meaning(i) = Cells(meaning_standard_point(i, 1) + index_i, meaning_standard_point(i, 2)).Value
            End If
        Next
        
        '일단 받아왓으니 이제 출력을 하면 될듯 하다
        '단어 출력
        day.Cells(6 + index_i, 3).Value = word
        '뜻 출력
        day.Cells(6 + index_i, 4).Value = meaning(1)
        day.Cells(6 + index_i, 5).Value = meaning(2)
        day.Cells(6 + index_i, 6).Value = meaning(3)
        day.Cells(6 + index_i, 7).Value = meaning(4)
        '틀린 횟수를 0으로 초기화
        day.Cells(6 + index_i, 8).Value = 0
        
        index_i = index_i + 1
    Loop
    
    'index_i를 단어의 갯수로 표현하고 싶다(단어의 갯수를 시트에 표현해야 하니까)
    index_i = index_i - 1
    
    
    
    '테두리를 그려서 예쁘게 만들어주기
    Set thin_border = day.Range(day.Cells(6, 3), day.Cells(6 + index_i, 8))
    
    thin_border.Borders.Weight = xlThin
    
    Set thick_border = day.Range("C5:H5")

    thick_border.Borders.Weight = xlThick
    
    '테두리 만들어졌으면 표로도 만들어주기
    sheet_name = sheet_name & "표"
    day.ListObjects.Add(xlSrcRange, day.Range(day.Cells(5, 3), day.Cells(6 + index_i, 8)), , xlYes).Name = sheet_name
    day.ListObjects(sheet_name).TableStyle = "TableStyleMedium15"

    '이제 표까지 만들어졌으니 단어장 시트에 추가적으로 있으면 좋은 정보들을 적어주자
    day.Range(day.Cells(1, 3), day.Cells(2, 8)).Borders.Weight = xlThick
    day.Range(day.Cells(1, 3), day.Cells(2, 8)).Merge
    day.Cells(1, 3).Value = "단어장 시트입니다. 수정 사항이 있으면, 양식에 맞게 수정해주세요. 단어 갯수가 줄어들거나 늘어날 경우 수정해주셔야 합니다." & vbCrLf & "(본 시트의 단어장이 필요없어지면 시트 삭제를 하시면 됩니다.)"
    day.Cells(1, 3).VerticalAlignment = xlCenter
    day.Cells(1, 3).HorizontalAlignment = xlCenter
    
    '단어의 갯수도 출력
    day.Range(day.Cells(3, 3), day.Cells(3, 4)).Borders.Weight = xlThick
    day.Cells(3, 3).Value = "단어 갯수:"
    day.Cells(3, 3).HorizontalAlignment = xlCenter
    day.Cells(3, 4).Value = index_i + 1
    day.Cells(3, 4).HorizontalAlignment = xlCenter
    
    '이제 다 옮겼으니까 입력한 단어들을 없애야 함
    Range(Cells(8, 4), Cells(8 + index_i, 8)).ClearContents
    
End Sub

'해당 sheet가 열릴 때 이벤트
Private Sub Worksheet_Activate()
    Dim num_of_days As Integer
    num_of_days = ActiveWorkbook.Worksheets.Count
    num_of_days = num_of_days - 5
    If num_of_days <> Cells(1, 16).Value Then
        Cells(1, 16).Value = num_of_days
    End If
    For i = 1 To num_of_days
        Dim checking_name As String
        checking_name = Format(i)
        checking_name = "Day" & checking_name
        If Sheets(5 + i).Name <> checking_name Then
            Sheets(5 + i).Name = checking_name
        End If
    Next
    
    
End Sub

'엑셀 파일이 열릴 때 실행될 매크로
Sub excel_open()
    '혹시 사용자가 day1개라도 삭제를 했을 때를 대비하는 부분
    Dim num_of_days As Integer
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
End Sub


