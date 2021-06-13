Attribute VB_Name = "Module1"
Sub 표_만드는_방법()
Attribute 표_만드는_방법.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 표_만드는_방법 매크로
'

'
    ActiveWindow.SmallScroll Down:=42
    Range("C5:H45").Select
    Range("H45").Activate
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$5:$H$45"), , xlYes).Name = _
        "표1"
    Range("표1[#All]").Select
    ActiveSheet.ListObjects("표1").TableStyle = "TableStyleMedium15"
    ActiveWindow.SmallScroll Down:=30
End Sub
Sub 범위설정이가능한가()
Attribute 범위설정이가능한가.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 범위설정이가능한가 매크로
' 놀랍게도 가능하다!

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='단어입력'!$A$5:$A$14"
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
End Sub
