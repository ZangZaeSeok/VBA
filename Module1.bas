Attribute VB_Name = "Module1"
Sub ǥ_�����_���()
Attribute ǥ_�����_���.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ǥ_�����_��� ��ũ��
'

'
    ActiveWindow.SmallScroll Down:=42
    Range("C5:H45").Select
    Range("H45").Activate
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$5:$H$45"), , xlYes).Name = _
        "ǥ1"
    Range("ǥ1[#All]").Select
    ActiveSheet.ListObjects("ǥ1").TableStyle = "TableStyleMedium15"
    ActiveWindow.SmallScroll Down:=30
End Sub
Sub ���������̰����Ѱ�()
Attribute ���������̰����Ѱ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���������̰����Ѱ� ��ũ��
' ����Ե� �����ϴ�!

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='�ܾ��Է�'!$A$5:$A$14"
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
