# SmartView_Function
Excel菜单实现的SmartView上传+取数功能

Public Type amyess
    name As String
    server As String
    application As String
    database As String
End Type

Const UserN = "e0c00e3"
Const Pin = "Ltu1818."

Sub RetrieveSelection()

On Error Resume Next
Dim ret As Long
    ret = EssVSetSheetOption(Empty, 11, False)
    ret = EssVSetSheetOption(exmpty, 6, False)
    ret = EssVSetSheetOption(Empty, 7, False)
    
    ret = EssVSetSheetOption(exmpty, 6, False)
    ret = EssVSetSheetOption(Empty, 7, False)
    ret = EssVSetSheetOption(Empty, 11, True)
 
 Dim EssName(1 To 3) As amyess
'    EssName(1).name = "GN_"
'    EssName(1).server = "Pcnnt10040a"
'    EssName(1).application = "GLCN"
'    EssName(1).database = "GLCN"
        EssName(1).name = "US"
        EssName(1).server = "ihoess"
        EssName(1).application = "iPlanV2"
        EssName(1).database = "iPlanV2"
    
    EssName(2).name = "SAP"
    EssName(2).server = "Pcnnt10040a"
    EssName(2).application = "SAPGLCN"
    EssName(2).database = "SAPGLCN"
    
    EssName(3).name = "FLA"
    EssName(3).server = "Pcnnt10040a"
    EssName(3).application = "FLASH"
    EssName(3).database = "Main"

 application.ScreenUpdating = False
 application.DisplayAlerts = False

 Dim i As Integer
'     UserN = InputBox("Input UserN", UserN, UserN)
'     Pin = InputBox("Input Pin", UserN, Pin)
 YesNo = MsgBox("获取数据？" & vbCrLf & "Yes:SAP  No:US/GNCN  Cancel:Flash", vbYesNoCancel, "？")
    Select Case YesNo
       Case vbYes
           i = 2
       Case vbNo
           i = 1
       Case vbCancel
           i = 3
    End Select
    
 Call EssVConnect(Empty, UserN, Pin, EssName(i).server, EssName(i).application, EssName(i).database)
 Call EssVRetrieve(Empty, Selection, 1)
 Call EssVDisconnect(Empty)
 Beep

End Sub


Sub UploadSelection()

 application.ScreenUpdating = False
 application.DisplayAlerts = False

Dim ret As Long
    ret = EssVSetSheetOption(Empty, 11, False)
    ret = EssVSetSheetOption(exmpty, 6, False)
    ret = EssVSetSheetOption(Empty, 7, False)
 
Dim Calc As String, i As Integer
'        Pin = InputBox("Input Pin", UserN)
 YesNo = MsgBox("上传数据？" & vbCrLf & "Yes:Buget  No:3YP  Cancel:US/IRR", vbYesNoCancel, "？")
 Select Case YesNo
    Case vbYes
        Calc = "BGRMBS"         'Budget
'        Calc = "CPBGTO0"         'Budget0
'        Calc = "BK7RMBS"         '3YrPlanBK7, Open to OPS FIN, by store: Annual Plan by month
'        Calc = "CPBG-BK1"         '3YrPlanBK1/2/3/4/5/6,Copy version, fill difference version
        Call EssVConnect(Empty, UserN, Pin, "Pcnnt10040a", "SAPGLCN", "SAPGLCN")
        Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
        Call EssVSendData(Empty, Selection)
        Call EssVCalculate(Empty, Calc, Empty)
    
    Case vbNo
        Calc = "Cal5YrPS"       '3YRPlan， 3YRPlan尚未使用
'        Calc = "CP5Y-BK1"         '3YrPlanBK1/2/3/4/5/6Copy version,CP5Y-BK7尚未使用
        Call EssVConnect(Empty, UserN, Pin, "Pcnnt10040a", "SAPGLCN", "SAPGLCN")
        Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
        Call EssVSendData(Empty, Selection)
        Call EssVCalculate(Empty, Calc, True)
    Case vbCancel
'################################################################IRR 计算器
    '        Calc = "Cal_IRR"
    '        Call EssVConnect(Empty, UserN, Pin, "Pcnnt10040a", "SAPGLCN", "SAPGLCN")
    '        Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
    '        Call EssVSendData(Empty, Selection)
    '        i = EssVCalculate(Empty, Calc, True)
 '################################################################AOP 计算器
'        Call EssVConnect(Empty, UserN, Pin, "ihoess", "iPlanV2", "iPlanV2")
'        Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
'        Call EssVSendData(Empty, Selection)
'            YesNo = MsgBox("类型" & vbCrLf & "Yes:FCST Base  No:AOP PL  Cancel:AOP BS", vbYesNoCancel, "？")
'            If YesNo = vbYes Then
'                Call EssVCalculate(Empty, "CNTYPB", Empty)
'            ElseIf YesNo = vbNo Then
'                Call EssVCalculate(Empty, "CNNYWP", Empty)   'CNNYWP   CN3YWP
'            Else
'                Call EssVCalculate(Empty, "CNTYPB", Empty)   'CNTYPB
'                Call EssVCalculate(Empty, "CNNYWP", Empty)   'CNNYWP
'            End If
 '################################################################3YP 计算器
        Call EssVConnect(Empty, UserN, Pin, "ihoess", "iPlanV2", "iPlanV2")
        Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
        Call EssVSendData(Empty, Selection)
'            YesNo = MsgBox("类型" & vbCrLf & "Yes:FCST Base  No:3YP PL  Cancel:3YP BS", vbYesNoCancel, "？")
'            If YesNo = vbYes Then
                Call EssVCalculate(Empty, "CN3YWP", Empty)    '"CNcl3YWP"
'            ElseIf YesNo = vbNo Then
'                Call EssVCalculate(Empty, "CN3YWP", Empty)
'            Else
'                Call EssVCalculate(Empty, "CN3YWP", Empty)
'            End If
 End Select
    
' Call EssVConnect(Empty, UserN, Pin, "Pcnnt10040a", "SAPGLCN", "SAPGLCN")
' Call EssVRetrieve(Empty, Selection, 3)         '3 is Lockdata?
' Call EssVSendData(Empty, Selection)
' i = EssVCalculate(Empty, Calc, True)
'            If i = 0 Then
'            MsgBox ("Calculation complete.")
'            Else
'            MsgBox ("Calculation failed.")
'            End If

 Call EssVDisconnect(Empty)
 Beep

End Sub


Sub ZISelection()

On Error Resume Next
Dim ret As Long
    ret = EssVSetSheetOption(Empty, 11, False)      'Disable Formular preservation
    ret = EssVSetSheetOption(exmpty, 6, False)      'Disable suppress #Missing setting
    ret = EssVSetSheetOption(Empty, 7, False)       'Disable suppress Zeroes setting
    
    ret = EssVSetSheetOption(Empty, 11, False)
    ret = EssVSetSheetOption(exmpty, 6, True)
    ret = EssVSetSheetOption(Empty, 7, True)
    
 Dim EssName(1 To 3) As amyess
    EssName(1).name = "GN_"
    EssName(1).server = "Pcnnt10040a"
    EssName(1).application = "GLCN"
    EssName(1).database = "GLCN"
    
    EssName(2).name = "SAP"
    EssName(2).server = "Pcnnt10040a"
    EssName(2).application = "SAPGLCN"
    EssName(2).database = "SAPGLCN"
    
    EssName(3).name = "US"
    EssName(3).server = "ihoess"
    EssName(3).application = "iPlanV2"
    EssName(3).database = "iPlanV2"

 application.ScreenUpdating = False
 application.DisplayAlerts = False

 Dim i As Integer
'     UserN = InputBox("Input UserN", UserN, UserN)
'     Pin = InputBox("Input Pin", UserN, Pin)
 YesNo = MsgBox("扩展数据？" & vbCrLf & "Yes:SAP  No:GNCN  Cancel:US", vbYesNoCancel, "？")
    Select Case YesNo
       Case vbYes
           i = 2
       Case vbNo
           i = 1
       Case vbCancel
           i = 3
    End Select
    
 Call EssVConnect(Empty, UserN, Pin, EssName(i).server, EssName(i).application, EssName(i).database)
 Call EssVZoomIn(Empty, Empty, Selection, 3, False)
 Call EssVDisconnect(Empty)
 Beep

End Sub


Sub MultSheet_SAPRetrieve()
    Dim SapRange As String
    SapRange = Selection.Address
'    SapRange = Right(SapRange, Len(SapRange) - InStrRev(SapRange, "!"))
    Dim i, n As Long
        i = ActiveWindow.SelectedSheets.Count
    Dim SapSheet() As String
    ReDim SapSheet(i)
    For n = 1 To i
        SapSheet(n) = ActiveWindow.SelectedSheets(n).name
    Next
    For n = 1 To i
        If Left(Sheets(SapSheet(n)).Range(SapRange).Cells(13, 1).Value, 3) <> "ACC" Then GoTo amybreak
            Sheets(SapSheet(n)).Select
            Call EssVConnect(Empty, UserN, Pin, "pcnnt10040a", "SAPGLCN", "SAPGLCN")
            Call EssVRetrieve(Empty, Sheets(SapSheet(n)).Range(SapRange), 1)
            Call EssVDisconnect(Empty)
amybreak:
    Next
    
    
End Sub
