# send_bulk_sal_slip
send lots of salary slips in excel
Sub Outlook夾檔發信()


 
 '計算有幾個需發送信件欄位
 Dim toRange As Range
 Set toRange = Worksheets("發信指令").Range("A:A")
 answer = Application.WorksheetFunction.CountA(toRange) + 2
 Worksheets("發信指令").Range("E2") = answer
 


    'Dim OutApp As Object
    'Dim OutMail As Object
      

    For i = 7 To answer
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    

    'On Error Resume Next
    With OutMail
        .To = Sheets("發信指令").Cells(i, 2).Text 'mail，對應B2儲存Range("B2").Text格的位置或可直接輸入"dmmiao@simenvi.com.tw"
        .Cc = ""
        .Bcc = ""
        .Subject = Sheets("發信指令").Range("A2").Text  '主旨
        .Body = Sheets("發信指令").Range("B2").Text '內容
        '.Attachments.Add ActiveWorkbook.FullName  若無附件請選擇此語法
         .Attachments.Add (Sheets("發信指令").Cells(i, 3).Text) '附件檔位置
        .Send   'or use .Display
    End With
    'On Error GoTo 0

    'Set OutMail = Nothing
    'Set OutApp = Nothing
    
    Next i
    
    
End Sub
