Attribute VB_Name = "Module1"
Sub update_TL()

Dim tacdongDict As New Scripting.dictionary
Dim countDict As New Scripting.dictionary
Dim tacDong As New TacDongInfo
Dim count As Integer
Dim lastRow As Long
'Dim wkb As Workbook
Dim current As String
Dim payMent As Long
Dim duNo As Long

 Call TurnOffStuff
    
    With ThisWorkbook.Sheets("BCTD")
    lastRow = .Cells(.Rows.count, "C").End(xlUp).Row
     
    For thisRow = lastRow To 2 Step -1
        If Not tacdongDict.Exists(.Cells(thisRow, "C").Value) Then
        
        tacDong.ma_danh_gia = .Cells(thisRow, "O").Value
        tacDong.ngay_tac_dong = .Cells(thisRow, "E").Value
        tacDong.tien_hen = .Cells(thisRow, "H").Value
        .Cells(thisRow, "G") = Format(Date, "dd-mm-yyyy")
        tacDong.ngay_hen = .Cells(thisRow, "G").Value
        
        tacdongDict.Add .Cells(thisRow, "C").Value, tacDong
        End If
        
        If Not countDict.Exists(.Cells(thisRow, "C").Value) Then
             countDict.Add .Cells(thisRow, "C").Value, 1
        Else
             countDict(.Cells(thisRow, "C").Value) = countDict(.Cells(thisRow, "C").Value) + 1
        End If
        
        Set tacDong = Nothing
        
    Next thisRow
    
End With

'wkb.Close SaveChanges:=False

With ThisWorkbook.Sheets("TL")
    lastRow = .Cells(.Rows.count, "B").End(xlUp).Row
    For thisRow = 2 To lastRow
        current = .Cells(thisRow, "B").Value
        If tacdongDict.Exists(current) Then
                .Cells(thisRow, "AA").Value = tacdongDict(current).ngay_tac_dong
                .Cells(thisRow, "AB").Value = tacdongDict(current).ma_danh_gia
                .Cells(thisRow, "AC").Value = tacdongDict(current).ngay_hen
                .Cells(thisRow, "AD").Value = tacdongDict(current).tien_hen
        End If
        
        'Tinh Pos con lai
        payMent = .Cells(thisRow, "Q").Value
        duNo = .Cells(thisRow, "D").Value
        .Cells(thisRow, "AE").Value = duNo - payMent
        
        'Tinh tinh trang hs
        If payMent >= duNo Then
        .Cells(thisRow, "AF").Value = "THANH LU"
        ElseIf payMent > 0 Then
        .Cells(thisRow, "AF").Value = "GOP"
        Else
        .Cells(thisRow, "AF").Value = "CTT"
        End If
        'tinh so luot bao cao
        If countDict.Exists(current) Then
         .Cells(thisRow, "AG").Value = countDict(current)
        End If
    Next thisRow
End With
Set tacdongDict = Nothing

Call TurnOnStuff

End Sub

Sub TurnOffStuff()

    Application.Calculation = xlManual
    Application.ScreenUpdating = False

    
End Sub

Sub TurnOnStuff()

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
End Sub


