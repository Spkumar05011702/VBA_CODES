# Case Statement 

    YesNo = MsgBox("Have you logged in for respective SAP bottler..?" & vbNewLine & vbNewLine & _
    "Have you download files form capturis portal.?" & vbNewLine & vbNewLine & _
    "Do you Wish to continue?", vbYesNo, "AP Reversal.!")
    
    Select Case YesNo      
    Case vbYes
    
    
    Case vbNo 
        MsgBox "You Have Cancelled the task.!", vbExclamation, "Electric Accrual.!"    
    End Select
  

  
# Choose file in run time

      MsgBox "Please choose base " & bName & " Electric Accrual.!", vbInformation 
      FileToOpen = Application.GetOpenFilename(Title:="Open " & bName & " Electric Accrual Base Report" _
    , FileFilter:="Excel Files(*.xls*),*xls*", MultiSelect:=False)
    
      If FileToOpen <> False Then
        Set tWb = Application.Workbooks.Open(FileToOpen)
        tWName = tWb.Name
        mSht.Range("A1").Value = tWName
        Set tWs = tWb.Sheets("KSB1_YTD 2022"   
      Else
        MsgBox "No Electric Accrual. Report chosen to process, Macro terminated.!", vbExclamation
        Exit Sub
      End If
                                    
                                    Or
      
      '******************************************'## open input JE file
    
    MsgBox "Please choose " & bCode & " BTB Accrual file.!", vbInformation, "Property Tax Journal.!"
    
    FileToOpen = Application.GetOpenFilename(Title:="Open " & bCode & " BTB Accrual file.!" _
    , FileFilter:="Excel Files(*.xls*),*xls*", MultiSelect:=False)
    
    If FileToOpen <> False Then
        Set tWB = Application.Workbooks.Open(FileToOpen)
        tName = tWB.Name
        
        TextPostn = InStr(LCase(tName), LCase("BTB Accrual"))
        If TextPostn <> 0 Then
            
            fName = Left(tName, TextPostn - 1) & " BTB Accrual " & fStamp & ".xlsx"
            
            tWB.SaveAs fPath & "\" & fName
            Set jeWB = ActiveWorkbook
            
        Else
            
            MsgBox "Selected file has not key word BTB Accrual in file name, Macro terminated.!" & vbNewLine & _
            "Please Correct the file name and run macro again.!", vbExclamation, "Property Tax Journal.!"
            Exit Sub
            
        End If
        
    Else
        MsgBox "No Raw file chosen to process, Macro terminated.!", vbExclamation, "Property Tax Journal.!"
        Exit Sub
    End If
'******************************************'## start working on output JE file
      

# Calculate filter row
    ActiveSheet.Range("$E$4:$Q$" & iLR).AutoFilter Field:=6, Criteria1:=Array("E", "Tax on Invoice"), Operator:=xlFilterValues
    fltRow = Application.WorksheetFunction.Subtotal(3, t1Ws.Range("B1:B" & Tlr))'## Count of filter row
       If fltRow > 1 Then
            fltRowno = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row '## count first filter row number     
       End If

# Create a Working Sheet and access perticular sheet

            If ActiveSheet.FilterMode Then ActiveSheet.AutoFilterMode = False 
            Sheets.Add After:=Worksheets("Prior Month")
            ActiveSheet.Name = "Working" '## Create a Working Sheet 
            For Each ws In Worksheets
                If ws.Name Like "* Use Tax Review" Then 
                
                End If
            Next ws
# Like oprerator
    If LCase(Cells(5, i).Value) Like LCase("*" & mMonth & "") Or LCase(Cells(5, i).Value) Like LCase("*" & mmMonth & "") Then
    
    end if
# Nested if botller condition

                If bCode = 4300 Then
                bName = "BSNA"
                
            ElseIf bCode = 4700 Then
                bName = "Heartland"
                
            ElseIf bCode = 4800 Then
                bName = "Southwest"
                
            ElseIf bCode = 489 Then
                bName = "TCL"
                
            ElseIf bCode = 4900 Then
                bName = "Abarta"
                
            ElseIf bCode = 5200 Then
                bName = "Liberty"
                
            ElseIf bCode = "" Then
                bName = ""
                
            End If
            
            mWS.Activate
            
            For Each ws In Worksheets
                
                If ws.Name = "working" Then
                    ws.Delete
                End If
                
             Next ws
                
             **Sheets.Add(After:=ActiveSheet).Name = "working" ' Create working sheet**






# add 

Sub test()

    Cells.UnMerge
    ActiveSheet.Range("I2").Value = ActiveSheet.Range("I1").Value
    ActiveSheet.Range("L2").Value = ActiveSheet.Range("L1").Value
    Rows(1).EntireRow.Delete
    'lr = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'Rows(lr).EntireRow.Delete
    lr = ActiveSheet.Cells(Rows.Count, 12).End(xlUp).Row
    ActiveSheet.Range("a2:F2").Copy
    ActiveSheet.Range("M2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    ActiveSheet.Range("M3:R" & lr).Value = "=IF(A3=0,M2,A3)"
    Application.Calculate
    ActiveSheet.Range("M2:R" & lr).Copy
    ActiveSheet.Range("a2").PasteSpecial xlValues
    'ActiveSheet.Range("M2:R" & lr).PasteSpecial xlValues
    Columns("M:R").EntireColumn.Delete
    ActiveSheet.Range("a1").Activate
    
    
    
    
    ActiveSheet.Range("$A$1:$L$" & lr).AutoFilter Field:=2, Criteria1:= _
        "=*gesamt", Operator:=xlAnd
    fltrow = Application.WorksheetFunction.Subtotal(3, Range("B1:B" & lr))
        
    If fltrow > 1 Then
        ActiveSheet.Range("B2:B" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    
    ActiveSheet.Range("$A$1:$L$" & lr).AutoFilter Field:=3, Criteria1:= _
        "=*gesamt", Operator:=xlAnd
    fltrow = Application.WorksheetFunction.Subtotal(3, Range("C1:C" & lr))
        
    If fltrow > 1 Then
      ActiveSheet.Range("$A$2:$L$" & lr).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    
    ActiveSheet.Range("$A$1:$L$" & lr).AutoFilter Field:=4, Criteria1:= _
        "=*gesamt", Operator:=xlAnd
    fltrow = Application.WorksheetFunction.Subtotal(3, Range("D1:D" & lr))
        
    If fltrow > 1 Then
        ActiveSheet.Range("D2:D" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    
    ActiveSheet.Range("$A$1:$L$" & lr).AutoFilter Field:=5, Criteria1:= _
        "=*gesamt", Operator:=xlAnd
    fltrow = Application.WorksheetFunction.Subtotal(3, Range("E1:E" & lr))
        
    If fltrow > 1 Then
        ActiveSheet.Range("E2:E" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    
    ActiveSheet.Range("$A$1:$L$" & lr).AutoFilter Field:=6, Criteria1:= _
        "=*gesamt", Operator:=xlAnd
    fltrow = Application.WorksheetFunction.Subtotal(3, Range("F1:F" & lr))
        
    If fltrow > 1 Then
        ActiveSheet.Range("F2:F" & lr).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    ActiveSheet.Range("a1").Activate
End Sub
