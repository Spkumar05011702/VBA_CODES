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