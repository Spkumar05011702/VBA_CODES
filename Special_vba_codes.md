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
