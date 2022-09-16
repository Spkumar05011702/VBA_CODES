# Case statement 

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
      
# Web Scraping login web page in internet explorer
    Sub Login_Librty()
    Dim ie As New InternetExplorer
    Dim ohtml As HTMLDocument
    Dim htmlas, htmldivs As MSHTML.IHTMLElementCollection
    Dim htmla, htmldiv As MSHTML.IHTMLElement
    
    Set mWS = ThisWorkbook.Worksheets("macro_sheet")
    url = ""' need to url
    
    YesNo = MsgBox("Have you filled in dashboard capturis id and password..?" & vbNewLine & vbNewLine & _
    "Do you Wish to continue?", vbYesNo, "AP Reversal.!")
    
    Select Case YesNo
        
    Case vbYes
        
        ie.Visible = True
        ie.Silent = True
        
        ie.navigate url
        
        While ie.readyState <> READYSTATE_COMPLETE
            DoEvents
        Wend
        
        Set ohtml = ie.document
        
        ohtml.getElementsByName("emailid").Item(0).Value = mWS.Range("e4").Value
        ohtml.getElementsByName("passwd").Item(0).Value = mWS.Range("e5").Value
        ohtml.getElementsByTagName("input").Item(2).Click ' Login page with id password
        
        Application.Wait (Now + TimeValue("0:00:10"))
        
        ohtml.getElementsByClassName("sub-menu-content-column").Item(8).getElementsByTagName("a").Item(0).Click
        Application.Wait (Now + TimeValue("0:00:10"))
        ohtml.getElementsByName("cfrmm").Item(0).Value = Format(mWS.Range("e7").Value, "MM") '"08"
        ohtml.getElementsByName("cfrdd").Item(0).Value = Format(mWS.Range("e7").Value, "DD")
        ohtml.getElementsByName("cfryy").Item(0).Value = Format(mWS.Range("e7").Value, "YYYY")
        ohtml.getElementsByName("invgrp").Item(0).Value = "6
        ohtml.getElementsByClassName("gobutton").Item(1).Click

        While ie.readyState <> READYSTATE_COMPLETE
            DoEvents
        Wend
        Application.Wait (Now + TimeValue("0:00:10"))
        
        ohtml.getElementsByClassName("nav-item").Item(4).getElementsByTagName("a").Item(0).Click
        
        Application.Wait (Now + TimeValue("0:00:10"))
        Application.SendKeys "{TAB}"
        Application.SendKeys "{TAB}"
    Case vbNo
        
        MsgBox "You Have Cancelled the task.!", vbExclamation, "Electric Accrual.!"
        
    End Select
    End Sub

# Calculate filter row
    ActiveSheet.Range("$E$4:$Q$" & iLR).AutoFilter Field:=6, Criteria1:=Array("E", "Tax on Invoice"), Operator:=xlFilterValues
    fltRow = Application.WorksheetFunction.Subtotal(3, t1Ws.Range("B1:B" & Tlr))'## Count of filter row
       If fltRow > 1 Then
            fltRowno = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row '## count first filter row number     
       End If
