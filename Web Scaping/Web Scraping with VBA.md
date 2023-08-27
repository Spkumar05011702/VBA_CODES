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


# Basic Webscrping code	(Selenium-https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0) 841219 Spkumar

Private edgebrowser As Selenium.EdgeDriver
Sub TC_Browsers()
	
	Set edgebrowser = New Selenium.EdgeDriver
	With Application
		.DisplayAlerts = False
		.ScreenUpdating = False
	End With
	
	ThisWorkbook.Sheets(1).Cells.Clear
	
	edgebrowser.start baseUrl:="https://autometrics.in/"
	edgebrowser.Window.Maximize
	start:
	edgebrowser.Get "https://autometrics.in/" 'https://autometrics.in/sndtransporterconsignments
	
	edgebrowser.FindElementByName("email").SendKeys "AUTO"
	edgebrowser.FindElementByName("password").SendKeys ""
	MyInput = InputBox("Enter your captch text here from browser", "Captcha")
	edgebrowser.FindElementByClass("rnc-input").SendKeys "" & MyInput & ""
	edgebrowser.FindElementById("loginButton").Click
	edgebrowser.Wait (10000)
	Dim val As Selenium.WebElement
	On Error Resume Next
	Set val = edgebrowser.FindElementByXPath("/html/body/div/div/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/div[9]")
	If Err.Number = 7 Then
		edgebrowser.Close
		GoTo start	
	End If
	edgebrowser.Actions.ClickContext(val).Perform
	edgebrowser.Wait (5000)
	
	edgebrowser.FindElementByXPath("/html/body/div/div/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div/div/div[6]/div/div/div[5]/span[2]").Click
	edgebrowser.FindElementByXPath("/html/body/div/div/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div/div/div[7]/div/div/div[2]/span[2]").Click
	
	MsgBox "Please choose export file from download folder.!", vbInformation
	FileToOpen = Application.GetOpenFilename(Title:="Open export file" _
	, FileFilter:="Excel Files(*.xls*),*xls*", MultiSelect:=False)
	
	If FileToOpen <> False Then
		edgebrowser.Close
		Set tWb = Application.Workbooks.Open(FileToOpen)
		Set tWs = tWb.Sheets(1)
		tWs.Activate
		tlr = tWs.Cells(Rows.Count, 1).End(xlUp).Row
		tCol = tWs.Cells(1, Columns.Count).End(xlToLeft).Column
		tWs.Range(Cells(1, 1), Cells(tlr, tCol)).Copy
		ThisWorkbook.Sheets(1).Range("a5").PasteSpecial xlPasteValues
		ThisWorkbook.Sheets(1).Cells.EntireColumn.AutoFit
		tWb.Close False
	Else
		MsgBox "No export file from download folder chosen to process, Macro terminated.!", vbExclamation
		
		Exit Sub
	End If
	
	
	MsgBox "Done", vbInformation
	
	
	
'edgebrowser.Get ("https://autometrics.in/sndtransporterconsignments")
	
	
End Sub

