# SAP Procedger 
	Sub SAP_PRO()
			Dim mPath As String, tLR As Double
			Set mWS = ThisWorkbook.Worksheets("macro_sheet")
			bCode = mWS.Range("$E$6").Value
			Posting_FromDate = mWS.Range("$E$7").Value
			Posting_ToDate = mWS.Range("$E$8").Value

			mPath = ThisWorkbook.Path
			Dim SapGuiAuto As Object
			Dim sApplication As Object
			Dim Connection As Object
			Dim session As Object

			If bCode = 4700 Then
				vVarient = ""
				lLayout = ""
				bName = "Heartland"
				
			ElseIf bCode = 5300 Then
				vVarient = "MDAOM 5300 V2"
				lLayout = "/CA 2LAYOUT"
				bName = "Canada"

			ElseIf bCode = 4300 Then
				vVarient = ""
				lLayout = ""
				bName = " BSNA"

			ElseIf bCode = 4800 Then
				vVarient = ""
				lLayout = ""
				bName = "Southwest"
				
			ElseIf bCode = 4900 Then
				vVarient = ""
				lLayout = ""
				bName = "Abarta"

			ElseIf bCode = 5200 Then
				vVarient = ""
				lLayout = ""
				bName = " Liberty"

			ElseIf bCode = 5300 Then
				vVarient = ""
				lLayout = ""
				bName = "Canada"

			End If

			Set SapGuiAuto = GetObject("SAPGUI")
			Set sApplication = SapGuiAuto.GetScriptingEngine
			Set Connection = sApplication.Children(0)
			Set session = Connection.Children(0)	
		' Bring SAP to front
		
			Set objShell = CreateObject("wscript.shell")
			objShell.AppActivate (CStr(session.ActiveWindow.Text))
			
		'SAP Script Start from Hare
	
	
	
	
	
	
			For Each wn In Application.Windows
				On Error Resume Next
				fName = 0
				fName = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Find("Worksheet in Basis", wn.Caption), 0)
				On Error GoTo 0
				If fName > 0 Then
					Workbooks(wn.Caption).Activate
					ActiveWindow.WindowState = xlMaximized
					'Application.Wait (Now() + TimeValue("00:05:00"))
					ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\Raw Data" & "\" & bCode & ".xlsx"
					'Application.Wait (Now() + TimeValue("00:00:10"))
				
				End If
			Next wn	
	
	End Sub


#Log SAP in with the data

	Sub SapConn()

		Dim Appl As Object
		Dim Connection As Object
		Dim session As Object
		Dim WshShell As Object
		Dim SapGui As Object

		'Of course change for your file directory
		Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
		Set WshShell = CreateObject("WScript.Shell")

		Do Until WshShell.AppActivate("SAP Logon ")
			Application.Wait Now + TimeValue("0:00:01")
		Loop

		Set WshShell = Nothing

		Set SapGui = GetObject("SAPGUI")
		Set Appl = SapGui.GetScriptingEngine
		Set Connection = Appl.Openconnection("paste name of module", _
			True)
		Set session = Connection.Children(0)

		'if You need to pass username and password
		session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "900"
		session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "user"
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "password"
		session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"

		If session.Children.Count > 1 Then

			answer = MsgBox("You've got opened SAP already," & _
		"please leave and try again", vbOKOnly, "Opened SAP")

			session.findById("wnd[1]/usr/radMULTI_LOGON_OPT3").Select
			session.findById("wnd[1]/usr/radMULTI_LOGON_OPT3").SetFocus
			session.findById("wnd[1]/tbar[0]/btn[0]").press

			Exit Sub

		End If

		session.findById("wnd[0]").maximize
		session.findById("wnd[0]").sendVKey 0 'ENTER

		'and there goes your code in SAP

	End Sub
# T-Code -S_ALR_87009994- Credit Limit 

	Sub credit_limit_Sap()
		Dim mPath As String, tLR As Double
		Set mWS = ThisWorkbook.Worksheets("macro_sheet")
		Set SapGuiAuto = GetObject("SAPGUI")
		Set sApplication = SapGuiAuto.GetScriptingEngine
		Set Connection = sApplication.Children(0)
		Set session = Connection.Children(0)
		' Bring SAP to front
		Set objShell = CreateObject("wscript.shell")
		objShell.AppActivate (CStr(session.ActiveWindow.Text))	
		session.findById("wnd[0]/tbar[0]/okcd").Text = "/NS_ALR_87009994"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/ctxtKKBER-LOW").Text = "C100"
		session.findById("wnd[0]/usr/ctxtKKBER-LOW").SetFocus
		session.findById("wnd[0]/usr/ctxtKKBER-LOW").caretPosition = 4
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
		session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
		session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ThisWorkbook.Path & "\" '"C:\Users\nchennapay\Downloads\New folder\"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Credit_Limit.XLS"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
		session.findById("wnd[1]/tbar[0]/btn[11]").press	
	End Sub

# T-Code- NFBL5N -file download

	Sub SAP_NFBL5N()
			Dim mPath As String, tLR As Double	
			Set mWS = ThisWorkbook.Worksheets("macro_sheet")
			bCode = mWS.Range("$E$6").Value
			Posting_FromDate = mWS.Range("$E$7").Value
			Posting_ToDate = mWS.Range("$E$8").Value	
			mPath = ThisWorkbook.Path
			Dim SapGuiAuto As Object
			Dim sApplication As Object
			Dim Connection As Object
			Dim session As Object	
			Set SapGuiAuto = GetObject("SAPGUI")
			Set sApplication = SapGuiAuto.GetScriptingEngine
			Set Connection = sApplication.Children(0)
			Set session = Connection.Children(0)
			' Bring SAP to front	
			Set objShell = CreateObject("wscript.shell")
			objShell.AppActivate (CStr(session.ActiveWindow.Text))
			'SAP Script Start from Hare
			session.findById("wnd[0]/tbar[0]/okcd").Text = "/NFBL5N"
			session.findById("wnd[0]").sendVKey 0
			session.findById("wnd[0]/usr/chkX_SHBV").Selected = False
			session.findById("wnd[0]/usr/chkX_NORM").Selected = True
			session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").Text = ""
			session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").Text = "c100"
			session.findById("wnd[0]/usr/ctxtPA_STIDA").Text = Format(Posting_FromDate, "dd.mm.yyyy") '"07.06.2023"
			session.findById("wnd[0]/usr/ctxtPA_VARI").Text = "/GL & ARREAR"
			session.findById("wnd[0]/usr/ctxtPA_VARI").SetFocus
			session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 12
			session.findById("wnd[0]/tbar[1]/btn[8]").press
			session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
			session.findById("wnd[1]/usr/radRB_OTHERS").SetFocus
			session.findById("wnd[1]/usr/radRB_OTHERS").Select
			session.findById("wnd[1]/usr/cmbG_LISTBOX").Key = "08"
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			'session.findById("wnd[1]/tbar[0]/btn[0]").press	
			For Each wn In Application.Windows
					On Error Resume Next
					fName = 0
					fName = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Find("Worksheet in ALVXXL01 (1)", wn.Caption), 0)
					On Error GoTo 0
					If fName > 0 Then
						Workbooks(wn.Caption).Activate
						ActiveWindow.WindowState = xlMaximized
						'Application.Wait (Now() + TimeValue("00:05:00"))
						ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\Raw_Data_Credit Limit Report - BI.xlsx"	
					End If
			Next wn	
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				Workbooks("Worksheet in ALVXXL01 (1)").Activate
				ActiveWorkbook.Close False	
	End Sub
	
# AR - PD - T-Code - nS_ALR_87012178
	sub sap_S_ALR_87012178()
			Set SapGuiAuto = GetObject("SAPGUI")
			Set sApplication = SapGuiAuto.GetScriptingEngine
			Set Connection = sApplication.Children(0)
			Set session = Connection.Children(0)
			' Bring SAP to front
			Set objShell = CreateObject("wscript.shell")
			objShell.AppActivate (CStr(session.ActiveWindow.Text))	
		  	session.findById("wnd[0]/tbar[0]/okcd").Text = "/nS_ALR_87012178"
		        session.findById("wnd[0]").sendVKey 0
		        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").Text = "c100"
		        session.findById("wnd[0]/usr/ctxtDD_STIDA").Text = Format(Posting_FromDate, "dd.mm.yyyy") '"24.05.2023"
		        session.findById("wnd[0]/usr/txtMONAT").Text = "16"
		        session.findById("wnd[0]/usr/ctxtAKONTS-HIGH").SetFocus
		        session.findById("wnd[0]/usr/ctxtAKONTS-HIGH").caretPosition = 0
		        session.findById("wnd[0]/usr/btn%_AKONTS_%_APP_%-VALU_PUSH").press
		        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]").Text = "510050"
		        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,1]").Text = "510000"
		        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,2]").Text = "513100"
		        'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,2]").SetFocus
		        'session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,2]").caretPosition = 6
		        session.findById("wnd[1]/tbar[0]/btn[8]").press
		        session.findById("wnd[0]/usr/txtVERDICHT").Text = "6"
		        session.findById("wnd[0]/usr/ctxtXBUKRDAT").Text = "2"
		        session.findById("wnd[0]/usr/txtRASTBIS2").Text = "30"
		        session.findById("wnd[0]/usr/txtRASTBIS3").Text = "60"
		        session.findById("wnd[0]/usr/txtRASTBIS4").Text = "90"
		        session.findById("wnd[0]/usr/txtRASTBIS5").Text = "120"
		        session.findById("wnd[0]/usr/txtRASTBIS5").SetFocus
		        session.findById("wnd[0]/usr/txtRASTBIS5").caretPosition = 0
		        session.findById("wnd[0]/tbar[1]/btn[8]").press
		        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
		        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
		        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
		        session.findById("wnd[1]/tbar[0]/btn[0]").press
		        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ThisWorkbook.Path & "\" '"C:\Users\nchennapay\Downloads\AR\"
		        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "" & i & "_AR & PD " & Format(Date, "dd.mm.yyyy") & ".XLS"
		        session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
		        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 33
		        session.findById("wnd[1]/tbar[0]/btn[0]").press
		        'session.findById("wnd[0]/tbar[0]/btn[3]").press
   	End Sub


# T- Code:-me23n(extract data)

		Sub sap_log_script(ByVal po As String, ByVal itm As String, ByVal j As Integer, ByVal a As Integer)
			
			
			Set WWS = ThisWorkbook.Worksheets("Working")
			WWS.Visible = True
			WWS.Cells.Clear
			
			On Error Resume Next
			Set SapGuiAuto = GetObject("SAPGUISERVER")
			Set SapApplication = SapGuiAuto.GetScriptingEngine
			Set SapConnection = SapApplication.Children(0)
			On Error GoTo 0
			
			If IsObject(SapConnection) = False Then
				MsgBox "Unable to establish a connection with SAP. Please try again!"
				Exit Sub
			End If
			
			If Not IsObject(session) Then
				Set session = SapConnection.Children(0)
			End If
			
			If IsObject(WScript) Then
				WScript.ConnectObject session, "on"
				WScript.ConnectObject SapApplication, "on"
			End If
			
			
			
			session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme23n"
			session.findById("wnd[0]").sendVKey 0
			session.findById("wnd[0]/tbar[1]/btn[17]").press
			session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = po '"4512090088"
			session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").caretPosition = 10
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			
			session.findById("wnd[0]").sendVKey 27
		'Application.Wait (Now() + TimeValue("00:00:10"))
		'On Error Resume Next
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").Columns.elementAt(1).Selected = True
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/btnEDITFILTER").press
			session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = itm '"20"
			session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 2
			session.findById("wnd[1]").sendVKey 0
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").SetFocus
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").Key = "   1"
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").SetFocus
			
			
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = 1
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = 1
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17").Select
			
			
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").currentCellRow = -1
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").selectColumn "BEWTK"
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").pressToolbarButton "&MB_FILTER"
			session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "RE-L"
			session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			
			
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").currentCellColumn = "BELNR"
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").selectedRows = "0"
			session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1106/subSUB2:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").clickCurrentCell
			
			Invoice_date = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").Text 'setFocus
			Reference = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").Text 'SetFocus
			
			Company_Code = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-BUTXT").Text 'SetFocus
			
			Dim Code() As String
			Code = Split(Company_Code, " ")
			Company_Code = Code(0)
			
			txt = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT").Text 'SetFocus
			
			amt = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").Text 'setFocus
			
			ThisWorkbook.Sheets("PO Template").Range("c" & a).Value = Invoice_date
			ThisWorkbook.Sheets("PO Template").Range("d" & a).Value = Reference
			ThisWorkbook.Sheets("PO Template").Range("e" & a).Value = Company_Code
			ThisWorkbook.Sheets("PO Template").Range("f" & a).Value = txt
			ThisWorkbook.Sheets("PO Template").Range("g" & a).Value = amt
			
		'Call sap_log_script_miro(po, itm, Invoice_date, Reference, Company_Code, txt, amt)
		End Sub
# T-Code:-miro ( Fill all line item)
		Sub sap_log_script_miro(ByVal po As String, ByVal Invoice_date As String, ByVal Reference As String, ByVal Company_Code As String, ByVal txt As String, ByVal lp_int As Integer)
			
			
			Set WWS = ThisWorkbook.Worksheets("Working")
			WWS.Visible = True
		'WWS.Cells.Clear
			
			On Error Resume Next
			Set SapGuiAuto = GetObject("SAPGUISERVER")
			Set SapApplication = SapGuiAuto.GetScriptingEngine
			Set SapConnection = SapApplication.Children(0)
			On Error GoTo 0
			
			If IsObject(SapConnection) = False Then
				MsgBox "Unable to establish a connection with SAP. Please try again!"
				Exit Sub
			End If
			
			If Not IsObject(session) Then
				Set session = SapConnection.Children(0)
			End If
			
			If IsObject(WScript) Then
				WScript.ConnectObject session, "on"
				WScript.ConnectObject SapApplication, "on"
			End If
			
			
			session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmiro"
			session.findById("wnd[0]").sendVKey 0
			On Error Resume Next
			session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").Text = Company_Code '"5001"
			session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").caretPosition = 4
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			
			
		'session.findById("wnd[0]").resizeWorkingPane 132, 12, False
			session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmiro"
			session.findById("wnd[0]").sendVKey 0
			session.findById("wnd[0]/usr/cmbRM08M-VORGANG").Key = "2"
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").SetFocus
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").caretPosition = 0
			session.findById("wnd[0]/mbar/menu[1]/menu[0]").Select
			session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").Text = Company_Code '"5001"
			session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").caretPosition = 4
			session.findById("wnd[1]/tbar[0]/btn[0]").press
			session.findById("wnd[0]/usr/cmbRM08M-VORGANG").SetFocus
			
			
			
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").Text = Invoice_date ' "04/27/2022"
			session.findById("wnd[0]").sendVKey 0
		'Application.Wait (Now() + TimeValue("00:00:10"))
			session.findById("wnd[0]").sendVKey 0
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").Text = Reference '"TCPS211378"
			session.findById("wnd[0]").sendVKey 0
			
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT").Text = txt & "K1" '"TCPS211378K1"
			
			iLr = ThisWorkbook.Sheets("PO Template").Cells(Rows.Count, 11).End(xlUp).Row
			lp = iLr Mod 10
			Fd = Application.WorksheetFunction.Floor(iLr, 10) / 10
			temp = 1
			Scroll = 0
			session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/btnRM08M-XMSEL").press
			i = 0
			For i = 0 To Fd
				
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,0]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value ' "4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,0]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value '"20"
				
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,1]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,1]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value '"40"
				
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,2]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value ' "4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,2]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value '"50"
				
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,3]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,3]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value '"140"
				
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,4]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,4]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value ' "150"
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,5]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,5]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value '"190"
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,6]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,6]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value ' "200"
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,7]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,7]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value ' "250"
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,8]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value '"4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,8]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value ' "260"
				temp = temp + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,9]").Text = ThisWorkbook.Sheets("PO Template").Range("k" & temp).Value ' "4513853752"
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,9]").Text = ThisWorkbook.Sheets("PO Template").Range("l" & temp).Value ' "270"
		'session.findById("wnd[0]").sendVKey 82
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,10]").SetFocus
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,10]").caretPosition = 0
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll '1
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '2
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '3
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '4
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll '5
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '6
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '7
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '8
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '9
				Scroll = Scroll + 1
				session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = Scroll  '10
		'If i >= 0 Then session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = temp '1
				
				Next i
				
				session.findById("wnd[1]/tbar[0]/btn[8]").press
				
		'session.findById("wnd[0]").sendVKey 0
		'session.findById("wnd[0]").sendVKey 0
				
				session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
				session.findById("wnd[0]").sendVKey 0
				session.findById("wnd[0]").sendVKey 0
				k = 0
				temp = 0
				lr_S = lp_int
				For k = 0 To iLr
					session.findById("wnd[0]").sendVKey 0
					session.findById("wnd[0]").sendVKey 0
					Received = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WEMNG[38," & temp & "]").Text 'SetFocus
					Settled = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-REMNG[39," & temp & "]").Text 'SetFocus
					Final_Q = Int(Settled) - Int(Received)
					Amount_1 = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[35," & temp & "]").Text 'SetFocus
					If Amount_1 = "" Then Amount_1 = 0
					Final_Amount = (CDec(Amount_1) / Settled) * Final_Q
					ThisWorkbook.Sheets("PO Template").Activate
					
					itm = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELP[31," & temp & "]").Text 'setFocus
					session.findById("wnd[0]").sendVKey 0
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[35," & temp & "]").Text = "" & Format(Final_Amount, "0.00") & "" '"24"
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-MENGE[33," & temp & "]").Text = "" & Final_Q & "" '"2"
					temp = temp + 1
					If (k + 1) Mod 6 = 0 Then
						session.findById("wnd[0]").sendVKey 82
						temp = 0
					End If
					Next k
					
					balance = session.findById("wnd[0]/usr/txtRM08M-DIFFERENZ").Text 'setFocus
					
					If Int(balance) = 0 Then
						ThisWorkbook.Sheets("PO Template").Range("H" & lr_S).Value = "This po already processed."
						Exit Sub
					End If
					
					session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").Text = Left(balance, Len(balance) - 1) '"123"
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").SetFocus
		'session.findById("wnd[0]").sendVKey 0
		'session.findById("wnd[0]").sendVKey 0
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").caretPosition = 3
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI").Select
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/cmbINVFO-BLART").Key = "K1"
					session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/cmbINVFO-BLART").SetFocus
					check_bal = session.findById("wnd[0]/usr/txtRM08M-DIFFERENZ").Text 'setFocus
					If Int(check_bal) <> 0 Then
						ThisWorkbook.Sheets("PO Template").Range("H" & lr_S).Value = "Getting Error:- balance is not able to 0 "
						Exit Sub
					End If
					
					session.findById("wnd[0]/tbar[1]/btn[43]").press
					session.findById("wnd[1]").Close
					
					session.findById("wnd[0]/tbar[1]/btn[13]").press
					error_m = session.findById("wnd[1]/tbar[0]/btn[17]").Text 'press
					session.findById("wnd[1]/tbar[0]/btn[12]").press
					If Int(error_m) > 0 Then
		'ThisWorkbook.Sheets("Status").Range("A" & lr_S + 1).Value = po  'grengauh     Lawson@@@02
						session.findById("wnd[1]").Close
						ThisWorkbook.Sheets("PO Template").Range("H" & lr_S).Value = "Getting Error- " & error_m & " "
						Exit Sub
					Else
						session.findById("wnd[1]").Close
						session.findById("wnd[0]").sendVKey 11
						ThisWorkbook.Sheets("PO Template").Range("H" & lr_S).Value = "Processed...."
					End If
					
					
		End Sub
				
		
		
		


