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
   	end sub
