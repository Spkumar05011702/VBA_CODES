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
