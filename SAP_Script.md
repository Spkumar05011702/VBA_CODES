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
	
	
	
	End Sub
