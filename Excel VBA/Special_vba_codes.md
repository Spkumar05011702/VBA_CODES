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
        Set tWs = tWb.Sheets("KSB1_YTD 2022")  
      Else
        MsgBox "No Electric Accrual. Report chosen to process, Macro terminated.!", vbExclamation
        Exit Sub
      End If
                                                                       
#  Choose file in run time specific condition in file name
    
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

# Calculate filter row
    ActiveSheet.Range("$E$4:$Q$" & iLR).AutoFilter Field:=6, Criteria1:=Array("E", "Tax on Invoice"), Operator:=xlFilterValues
    fltRow = Application.WorksheetFunction.Subtotal(3, t1Ws.Range("B1:B" & Tlr))'## Count of filter row
       If fltRow > 1 Then
            fltRowno = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row '## count first filter row number 
            Rows("" & fltRowno & ":" & iLR & "").SpecialCells(xlCellTypeVisible).Select
            Selection.EntireRow.Delete
       End If

# Delete filtered rows in Excel 

    ActiveSheet.Range("$A$1:$I$" & lines).SpecialCells _
        (xlCellTypeVisible).EntireRow.Delete

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
# Add sheets and delete sheet

        tWs.Activate
        For Each ws In Worksheets
        
            If ws.Name = "working" Then
                ws.Delete
            End If
        
        Next ws
        
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "working"
        Set workingWS = tWb.Worksheets("working")
			
# Specific range fill color in range

	jeWB.Range(Cells(1, 1), Cells(1, columncount + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

# Condition formating whole row specific condition in column
	With tWs1.Range(tWs1.Cells(2, 1), tWs1.Cells(lr1 + 1, 20))
        .FormatConditions.Add Type:=xlExpression, Formula1:="=IF($E2=""50"",1,0)"
        .FormatConditions(1).Font.Color = RGB(255, 0, 0)
    End With

# Removing filter any table

	With ActiveSheet.ListObjects(1)
        	If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
	End With


# Sync data from Excel to share point list

	Sub PullData()
		Set mysh = Sheets("query (1)")
		If mysh.Visible = False Then
		mysh.Visible = True
		End If
		mysh.Activate
		mysh.Range("A:CZ").Delete
		
		Dim src(0 To 1) As Variant
		spsite = "https://onewba.sharepoint.com/sites/BootsSP/" 'sharepoint url
		src(0) = spsite & "/_vti_bin"
		src(1) = "{A0260111-2EA1-4611-8CFF-8FAA7FD304DD}" ' list id from advance
		Sheet1.ListObjects.Add xlSrcExternal, src, True, xlYes, Sheet1.Range("A1")
	End Sub
	
# Combine Multiple Excel Files into One File :Created by Surendra Kumar from spkumar1702@gmail.com
	Sub ConslidateWorkbooks()
	Dim FolderPath As String
	Dim Filename As String
	Dim Sheet As Worksheet
	Application.ScreenUpdating = False
	FolderPath = Environ("userprofile") & "DesktopTest"
	Filename = Dir(FolderPath & "*.xls*")
	Do While Filename <> ""
		Workbooks.Open Filename:=FolderPath & Filename, ReadOnly:=True
		For Each Sheet In ActiveWorkbook.Sheets
			Sheet.Copy After:=ThisWorkbook.Sheets(1)
			Next Sheet
			Workbooks(Filename).Close
			Filename = Dir()
		Loop
		Application.ScreenUpdating = True
	End Sub

# Basic vba program how to read/write range 
	Sub Array_test()
		Dim arr As Variant, rg As Range
		Set rg = tSht.Range("a5").CurrentRegion
		arr = rg.Value 'or arr=tSht.Range("a5").CurrentRegion.value
		Dim rowcount As Long, columncount As Long
		rowcount = UBound(arr, 1)
		columncount = UBound(arr, 2)
		Range("A1").Resize(rowcount, columncount).Value = arr
	End Sub

# Removeing special charector from string
	Function RemoveSpecChar(sInput As String) As String
		Dim sSpecChar As String
		Dim i As Long
		sSpecChar = "\/:*?™""®<>|.&@# (_+`©~);-+=^$!,'"
		For i = 1 To Len(sSpecChar)
			sInput = Replace$(sInput, Mid$(sSpecChar, i, 1), "")
		Next i
		RemoveSpecChar = sInput
	End Function


# SQL CONNECTION String 
	DRIVER={SQL Server};Server=localhost\SQL2012;integrated security=true;Database=Company