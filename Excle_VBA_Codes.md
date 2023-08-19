# 1.Combine Multiple Excel Files into One File :Created by Surendra Kumar from spkumar1702@gmail.com
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

# 2. Basic vba program how to read/write range 
	Sub Array_test()
		Dim arr As Variant, rg As Range
		Set rg = tSht.Range("a5").CurrentRegion
		arr = rg.Value 'or arr=tSht.Range("a5").CurrentRegion.value
		Dim rowcount As Long, columncount As Long
		rowcount = UBound(arr, 1)
		columncount = UBound(arr, 2)
		Range("A1").Resize(rowcount, columncount).Value = arr
	End Sub

# 3. Removeing special charector from string
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