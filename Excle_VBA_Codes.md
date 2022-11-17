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


