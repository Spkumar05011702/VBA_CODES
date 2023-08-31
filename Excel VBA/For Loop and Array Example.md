# Array example
 	Sub excel_format(f_name, s_code, c_mnth, descript)
				Application.DisplayAlerts = False
			'f_name = "I:\GP\3MEDICARE\Gaurav\ramesh\JRS-P2P\excels\Correction Entries for the JRS Check Deposits.xlsx"
			's_code = "489"
			'c_mnth = "8"
				Dim src As Worksheet
				Dim tar As Worksheet
				
				Set c = GetObject(f_name)
				c.Sheets.Add after:=c.Sheets(c.Sheets.Count)
				Set src = c.Sheets(1)
				Set tar = c.Sheets(c.Sheets.Count)
				c.Application.Visible = True
				c.Parent.Windows(c.Parent.Windows.Count()).Visible = True
				c.Activate
				tar.Select
				tar.Cells.Select
				Selection.NumberFormat = "@"
				
				endrow = src.UsedRange.Row - 1 + src.UsedRange.Rows.Count
				If Trim(src.Range("b" & CStr(endrow)) & src.Range("c" & CStr(endrow)) & src.Range("d" & CStr(endrow)) & src.Range("j" & CStr(endrow)) & src.Range("k" & CStr(endrow)) & src.Range("l" & CStr(endrow))) = "" Then
					endrow = endrow - 1
					While Trim(src.Range("b" & CStr(endrow)) & src.Range("c" & CStr(endrow)) & src.Range("d" & CStr(endrow)) & src.Range("j" & CStr(endrow)) & src.Range("k" & CStr(endrow)) & src.Range("l" & CStr(endrow))) = ""
						endrow = endrow - 1
					Wend
				End If
				
				Dim ar()
				ar = Array("", "journal number", "date", "charge au", "charge account", "chg sub", "amount", "sign", "vendor", "keyrec", "reason code", "credit au", "credit acct", "crd sub", "type", "company code", "source code", "acct month", "desc", "cip", "budget", "imprint", "recycle", "vname1", "vname2", "deal", "store number", "distributor name", "store  state", "amount by gl code")
			'ar_index = 1
				col_index = 1
				For I = 1 To 29
					For j = 1 To 29
						If ar(I) = LCase(src.Cells(1, j)) Then
							Exit For
						End If
						Next j
						If j > 29 Then
							Debug.Print "error"
						Else
							If I = 2 Then
								tar.Range("b2:b" & CStr(endrow)).Select
								Selection.NumberFormat = "m/d/yyyy"
							ElseIf I = 6 Then
								tar.Range("f2:f" & CStr(endrow + 1)).Select
								Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
							End If
							
							If I = 18 Then
								For k = 1 To endrow
									tar.Cells(k, I) = src.Cells(k, j)
									Next k
								Else
									For k = 1 To endrow
										If k = 1 Then
											tar.Cells(k, I) = src.Cells(k, j)
										Else
											tar.Cells(k, I) = Replace(src.Cells(k, j), " ", "")
										End If
										Next k
									End If
								End If
								
								Next I
								
			'    If Trim(tar.Range("f" & CStr(endrow))) <> "" And _
			'        ((Trim(tar.Range("b" & CStr(endrow))) = "" And Trim(tar.Range("c" & CStr(endrow))) = "") Or _
			'        (Trim(tar.Range("k" & CStr(endrow))) = "" And Trim(tar.Range("l" & CStr(endrow))) = "")) Then
			'
			'        endrow = endrow - 1
			'
			'    End If
								
			'    tar.Columns("R:R").Select
			'    Selection.ColumnWidth = 30
								
								tar.Range("f" & CStr(endrow + 1)).Formula = "=sum(f2:f" & CStr(endrow) & ")"
								tar.Range("a" & CStr(endrow + 1)) = endrow - 1
								
								
			'Procedure for constraints and conditions checking
								descript = ""
								
			'    Dim empt()
			'    empt = Array(True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
								
								For I = 2 To endrow
									
									If Trim(tar.Range("a" & CStr(I))) = "" Then
										tar.Range("a" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-a" & I
									End If
									If Len(Trim(tar.Range("a" & CStr(I)))) > 30 Then
										tar.Range("a" & CStr(I)).Select
			'Orange
										Selection.Interior.Color = 49407
										descript = descript & "," & "a" & I
									End If
									If Trim(tar.Range("b" & CStr(I))) = "" Then
										tar.Range("b" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-b" & I
									End If
									If Trim(tar.Range("c" & CStr(I))) = "" Or Len(Trim(tar.Range("c" & CStr(I)))) > 6 Then
										tar.Range("c" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-c" & I
									End If
									If Trim(tar.Range("d" & CStr(I))) = "" Or Len(Trim(tar.Range("d" & CStr(I)))) > 6 Then
										tar.Range("d" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-d" & I
									End If
									If Trim(tar.Range("f" & CStr(I))) = "" Or tar.Range("f" & CStr(I)) < 0 Then
										tar.Range("f" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-f" & I
									End If
									If Trim(tar.Range("e" & CStr(I))) = "" Then
										tar.Range("e" & CStr(I)) = "0000"
									Else
										tar.Range("e" & CStr(I)) = Format(CStr(tar.Range("e" & CStr(I))), "0000")
									End If
									
									If Trim(tar.Range("m" & CStr(I))) = "" Then
										tar.Range("m" & CStr(I)) = "0000"
									Else
										tar.Range("m" & CStr(I)) = Format(CStr(tar.Range("m" & CStr(I))), "0000")
									End If
									
									tar.Range("g" & CStr(I)) = ""
									
									If Trim(tar.Range("h" & CStr(I))) = "" Then
										tar.Range("h" & CStr(I)) = "0"
									End If
									
									If Trim(tar.Range("I" & CStr(I))) = "" Then
										tar.Range("I" & CStr(I)) = "0"
									End If
									
									If Len(tar.Range("I" & CStr(I))) > 6 Then
										tar.Range("I" & CStr(I)).Select
										Selection.Interior.Color = 49407
										descript = descript & "," & "I" & I
									End If
									
									If Trim(tar.Range("j" & CStr(I))) = "" Then
										tar.Range("j" & CStr(I)) = "0"
									End If
									
									If Trim(tar.Range("k" & CStr(I))) = "" Or Len(Trim(tar.Range("k" & CStr(I)))) > 6 Then
										tar.Range("k" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-k" & I
									End If
									
									
									If Trim(tar.Range("l" & CStr(I))) = "" Or Len(Trim(tar.Range("l" & CStr(I)))) > 6 Then
										tar.Range("l" & CStr(I)).Select
			'Red
										Selection.Interior.Color = 255
										descript = descript & "," & "DONOT-L" & I
									End If
									
									
									tar.Range("N" & CStr(I)) = ""
									tar.Range("O" & CStr(I)) = ""
									tar.Range("P" & CStr(I)) = s_code
									tar.Range("Q" & CStr(I)) = c_mnth
									tar.Range("S" & CStr(I)) = ""
									tar.Range("T" & CStr(I)) = ""
									tar.Range("V" & CStr(I)) = ""
									tar.Range("Y" & CStr(I)) = ""
			'        For em = 1 To 25
			'            If Trim(tar.Cells(i, em)) <> "" Then
			'                empt(em) = False
			'            End If
			'        Next em
									Next I
									
			'      For emp = 1 To 25
			'        If empt(emp) Then
			'            tar.Cells(1, emp).Select
			'            Selection.EntireColumn.Hidden = True
			'        End If
			'      Next emp
									
			' Call printing
			'i = 1
									While c.Sheets.Count <> 1
										Sheets(1).Select
										ActiveWindow.SelectedSheets.Delete
										
									Wend
									
									Cells.Select
									Cells.EntireColumn.AutoFit
									
									c.Save
									
									c.Close
			'   c.Application.Quit
			'formatt = "ok"
									Set c = Nothing
									Application.DisplayAlerts = True
									Debug.Print descript
								End Sub
								Sub ttttest()
									Application.DisplayAlerts = False
									ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
			'ThisWorkbook.Sheets.Add(after:=ActiveSheet).Name = "Aggregated"
									While ThisWorkbook.Sheets.Count <> 1
			'ThisWorkbook.Sheets(1).Select
			'ActiveWindow.SelectedSheets.Delete
										ThisWorkbook.Sheets(1).Delete
									Wend
									c.Sheets(1).Range("a2:Y" & CStr(endrow1)).Copy Destination:=ThisWorkbook.Sheets(1).Range("a" & CStr(f_row))
									Application.DisplayAlerts = True
								End Sub
								Function formatt(f_name)
									Dim src As Sheets
									Dim tar As Sheets
									
									Set c = GetObject(f_name)
									c.Sheets.Add after:=c.Sheets(c.Sheets.Count)
									Set src = c.Sheets(1)
									Set tar = c.Sheets(c.Sheets.Count)
									
									formatt = src.Range("a1")
									
									c.Application.Visible = True
									c.Parent.Windows(c.Parent.Windows.Count()).Visible = True
									c.Save
									c.Close
			'   c.Application.Quit
			'formatt = "ok"
									Set c = Nothing
								End Function
			'Sub excel_printing(fintech)
								Sub excel_printing()
			'f_name = ThisWorkbook.Path & "\temp.xlsm"
			'fintech = False
									Application.DisplayAlerts = False
			'Set c = GetObject(f_name)
			'   c.Application.Visible = True
			'  c.Parent.Windows(c.Parent.Windows.Count()).Visible = True
									Set src = ThisWorkbook.Sheets(1)
									endrow = src.UsedRange.Row - 1 + src.UsedRange.Rows.Count
									
									Application.ActivePrinter = "CutePDF Writer on CPW4:"
			'Application.ActivePrinter = "CutePDF Writer"
									Cells.Select
									Cells.EntireColumn.AutoFit
			'Code to hide the columns which having complete blank
			'---------------------------------------------------------------
									endrow1 = endrow
									If Trim(src.Range("b" & CStr(endrow1)) & src.Range("c" & CStr(endrow1)) & src.Range("d" & CStr(endrow1)) & src.Range("j" & CStr(endrow1)) & src.Range("k" & CStr(endrow1)) & src.Range("l" & CStr(endrow1))) = "" Then
										endrow1 = endrow1 - 1
										While Trim(src.Range("b" & CStr(endrow1)) & src.Range("c" & CStr(endrow1)) & src.Range("d" & CStr(endrow1)) & src.Range("j" & CStr(endrow1)) & src.Range("k" & CStr(endrow1)) & src.Range("l" & CStr(endrow1))) = ""
											endrow1 = endrow1 - 1
										Wend
									End If
									Dim empt()
									empt = Array(True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
									
									For I = 2 To endrow1
										For em = 1 To 25
											If Trim(src.Cells(I, em)) <> "" Then
												empt(em) = False
											End If
											Next em
											Next I
											
											For emp = 1 To 25
												If empt(emp) Then
													src.Cells(1, emp).Select
													Selection.EntireColumn.Hidden = True
												End If
												Next emp
			'---------------------------------------------------------------
												
			'    If Trim(src.Range("b" & CStr(endrow)) & src.Range("c" & CStr(endrow)) & src.Range("d" & CStr(endrow)) & src.Range("j" & CStr(endrow)) & src.Range("k" & CStr(endrow)) & src.Range("l" & CStr(endrow))) = "" Then
			'        endrow = endrow - 1
			'        While Trim(src.Range("b" & CStr(endrow)) & src.Range("c" & CStr(endrow)) & src.Range("d" & CStr(endrow)) & src.Range("j" & CStr(endrow)) & src.Range("k" & CStr(endrow)) & src.Range("l" & CStr(endrow))) = ""
			'            endrow = endrow - 1
			'        Wend
			'    End If
												
												
												
			'If fintech = 0 Then
			'src.Columns("R:R").Select
			'Selection.ColumnWidth = 30
			'End If
												
												Range("A1").Select
												Range(Selection, Selection.End(xlToRight)).Select
												Range(Selection, Selection.End(xlDown)).Select
												Selection.Copy
												Application.CutCopyMode = False
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.PrintTitleRows = ""
													.PrintTitleColumns = ""
												End With
												Application.PrintCommunication = True
												ActiveSheet.PageSetup.PrintArea = ""
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.LeftHeader = ""
													.CenterHeader = ""
													.RightHeader = ""
													.LeftFooter = ""
													.CenterFooter = ""
													.RightFooter = ""
													.LeftMargin = Application.InchesToPoints(0.25)
													.RightMargin = Application.InchesToPoints(0.25)
													.TopMargin = Application.InchesToPoints(0.75)
													.BottomMargin = Application.InchesToPoints(0.75)
													.HeaderMargin = Application.InchesToPoints(0.3)
													.FooterMargin = Application.InchesToPoints(0.3)
													.PrintHeadings = False
													.PrintGridlines = False
													.PrintComments = xlPrintNoComments
													.PrintQuality = 600
													.CenterHorizontally = False
													.CenterVertically = False
													.Orientation = xlPortrait
													.Draft = False
													.PaperSize = xlPaperLetter
													.FirstPageNumber = xlAutomatic
													.Order = xlDownThenOver
													.BlackAndWhite = False
													.Zoom = 100
													.PrintErrors = xlPrintErrorsDisplayed
													.OddAndEvenPagesHeaderFooter = False
													.DifferentFirstPageHeaderFooter = False
													.ScaleWithDocHeaderFooter = True
													.AlignMarginsHeaderFooter = True
													.EvenPage.LeftHeader.Text = ""
													.EvenPage.CenterHeader.Text = ""
													.EvenPage.RightHeader.Text = ""
													.EvenPage.LeftFooter.Text = ""
													.EvenPage.CenterFooter.Text = ""
													.EvenPage.RightFooter.Text = ""
													.FirstPage.LeftHeader.Text = ""
													.FirstPage.CenterHeader.Text = ""
													.FirstPage.RightHeader.Text = ""
													.FirstPage.LeftFooter.Text = ""
													.FirstPage.CenterFooter.Text = ""
													.FirstPage.RightFooter.Text = ""
												End With
												Application.PrintCommunication = True
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.PrintTitleRows = ""
													.PrintTitleColumns = ""
												End With
												Application.PrintCommunication = True
												ActiveSheet.PageSetup.PrintArea = ""
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.LeftHeader = ""
													.CenterHeader = ""
													.RightHeader = ""
													.LeftFooter = ""
													.CenterFooter = ""
													.RightFooter = ""
													.LeftMargin = Application.InchesToPoints(0.25)
													.RightMargin = Application.InchesToPoints(0.25)
													.TopMargin = Application.InchesToPoints(0.75)
													.BottomMargin = Application.InchesToPoints(0.75)
													.HeaderMargin = Application.InchesToPoints(0.3)
													.FooterMargin = Application.InchesToPoints(0.3)
													.PrintHeadings = False
													.PrintGridlines = False
													.PrintComments = xlPrintNoComments
													.PrintQuality = 600
													.CenterHorizontally = False
													.CenterVertically = False
													.Orientation = xlLandscape
													.Draft = False
													.PaperSize = xlPaperLetter
													.FirstPageNumber = xlAutomatic
													.Order = xlDownThenOver
													.BlackAndWhite = False
													.Zoom = 100
													.PrintErrors = xlPrintErrorsDisplayed
													.OddAndEvenPagesHeaderFooter = False
													.DifferentFirstPageHeaderFooter = False
													.ScaleWithDocHeaderFooter = True
													.AlignMarginsHeaderFooter = True
													.EvenPage.LeftHeader.Text = ""
													.EvenPage.CenterHeader.Text = ""
													.EvenPage.RightHeader.Text = ""
													.EvenPage.LeftFooter.Text = ""
													.EvenPage.CenterFooter.Text = ""
													.EvenPage.RightFooter.Text = ""
													.FirstPage.LeftHeader.Text = ""
													.FirstPage.CenterHeader.Text = ""
													.FirstPage.RightHeader.Text = ""
													.FirstPage.LeftFooter.Text = ""
													.FirstPage.CenterFooter.Text = ""
													.FirstPage.RightFooter.Text = ""
												End With
												Application.PrintCommunication = True
												ActiveSheet.PageSetup.PrintArea = "$A$1:$Y$" & endrow
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.PrintTitleRows = "$1:$1"
													.PrintTitleColumns = ""
												End With
												Application.PrintCommunication = True
												ActiveSheet.PageSetup.PrintArea = "$A$1:$Y$" & endrow
												Application.PrintCommunication = False
												With ActiveSheet.PageSetup
													.LeftHeader = ""
													.CenterHeader = "&F"
													.RightHeader = ""
													.LeftFooter = ""
													.CenterFooter = "Page &P of &N"
													.RightFooter = ""
													.LeftMargin = Application.InchesToPoints(0.25)
													.RightMargin = Application.InchesToPoints(0.25)
													.TopMargin = Application.InchesToPoints(0.75)
													.BottomMargin = Application.InchesToPoints(0.75)
													.HeaderMargin = Application.InchesToPoints(0.3)
													.FooterMargin = Application.InchesToPoints(0.3)
													.PrintHeadings = False
													.PrintGridlines = True
													.PrintComments = xlPrintNoComments
													.PrintQuality = 600
													.CenterHorizontally = False
													.CenterVertically = False
													.Orientation = xlLandscape
													.Draft = False
													.PaperSize = xlPaperLetter
													.FirstPageNumber = xlAutomatic
													.Order = xlDownThenOver
													.BlackAndWhite = False
													.Zoom = False
													.FitToPagesWide = 1
													.FitToPagesTall = 100
													.PrintErrors = xlPrintErrorsDisplayed
													.OddAndEvenPagesHeaderFooter = False
													.DifferentFirstPageHeaderFooter = False
													.ScaleWithDocHeaderFooter = True
													.AlignMarginsHeaderFooter = True
													.EvenPage.LeftHeader.Text = ""
													.EvenPage.CenterHeader.Text = ""
													.EvenPage.RightHeader.Text = ""
													.EvenPage.LeftFooter.Text = ""
													.EvenPage.CenterFooter.Text = ""
													.EvenPage.RightFooter.Text = ""
													.FirstPage.LeftHeader.Text = ""
													.FirstPage.CenterHeader.Text = ""
													.FirstPage.RightHeader.Text = ""
													.FirstPage.LeftFooter.Text = ""
													.FirstPage.CenterFooter.Text = ""
													.FirstPage.RightFooter.Text = ""
												End With
												Application.PrintCommunication = True
												ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
												IgnorePrintAreas:=False
			'Stop
												Cells.Select
												Selection.EntireColumn.Hidden = False
												ThisWorkbook.Save
			'c.Save
			'c.Close
			'   c.Application.Quit
			'formatt = "ok"
			'Set c = Nothing
												
												
												Application.DisplayAlerts = True
											End Sub
											
											Sub template_post(fintech)
			'f_name = "I:\GP\3MEDICARE\Gaurav\ramesh\P2P\JRS-P2P\Copy of Journal.xlsx"
												f_name = "I:\GP\SES\JRS\JRS_Automation\JRS-P2P\Journal.xlsx"
			'f_name = "I:\GP\PUBLIC\Divakar\JRS\JRS_Automation\JRS-P2P\Journal.xlsx"
												Application.DisplayAlerts = False
												Set c = GetObject(f_name)
												c.Application.Visible = True
												c.Parent.Windows(c.Parent.Windows.Count()).Visible = True
												
												endrow1 = ThisWorkbook.Sheets(1).UsedRange.Row - 1 + ThisWorkbook.Sheets(1).UsedRange.Rows.Count
												c.Activate
												c.Sheets("journal").Select
												c.Sheets("journal").Cells.Select
												Selection.ClearContents
												ThisWorkbook.Sheets(1).Range("a1:Y" & CStr(endrow1)).Copy Destination:=c.Sheets("journal").Range("a1")
												
												Dim ar()
												ar = Array("", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y")
												
												For I = 1 To 25
													If Trim(Cells(2, I)) = "" Then
														Range(ar(I) & "2:" & ar(I) & CStr(ActiveSheet.Rows.Count)).Select
														Selection.ClearContents
													End If
													Next I
													
													
													c.Sheets("journal").Cells.Select
													c.Sheets("journal").Cells.EntireColumn.AutoFit
													I = 1
													While I <= c.Sheets.Count
														If c.Sheets(I).Name Like "*" & ThisWorkbook.Sheets(1).Range("p2") & "*" And ThisWorkbook.Sheets(1).Range("p2") Like "*[0-9]*" Then
															c.Sheets(I).Select
															c.Sheets(I).Cells.Select
															Selection.ClearContents
															If fintech = 1 Then
																ThisWorkbook.Sheets(1).Range("a1:ac" & CStr(endrow1)).Copy Destination:=c.Sheets(I).Range("a1")
															Else
																ThisWorkbook.Sheets(1).Range("a1:Y" & CStr(endrow1)).Copy Destination:=c.Sheets(I).Range("a1")
															End If
															c.Sheets(I).Select
															c.Sheets(I).Cells.Select
															c.Sheets(I).Cells.EntireColumn.AutoFit
														End If
														I = I + 1
													Wend
													
													c.Save
													
													c.Close
			'   c.Application.Quit
			'formatt = "ok"
													Set c = Nothing
		
	 
	 End Sub
									
									
									
