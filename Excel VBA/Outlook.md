# mail body with table
  	Sub CreateEmailWithFormattedTableAndSignature()
        	Dim OutApp As Object
        	Dim OutMail As Object
        	Dim MailBody As String
        	Dim DataRange As Range
        	Dim RowIndex As Long
        	
        	' Set the range containing your data
        	Set DataRange = ThisWorkbook.Sheets("Sheet1").Range("A2:D" & ThisWorkbook.Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row) ' Adjust the range as needed
        	
        	' Initialize Outlook
        	Set OutApp = CreateObject("Outlook.Application")
        	Set OutMail = OutApp.CreateItem(0)
        	
        	' Build the email body using data from the range
        	MailBody = "<html><body>" & _
        	"<p>Hello,</p>" & _
        	"<p>We have multiple discrepancies with data relating to CARD and AMEX takings, please see the below.</p>" & _
        	"<p><b>Discrepancies are as follow: </b></p>" & _
        	"<table border='1' style='border-collapse: collapse;'>" & _
        	"<tr style='font-weight: bold;'>" & _
        	"<td>Card Name</td>" & _
        	"<td>Amount</td>" & _
        	"<td>Amount1</td>" & _
        	"<td>Diff</td>" & _
        	"</tr>"
        	
        	For RowIndex = 1 To DataRange.Rows.Count
        		MailBody = MailBody & "<tr>" & _
        		"<td>" & DataRange.Rows(RowIndex).Cells(1).Value & "</td>" & _
        		"<td>" & DataRange.Rows(RowIndex).Cells(2).Value & "</td>" & _
        		"<td>" & DataRange.Rows(RowIndex).Cells(3).Value & "</td>" & _
        		"<td>" & DataRange.Rows(RowIndex).Cells(4).Value & "</td>" & _
        		"</tr>"
        		Next RowIndex
        		
        		MailBody = MailBody & "</table>" & _
        		"<p>Can you please reconcile your EOD (end of day takings) against your PDQ printout for the date(s) mentioned above to identify the discrepancy.</p>" & _
        		"<p>A discrepancy could relate to one of the following:</p>" & _
        		"<ul>" & _
        		"<li>Purchase / refund put through the PDQ and not the till / OPS</li>" & _
        		"<li>Refunds relating to cancelled or returned orders not being processed correctly</li>" & _
        		"</ul>" & _
        		"<p>If the discrepancy relates to one of the above, this will now need to be corrected via OPS. If you need any help with processing the correction, please refer to OPS guidance available on the following links:</p>"
        		
        		MailBody = MailBody & "<ul>" & _
        		"<li><a href='https://onewba.sharepoint.com/:b:/r/sites/Loss2/Shared%20Documents/Banking%20%26%20Discounts/Banking%20Discrepancy%20Tips.pdf?csf=1&web=1'>Banking Discrepancy Tips</a></li>" & _
        		"<li><a href='https://onewba.sharepoint.com/:b:/r/sites/Loss2/Shared%20Documents/Banking%20%26%20Discounts/Daily%20Finance%20Banking%20Audit.pdf?csf=1&web=1'>Daily Finance Banking Audit</a></li>" & _
        		"</ul>" & _
        		"<p>If you require further assistance, please log a call via the operations helpdesk 01159 18 19 20 opt 2, 2</p>"
        		
        		MailBody = MailBody & "<p>EOD figures entered incorrectly: If this is the case, this cannot be corrected. You will need to inform us by replying to this email as we will need to look into this further.</p>" & _
        		"<p>This discrepancy will need to be identified and rectified as a matter of urgency. Please respond confirming the below within the next 7 days:</p>"
        		
        		MailBody = MailBody & "<ul>" & _
        		"<li>What the issue was</li>" & _
        		"<li>What action has been taken</li>" & _
        		"<li>What date the discrepancy was rectified</li>" & _
        		"</ul>" & _
        		"<p>If this is not rectified, the discrepancy may be written off and will potentially hit your P&L figures and loss reports. Any discrepancies that are written off are logged and monitored and may be referred for further investigation.</p>"
        		
        		MailBody = MailBody & "<p><strong>Please ensure you reply to all when responding to emails from Opticians Reconciliation</strong></p>" & _
        		"<p><br><b>Opticians Reconciliation Team</b><br>" & _
        		"<b>Walgreens Boots Alliance</b><br>" & _
        		"<b>Boots UK</b>, D90 ES10, Thane Road, Beeston, Nottingham, NG90 1BS<br>" & _
        		"<b>Member of Walgreens Boots Alliance</b><br>" & _
        		"For help with raising a shopping cart in EBP/SRM kindly refer to the below links:" & _
        		"<br>Support Office Colleagues - <font color='blue'>Support Office How To's</font><br>" & _
        		"Store Colleagues -<font color='blue'> Stores How To's</font></p>" & _
        		"</body></html>"
        		
        	' Configure the email properties
        		With OutMail
        			.To = "recipient@example.com"
        			.Subject = "Discrepancies in CARD and AMEX Takings"
        			.HTMLBody = MailBody
        			.Display ' Use .Send to send directly without displaying the email
        		End With
        		
        	' Clean up objects
        		Set OutMail = Nothing
        		Set OutApp = Nothing
	End Sub
	
