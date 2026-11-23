Attribute VB_Name = "EMS_Output"
Option Compare Database
Option Explicit

Public Function EMSPreviewCurrent()
    Dim invoiceId As Long
    invoiceId = SelectedInvoice()
    If invoiceId = 0 Then Exit Function
    EMSFillStatement invoiceId
    DoCmd.OpenReport "rptStatement", acViewPreview
End Function

Public Function EMSExportCurrent()
    Dim rs As Object
    Dim count As Long
    Dim folder As String
    If Nz(Forms!frmEvent!EventID, 0) = 0 Then Exit Function
    If Forms!frmEvent.Dirty Then Forms!frmEvent.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT InvoiceID FROM tblInvoice WHERE EventID=" & Forms!frmEvent!EventID, EMS_SNAPSHOT)
    folder = EMSFolder()
    Do While Not rs.EOF
        EMSExportInvoice CLng(rs!InvoiceID), folder
        count = count + 1
        rs.MoveNext
    Loop
    rs.Close
    If count = 0 Then
        MsgBox "Calculate the event before exporting.", vbInformation
    Else
        MsgBox count & " portrait statement(s) saved in" & vbCrLf & folder, vbInformation
    End If
End Function

Public Function EMSEmailCurrent()
    Dim rs As Object
    Dim count As Long
    If Nz(Forms!frmEvent!EventID, 0) = 0 Then Exit Function
    If Forms!frmEvent.Dirty Then Forms!frmEvent.Dirty = False
    Set rs = CurrentDb.OpenRecordset("SELECT InvoiceID FROM tblInvoice WHERE EventID=" & Forms!frmEvent!EventID, EMS_SNAPSHOT)
    Do While Not rs.EOF
        EMSEmailInvoice CLng(rs!InvoiceID)
        count = count + 1
        rs.MoveNext
    Loop
    rs.Close
    If count = 0 Then MsgBox "Calculate the event before emailing.", vbInformation
End Function

Public Sub EMSEmailInvoice(ByVal invoiceId As Long)
    On Error GoTo Failed
    Dim path As String
    Dim outlook As Object
    Dim message As Object
    Dim primary As String
    Dim copies As String
    Dim head As Object
    path = EMSExportInvoice(invoiceId, EMSFolder())
    Set head = CurrentDb.OpenRecordset("SELECT * FROM tblPrintHead", EMS_SNAPSHOT)
    PayeeMail CLng(Scalar("SELECT PayeeID FROM tblInvoice WHERE InvoiceID=" & invoiceId)), primary, copies
    On Error Resume Next
    Set outlook = GetObject(, "Outlook.Application")
    If outlook Is Nothing Then Set outlook = CreateObject("Outlook.Application")
    On Error GoTo Failed
    If outlook Is Nothing Then
        MsgBox "Outlook is not open on this computer. The PDF is here:" & vbCrLf & path, vbExclamation
        Exit Sub
    End If
    Set message = outlook.CreateItem(0)
    message.To = primary
    message.CC = copies
    message.Subject = "Levy UNM settlement - " & head!EventWhen & " - " & head!PayeeName
    message.Body = "Settlement statement attached for " & head!PayeeName & "." & vbCrLf & _
        "Invoice " & head!InvoiceNumber & vbCrLf & "Amount due " & head!AmountDue & vbCrLf & vbCrLf & _
        "Please review before sending."
    message.Attachments.Add path
    message.Display
    head.Close
    Exit Sub
Failed:
    MsgBox "The email could not be opened." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Function EMSExportInvoice(ByVal invoiceId As Long, ByVal folder As String) As String
    Dim path As String
    Dim number As String
    EMSFillStatement invoiceId
    number = CStr(Nz(Scalar("SELECT InvoiceNumber FROM tblInvoice WHERE InvoiceID=" & invoiceId), "statement"))
    path = folder & "\" & SafeFile(number) & ".pdf"
    On Error Resume Next
    DoCmd.Close acReport, "rptStatement", acSaveNo
    On Error GoTo 0
    DoCmd.OutputTo acOutputReport, "rptStatement", acFormatPDF, path, False
    EMSExportInvoice = path
End Function

Private Function SelectedInvoice() As Long
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Function
    SelectedInvoice = CLng(Nz(Forms!frmEvent!lstInvoices, 0))
    If SelectedInvoice = 0 Then SelectedInvoice = CLng(Nz(Scalar("SELECT Min(InvoiceID) FROM tblInvoice WHERE EventID=" & Nz(Forms!frmEvent!EventID, 0)), 0))
End Function

Private Sub PayeeMail(ByVal payeeId As Long, ByRef primary As String, ByRef copies As String)
    Dim rs As Object
    primary = ""
    copies = ""
    Set rs = CurrentDb.OpenRecordset("SELECT EmailAddress, IsPrimary FROM tblPayeeEmail WHERE PayeeID=" & payeeId & " ORDER BY IsPrimary DESC, EmailID", EMS_SNAPSHOT)
    Do While Not rs.EOF
        If Len(primary) = 0 Then
            primary = CStr(Nz(rs!EmailAddress, ""))
        ElseIf Len(CStr(Nz(rs!EmailAddress, ""))) > 0 Then
            If Len(copies) > 0 Then copies = copies & "; "
            copies = copies & CStr(rs!EmailAddress)
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Public Sub EMSFillStatement(ByVal invoiceId As Long)
    Dim invoice As Object
    Dim eventRow As Object
    Dim payee As Object
    Dim lines As Object
    Dim extras As Object
    Dim showAdj As Boolean
    Dim sortIndex As Long
    Dim company As String
    Dim center As String
    Dim cleaning As Currency
    Dim misc As Currency
    EMSClearPrint
    Set invoice = CurrentDb.OpenRecordset("SELECT * FROM tblInvoice WHERE InvoiceID=" & invoiceId, EMS_SNAPSHOT)
    Set eventRow = CurrentDb.OpenRecordset("SELECT * FROM tblEvent WHERE EventID=" & invoice!EventID, EMS_SNAPSHOT)
    Set payee = CurrentDb.OpenRecordset("SELECT * FROM tblPayee WHERE PayeeID=" & invoice!PayeeID, EMS_SNAPSHOT)
    Set lines = CurrentDb.OpenRecordset("SELECT * FROM tblInvoiceLine WHERE InvoiceID=" & invoiceId & " ORDER BY SortOrder, LineID", EMS_SNAPSHOT)
    showAdj = False
    sortIndex = 1
    Do While Not lines.EOF
        If CCur(Nz(lines!Adjustments, 0)) <> 0 Then showAdj = True
        InsertPrintLine sortIndex, CStr(lines!LocationName), lines
        sortIndex = sortIndex + 1
        lines.MoveNext
    Loop
    InsertTotal sortIndex, invoice
    company = Nz(Scalar("SELECT TOP 1 CompanyCode FROM tblOffice"), "241")
    center = Nz(Scalar("SELECT TOP 1 CostCenter FROM tblOffice"), "36127")
    Set extras = CurrentDb.OpenRecordset("SELECT Label, Amount, GLCode FROM tblSupplement WHERE EventID=" & invoice!EventID & " AND PayeeID=" & invoice!PayeeID & " ORDER BY SupplementID", EMS_SNAPSHOT)
    sortIndex = 1
    Do While Not extras.EOF
        InsertExtra sortIndex, CStr(extras!Label), CCur(Nz(extras!Amount, 0)), ""
        sortIndex = sortIndex + 1
        extras.MoveNext
    Loop
    cleaning = CCur(Nz(Scalar("SELECT Sum(Cleaning) FROM tblAssign WHERE EventID=" & invoice!EventID & " AND PayeeID=" & invoice!PayeeID), 0))
    If CStr(invoice!Kind) = "NPO" Then
        misc = CCur(Nz(invoice!Supplements, 0)) + CCur(Nz(invoice!LocationAdjustments, 0)) + cleaning
        InsertGl company, center, 101, "Group Commission", "611225", CCur(invoice!Commission)
        InsertGl company, center, 102, "Gratuities", "212113", CCur(invoice!Gratuity)
        InsertGl company, center, 103, "Misc. (to be reclassed)", "673000", misc
        InsertGl company, center, 104, "Cleaning Fee - Expense", "674102", -cleaning
    Else
        misc = CCur(Nz(invoice!Supplements, 0)) + CCur(Nz(invoice!LocationAdjustments, 0))
        InsertGl company, center, 101, "Food commission", "314009", CCur(invoice!FoodCommission)
        InsertGl company, center, 102, "Beverage commission", "611120", CCur(invoice!AlcoholCommission)
        InsertGl company, center, 103, "Gratuities", "212113", CCur(invoice!Gratuity)
        If CCur(Nz(invoice!CardFee, 0)) <> 0 Then InsertGl company, center, 104, "Credit card fees", "694000", -CCur(invoice!CardFee)
        If CCur(Nz(invoice!IceAmount, 0)) <> 0 Then InsertGl company, center, 105, "Ice", "411054", -CCur(invoice!IceAmount)
        If CCur(Nz(invoice!Shortage, 0)) <> 0 Then InsertGl company, center, 106, "Inventory / sales shortage", "", -CCur(invoice!Shortage)
        If CCur(Nz(invoice!OtherExpense, 0)) <> 0 Then InsertGl company, center, 107, "Other expense", "673000", -CCur(invoice!OtherExpense)
        If misc <> 0 Then InsertGl company, center, 108, "Misc. (to be reclassed)", "673000", misc
    End If
    InsertHead invoice, eventRow, payee, showAdj
    gPageFooter = CStr(payee!PayeeName) & " | " & CStr(invoice!InvoiceNumber) & " | " & CStr(Nz(eventRow!StatusText, ""))
    lines.Close
    extras.Close
    payee.Close
    eventRow.Close
    invoice.Close
End Sub

Private Sub EMSClearPrint()
    Exec "DELETE FROM tblPrintLine"
    Exec "DELETE FROM tblPrintExtra"
    Exec "DELETE FROM tblPrintHead"
End Sub

Private Sub InsertPrintLine(ByVal sortIndex As Long, ByVal locationName As String, ByVal lines As Object)
    Exec "INSERT INTO tblPrintLine (SortOrder, ServiceArea, NetSales, FoodNet, FoodCommission, AlcoholNet, AlcoholCommission, Commission, Gratuity, Adjustments, Payout) VALUES (" & _
        sortIndex & "," & SqlText(Left$(locationName, 150)) & "," & SqlText(MoneyText(CCur(lines!NetSales))) & "," & SqlText(MoneyText(CCur(lines!FoodNet))) & "," & _
        SqlText(MoneyText(CCur(lines!FoodCommission))) & "," & SqlText(MoneyText(CCur(lines!AlcoholNet))) & "," & SqlText(MoneyText(CCur(lines!AlcoholCommission))) & "," & _
        SqlText(MoneyText(CCur(lines!Commission))) & "," & SqlText(MoneyText(CCur(lines!Gratuity))) & "," & SqlText(MoneyText(CCur(lines!Adjustments))) & "," & _
        SqlText(MoneyText(CCur(lines!Payout))) & ")"
End Sub

Private Sub InsertTotal(ByVal sortIndex As Long, ByVal invoice As Object)
    Exec "INSERT INTO tblPrintLine (SortOrder, ServiceArea, NetSales, FoodNet, FoodCommission, AlcoholNet, AlcoholCommission, Commission, Gratuity, Adjustments, Payout) VALUES (" & _
        sortIndex & ",'LOCATION TOTAL'," & SqlText(MoneyText(CCur(invoice!GroupNet))) & "," & SqlText(MoneyText(CCur(invoice!FoodNet))) & "," & _
        SqlText(MoneyText(CCur(invoice!FoodCommission))) & "," & SqlText(MoneyText(CCur(invoice!AlcoholNet))) & "," & SqlText(MoneyText(CCur(invoice!AlcoholCommission))) & "," & _
        SqlText(MoneyText(CCur(invoice!Commission))) & "," & SqlText(MoneyText(CCur(invoice!Gratuity))) & "," & SqlText(MoneyText(CCur(invoice!LocationAdjustments))) & "," & _
        SqlText(MoneyText(CCur(invoice!AmountDue))) & ")"
End Sub

Private Sub InsertExtra(ByVal sortIndex As Long, ByVal label As String, ByVal amount As Currency, ByVal glText As String)
    Exec "INSERT INTO tblPrintExtra (SortOrder, Label, Amount, GLText) VALUES (" & sortIndex & "," & SqlText(Left$(label, 80)) & "," & SqlText(MoneyText(amount, True)) & "," & SqlText(glText) & ")"
End Sub

Private Sub InsertGl(ByVal company As String, ByVal center As String, ByVal sortIndex As Long, ByVal label As String, ByVal glCode As String, ByVal amount As Currency)
    Dim glText As String
    If Len(glCode) = 0 Then glText = company & " | " & center & " | GL: not mapped" Else glText = company & " | " & center & " | GL:" & glCode
    Exec "INSERT INTO tblPrintExtra (SortOrder, Label, Amount, GLText) VALUES (" & sortIndex & "," & SqlText(label) & "," & SqlText(MoneyText(amount, True)) & "," & SqlText(glText) & ")"
End Sub

Private Sub InsertHead(ByVal invoice As Object, ByVal eventRow As Object, ByVal payee As Object, ByVal showAdj As Boolean)
    Dim address As String
    address = Trim$(CStr(Nz(payee!Address1, "")) & " " & CStr(Nz(payee!Address2, "")))
    Exec "INSERT INTO tblPrintHead (Kind, PayeeName, AddressBlock, StatementDate, InvoiceNumber, AmountDue, EventWhen, EventType, Abbreviation, SAPID, BSSID, StatusText, FoodNet, AlcoholNet, GroupNet, GrossSales, Notes, PreparedBy, PreparedTitle, PreparedWhen, ReviewedBy, ReviewedTitle, ReviewedWhen, ApprovedBy, ApprovedTitle, ApprovedWhen, EventWhenShort, FoodRateLabel, AlcoholRateLabel, ShowAdjustments, Extra) VALUES (" & _
        SqlText(CStr(invoice!Kind)) & "," & SqlText(CStr(payee!PayeeName)) & "," & SqlText(Left$(address, 255)) & "," & _
        SqlText(Format$(eventRow!EventDate, "mm/dd/yyyy")) & "," & SqlText(CStr(invoice!InvoiceNumber)) & "," & SqlText(MoneyText(CCur(invoice!AmountDue))) & "," & _
        SqlText(Format$(eventRow!EventDate, "mmm d, yyyy")) & "," & SqlText(CStr(Nz(eventRow!EventType, ""))) & "," & SqlText(CStr(Nz(payee!Abbreviation, ""))) & "," & _
        SqlText(CStr(Nz(payee!SAPID, ""))) & "," & SqlText(CStr(Nz(payee!BSSID, ""))) & "," & SqlText(CStr(Nz(eventRow!StatusText, ""))) & "," & _
        SqlText(MoneyText(CCur(invoice!FoodNet))) & "," & SqlText(MoneyText(CCur(invoice!AlcoholNet))) & "," & SqlText(MoneyText(CCur(invoice!GroupNet))) & "," & SqlText(MoneyText(CCur(invoice!GrossSales))) & "," & _
        SqlText(CStr(Nz(eventRow!Notes, ""))) & "," & _
        SqlText(CStr(Nz(eventRow!PreparedBy, ""))) & "," & SqlText(CStr(Nz(eventRow!PreparedTitle, ""))) & "," & SqlText(DateText(eventRow!PreparedDate)) & "," & _
        SqlText(CStr(Nz(eventRow!ReviewedBy, ""))) & "," & SqlText(CStr(Nz(eventRow!ReviewedTitle, ""))) & "," & SqlText(DateText(eventRow!ReviewedDate)) & "," & _
        SqlText(CStr(Nz(eventRow!ApprovedBy, ""))) & "," & SqlText(CStr(Nz(eventRow!ApprovedTitle, ""))) & "," & SqlText(DateText(eventRow!ApprovedDate)) & "," & _
        SqlText(Format$(eventRow!EventDate, "mm/dd/yyyy")) & "," & SqlText(FoodRateLabel(invoice, eventRow)) & "," & SqlText(AlcoholRateLabel(invoice, eventRow)) & "," & _
        IIf(showAdj, "-1", "0") & ",'')"
End Sub

Private Function DateText(ByVal value As Variant) As String
    If IsNull(value) Then DateText = "" Else DateText = Format$(CDate(value), "mm/dd/yyyy")
End Function

Private Function FoodRateLabel(ByVal invoice As Object, ByVal eventRow As Object) As String
    Dim rate As Double
    If CStr(invoice!Kind) = "NPO" Then
        rate = RateOn(CLng(invoice!PayeeID), "DONATION_FOOD", CDate(eventRow!EventDate), 0.1)
    Else
        rate = RateOn(CLng(invoice!PayeeID), "FOOD", CDate(eventRow!EventDate), 0)
    End If
    FoodRateLabel = Format$(rate, "0%") & " COMM."
End Function

Private Function AlcoholRateLabel(ByVal invoice As Object, ByVal eventRow As Object) As String
    Dim beerRate As Double
    Dim liquorRate As Double
    Dim naRate As Double
    Dim eventDay As Date
    eventDay = CDate(eventRow!EventDate)
    If CStr(invoice!Kind) = "NPO" Then
        AlcoholRateLabel = Format$(RateOn(CLng(invoice!PayeeID), "DONATION_ALCOHOL", eventDay, 0.08), "0%") & " COMM."
        Exit Function
    End If
    naRate = RateOn(CLng(invoice!PayeeID), "NA", eventDay, 0)
    beerRate = RateOn(CLng(invoice!PayeeID), "BEER", eventDay, 0)
    liquorRate = RateOn(CLng(invoice!PayeeID), "LIQUOR", eventDay, 0)
    If Abs(naRate - beerRate) < 0.0000001 And Abs(naRate - liquorRate) < 0.0000001 Then
        AlcoholRateLabel = Format$(naRate, "0%") & " COMM."
    Else
        AlcoholRateLabel = "BEV COMM."
    End If
End Function

Private Function SafeFile(ByVal text As String) As String
    Dim bad As Variant
    Dim item As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In bad
        text = Replace(text, CStr(item), "-")
    Next item
    SafeFile = text
End Function

Public Function EMSStatementOpen()
    On Error GoTo Failed
    Dim head As Object
    Dim report As Report
    Dim index As Long
    Dim extras As Object
    Set report = Reports("rptStatement")
    Set head = CurrentDb.OpenRecordset("SELECT * FROM tblPrintHead", EMS_SNAPSHOT)
    If head.EOF Then Exit Function
    report!txtStatus = CStr(Nz(head!StatusText, ""))
    report!txtPayee = CStr(Nz(head!PayeeName, ""))
    report!txtAddress = CStr(Nz(head!AddressBlock, ""))
    report!txtDate = CStr(Nz(head!StatementDate, ""))
    report!txtInv = CStr(Nz(head!InvoiceNumber, ""))
    report!txtDue = CStr(Nz(head!AmountDue, ""))
    report!txtEventDate = CStr(Nz(head!EventWhen, ""))
    report!txtEventType = CStr(Nz(head!EventType, ""))
    report!txtAbbr = CStr(Nz(head!Abbreviation, ""))
    report!txtSAP = CStr(Nz(head!SAPID, ""))
    report!txtBSS = CStr(Nz(head!BSSID, ""))
    report!txtSumFood = CStr(Nz(head!FoodNet, ""))
    report!txtSumAlc = CStr(Nz(head!AlcoholNet, ""))
    report!txtSumGroup = CStr(Nz(head!GroupNet, ""))
    report!txtSumGross = CStr(Nz(head!GrossSales, ""))
    report!txtNotes = CStr(Nz(head!Notes, ""))
    report!txtPrep = CStr(Nz(head!PreparedBy, ""))
    report!txtPrepTitle = CStr(Nz(head!PreparedTitle, ""))
    report!txtPrepWhen = CStr(Nz(head!PreparedWhen, ""))
    report!txtPrepEvent = CStr(Nz(head!EventWhenShort, ""))
    report!txtRev = CStr(Nz(head!ReviewedBy, ""))
    report!txtRevTitle = CStr(Nz(head!ReviewedTitle, ""))
    report!txtRevWhen = CStr(Nz(head!ReviewedWhen, ""))
    report!txtRevEvent = CStr(Nz(head!EventWhenShort, ""))
    report!txtAppr = CStr(Nz(head!ApprovedBy, ""))
    report!txtApprTitle = CStr(Nz(head!ApprovedTitle, ""))
    report!txtApprWhen = CStr(Nz(head!ApprovedWhen, ""))
    report!txtApprEvent = CStr(Nz(head!EventWhenShort, ""))
    If CStr(Nz(head!Kind, "")) = "NPO" Then
        SetCaption report, "rFoodNet", "FOOD / NON-ALC NET"
        SetCaption report, "hFoodNet", "FOOD / NON-ALC NET"
        SetCaption report, "rAlcNet", "ALCOHOL NET"
        SetCaption report, "hAlcNet", "ALCOHOL NET"
    Else
        SetCaption report, "rFoodNet", "FOOD NET"
        SetCaption report, "hFoodNet", "FOOD NET"
        SetCaption report, "rAlcNet", "NA / ALCOHOL NET"
        SetCaption report, "hAlcNet", "NA / ALCOHOL NET"
    End If
    SetCaption report, "rFoodComm", CStr(Nz(head!FoodRateLabel, ""))
    SetCaption report, "hFoodComm", CStr(Nz(head!FoodRateLabel, ""))
    SetCaption report, "rAlcComm", CStr(Nz(head!AlcoholRateLabel, ""))
    SetCaption report, "hAlcComm", CStr(Nz(head!AlcoholRateLabel, ""))
    For index = 1 To 8
        report("txtSup" & index) = ""
        report("txtSupAmt" & index) = ""
        report("txtGl" & index) = ""
        report("txtGlCode" & index) = ""
        report("txtGlAmt" & index) = ""
    Next index
    Set extras = CurrentDb.OpenRecordset("SELECT * FROM tblPrintExtra WHERE SortOrder<100 ORDER BY SortOrder", EMS_SNAPSHOT)
    index = 1
    Do While Not extras.EOF And index <= 8
        report("txtSup" & index) = CStr(Nz(extras!Label, ""))
        report("txtSupAmt" & index) = CStr(Nz(extras!Amount, ""))
        index = index + 1
        extras.MoveNext
    Loop
    extras.Close
    Set extras = CurrentDb.OpenRecordset("SELECT * FROM tblPrintExtra WHERE SortOrder>=100 ORDER BY SortOrder", EMS_SNAPSHOT)
    index = 1
    Do While Not extras.EOF And index <= 8
        report("txtGl" & index) = CStr(Nz(extras!Label, ""))
        report("txtGlCode" & index) = CStr(Nz(extras!GLText, ""))
        report("txtGlAmt" & index) = CStr(Nz(extras!Amount, ""))
        index = index + 1
        extras.MoveNext
    Loop
    extras.Close
    LayoutColumns report, CBool(Nz(head!ShowAdjustments, False))
    head.Close
    Exit Function
Failed:
    MsgBox "The statement could not be filled." & vbCrLf & Err.Description, vbCritical
End Function

Public Function EMSStatementDetail()
    On Error Resume Next
    Dim report As Report
    Dim bold As Boolean
    Dim ctl As Control
    Set report = Reports("rptStatement")
    bold = (Nz(report!txtService, "") = "LOCATION TOTAL")
    For Each ctl In report.Section(acDetail).Controls
        ctl.FontBold = bold
    Next ctl
End Function

Public Function EMSPageFooter() As String
    EMSPageFooter = gPageFooter
End Function

Private Sub SetCaption(ByVal report As Report, ByVal controlName As String, ByVal caption As String)
    report.Controls(controlName).Caption = caption
End Sub

Private Sub LayoutColumns(ByVal report As Report, ByVal showAdj As Boolean)
    Dim widths As Variant
    Dim names As Variant
    Dim index As Long
    Dim leftPos As Long
    names = Array("Service", "Net", "FoodNet", "FoodComm", "AlcNet", "AlcComm", "Comm", "Grat", "Adj", "Pay")
    If showAdj Then
        widths = Array(1740, 1040, 1040, 940, 1040, 940, 1000, 1000, 900, 1100)
    Else
        widths = Array(2320, 1100, 1100, 1000, 1080, 920, 1020, 1040, 0, 1160)
    End If
    leftPos = 0
    For index = 0 To 9
        Place report.Controls("r" & names(index)), leftPos, CLng(widths(index))
        Place report.Controls("h" & names(index)), leftPos, CLng(widths(index))
        Place report.Controls("txt" & names(index)), leftPos, CLng(widths(index))
        leftPos = leftPos + CLng(widths(index))
    Next index
End Sub

Private Sub Place(ByVal ctl As Control, ByVal leftPos As Long, ByVal width As Long)
    ctl.Left = leftPos
    ctl.Width = width
    ctl.Visible = (width > 0)
End Sub
