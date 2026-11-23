Attribute VB_Name = "EMS_Build"
Option Compare Database
Option Explicit

' Open the database, Alt+F11, File > Import File, and import every module in the vba folder.
' In the Immediate window run BuildEMS once, then close and reopen.
' That builds the event list, the event screen, the payee screen, and the portrait statement.
' It does not run again on its own, and it does not lock menus or the navigation pane.

Private Const PAGE_WIDTH As Long = 10740

Public Sub BuildEMS()
    On Error GoTo Failed
    DoCmd.Hourglass True
    DoCmd.Echo False, "Building EMS"
    DropObject acForm, "frmEvent"
    DropObject acForm, "frmHub"
    DropObject acForm, "frmPayee"
    DropObject acForm, "fsubAssign"
    DropObject acForm, "fsubVendor"
    DropObject acForm, "fsubSupplement"
    DropObject acForm, "fsubStaff"
    DropObject acForm, "fsubEmail"
    DropObject acForm, "fsubRate"
    DropObject acReport, "rptStatement"
    CaptionFields
    BuildAssign
    BuildVendor
    BuildSupplement
    BuildStaff
    BuildEmail
    BuildRate
    BuildPayee
    BuildEvent
    BuildHub
    BuildStatement
    SetTextProperty "StartupForm", "frmHub"
    SetTextProperty "AppTitle", "EMS Payouts"
    DoCmd.Echo True
    DoCmd.Hourglass False
    MsgBox "EMS is ready." & vbCrLf & "Close this database and open it again. The event list is the first screen." & vbCrLf & _
        "Use New event when there is no MyVenue file. Import replaces the sales on an event you already have open.", vbInformation
    Exit Sub
Failed:
    DoCmd.Echo True
    DoCmd.Hourglass False
    MsgBox "Build stopped." & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub BuildHub()
    Dim frm As Form
    Dim ctl As Control
    Set frm = CreateForm()
    frm.Caption = "EMS Payouts"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.Width = 14000
    frm.OnLoad = "=EMSHubSearch()"
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 1500
    LabelForm frm.Name, "lblTitle", "Events", 200, 80, 4000, 360, 18, True
    LabelForm frm.Name, "lblHint", "New event types a settlement with no import. Search covers the date, the event, the payee, the stand, and the invoice.", 200, 480, 11000, 280, 9, False
    TextForm frm.Name, "txtFrom", 200, 900, 1600, 320, "From"
    TextForm frm.Name, "txtTo", 2000, 900, 1600, 320, "To"
    TextForm frm.Name, "txtFind", 3800, 900, 3600, 320, "Find"
    ButtonForm frm.Name, "cmdSearch", "Search", "=EMSHubSearch()", 7600, 880, 1400, 360
    ButtonForm frm.Name, "cmdNew", "New event", "=EMSNewEvent()", 9200, 880, 1600, 360
    ButtonForm frm.Name, "cmdImport", "Import reports", "=EMSImportNew()", 11000, 880, 1800, 360
    ButtonForm frm.Name, "cmdPayees", "Payees", "=EMSOpenPayees()", 13000, 880, 800, 360
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstEvents", 200, 200, 13400, 6200)
    ctl.ColumnCount = 5
    ctl.ColumnWidths = "0;1600;6200;2800;1800"
    ctl.BoundColumn = 1
    ctl.OnDblClick = "=EMSOpenSelected()"
    frm.Section(acDetail).Height = 6800
    SaveObject acForm, frm.Name, "frmHub"
End Sub

Private Sub BuildEvent()
    Dim frm As Form
    Dim ctl As Control
    Set frm = CreateForm()
    frm.Caption = "Event"
    frm.RecordSource = "tblEvent"
    frm.Width = 16200
    frm.OnCurrent = "=EMSEventOpen()"
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 780
    ButtonForm frm.Name, "cmdBack", "Events", "=EMSBackToEvents()", 80, 160, 1200, 400
    ButtonForm frm.Name, "cmdImport", "Import onto this event", "=EMSImportCurrent()", 1400, 160, 2600, 400
    ButtonForm frm.Name, "cmdCalc", "Calculate", "=EMSCalculateCurrent()", 4120, 160, 1500, 400
    ButtonForm frm.Name, "cmdPreview", "Preview", "=EMSPreviewCurrent()", 5740, 160, 1400, 400
    ButtonForm frm.Name, "cmdPdf", "PDF", "=EMSExportCurrent()", 7260, 160, 1200, 400
    ButtonForm frm.Name, "cmdMail", "Email", "=EMSEmailCurrent()", 8580, 160, 1400, 400
    ButtonForm frm.Name, "cmdDelete", "Delete event", "=EMSDeleteEvent()", 12800, 160, 1800, 400
    BoundBox frm.Name, "EventID", "", 0, 0, 10, 10, False
    BoundBox frm.Name, "EventDate", "Event date", 200, 40, 1600, 280, True
    BoundBox frm.Name, "EventName", "Event", 2000, 40, 5200, 280, True
    BoundBox frm.Name, "EventType", "Type", 7400, 40, 2400, 280, True
    BoundBox frm.Name, "StatusText", "Status", 10000, 40, 1600, 280, True
    BoundBox frm.Name, "Venue", "Venue", 200, 420, 4000, 280, True
    BoundBox frm.Name, "MinimumCommission", "Minimum commission", 4400, 420, 1800, 280, True
    BoundBox frm.Name, "PreparedBy", "Prepared by", 200, 800, 2400, 260, True
    BoundBox frm.Name, "ReviewedBy", "Reviewed by", 2800, 800, 2400, 260, True
    BoundBox frm.Name, "ApprovedBy", "Approved by", 5400, 800, 2400, 260, True
    BoundBox frm.Name, "PreparedDate", "Prepared", 8000, 800, 1500, 260, True
    BoundBox frm.Name, "ReviewedDate", "Reviewed", 9700, 800, 1500, 260, True
    BoundBox frm.Name, "ApprovedDate", "Approved", 11400, 800, 1500, 260, True
    BoundBox frm.Name, "Notes", "Notes", 200, 1160, 9000, 460, True
    Set ctl = AddControl(frm.Name, acForm, acTextBox, acDetail, "txtDue", 9400, 1340, 2800, 320)
    ctl.ControlSource = "=Nz(DSum('AmountDue','tblInvoice','EventID=' & Nz([EventID],0)),0)"
    ctl.Format = "Currency"
    ctl.FontBold = True
    LabelDetail frm.Name, "lblDue", "Amount due", 9400, 1160, 2000, 180, 8, True
    LabelDetail frm.Name, "lblHow", "Add a stand in the grid when there is no file, and type its nets. Import fills those nets from the price report. A blank people count holds only that stand's tips.", 200, 1760, 15000, 240, 9, False
    LinkSubform frm.Name, "fsubAssign", "fsubAssign", 120, 2060, 15800, 2500
    LinkSubform frm.Name, "fsubVendor", "fsubVendor", 120, 4660, 15800, 1200
    LinkSubform frm.Name, "fsubSupplement", "fsubSupplement", 120, 5960, 7800, 1100
    LinkSubform frm.Name, "fsubStaff", "fsubStaff", 8100, 5960, 7800, 1100
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstInvoices", 120, 7180, 7800, 1300)
    ctl.RowSource = "SELECT i.InvoiceID, p.Abbreviation, i.InvoiceNumber, i.AmountDue FROM tblInvoice AS i INNER JOIN tblPayee AS p ON i.PayeeID=p.PayeeID WHERE i.EventID=[Forms]![frmEvent]![EventID] ORDER BY p.Abbreviation"
    ctl.ColumnCount = 4
    ctl.ColumnWidths = "0;1200;3200;1600"
    ctl.BoundColumn = 1
    ctl.OnDblClick = "=EMSPreviewCurrent()"
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstExceptions", 8100, 7180, 7800, 1300)
    ctl.RowSource = "SELECT Kind, Detail FROM tblException WHERE EventID=[Forms]![frmEvent]![EventID] ORDER BY ExceptionID"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "1600;6000"
    frm.Section(acDetail).Height = 8600
    SaveObject acForm, frm.Name, "frmEvent"
End Sub

Private Sub BuildPayee()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Payees"
    frm.RecordSource = "tblPayee"
    frm.Width = 14000
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 700
    ButtonForm frm.Name, "cmdEvents", "Events", "=EMSBackToEvents()", 200, 160, 1400, 360
    ButtonForm frm.Name, "cmdNew", "New payee", "=GoToNewPayee()", 1800, 160, 1600, 360
    BoundBox frm.Name, "PayeeID", "", 0, 0, 10, 10, False
    BoundBox frm.Name, "PayeeName", "Payee", 200, 40, 5000, 280, True
    BoundBox frm.Name, "PayeeType", "Type", 5400, 40, 1200, 280, True
    BoundBox frm.Name, "Abbreviation", "Abbrev", 6800, 40, 1400, 280, True
    BoundBox frm.Name, "Active", "Active", 8400, 40, 800, 280, True
    BoundBox frm.Name, "SAPID", "SAP", 200, 420, 1800, 280, True
    BoundBox frm.Name, "BSSID", "BSS", 2200, 420, 1800, 280, True
    BoundBox frm.Name, "Address1", "Address", 4200, 420, 4500, 280, True
    BoundBox frm.Name, "Address2", "City", 8900, 420, 4500, 280, True
    BoundBox frm.Name, "Notes", "Notes", 200, 800, 8000, 360, True
    LinkSubform frm.Name, "fsubEmail", "fsubEmail", 200, 1400, 6400, 2200
    LinkSubform frm.Name, "fsubRate", "fsubRate", 6800, 1400, 6800, 2200
    frm.Section(acDetail).Height = 3800
    SaveObject acForm, frm.Name, "frmPayee"
End Sub

Public Function GoToNewPayee()
    DoCmd.GoToRecord acDataForm, "frmPayee", acNewRec
End Function

Private Sub BuildAssign()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Stands"
    frm.RecordSource = "tblAssign"
    frm.DefaultView = 2
    DatasheetBox frm.Name, "LocationName", 2400
    DatasheetBox frm.Name, "DepartmentName", 1800
    DatasheetBox frm.Name, "PosSales", 1200
    DatasheetBox frm.Name, "FoodNet", 1200
    DatasheetBox frm.Name, "NaNet", 1100
    DatasheetBox frm.Name, "AlcoholNet", 1200
    DatasheetBox frm.Name, "PosTips", 1100
    DatasheetCombo frm.Name, "Role", "Value List", "NPO;SUB;STAFF;OPEN", 1, "1400", 1100
    DatasheetCombo frm.Name, "PayeeID", "Table/Query", "SELECT PayeeID, PayeeName FROM tblPayee WHERE Active<>0 ORDER BY PayeeName", 2, "0;3600", 2400
    DatasheetBox frm.Name, "Headcount", 900
    DatasheetBox frm.Name, "StaffNames", 2200
    DatasheetBox frm.Name, "CookFee", 1000
    DatasheetBox frm.Name, "Bonus", 1000
    DatasheetBox frm.Name, "Shortage", 1000
    DatasheetBox frm.Name, "Cleaning", 1000
    SaveObject acForm, frm.Name, "fsubAssign"
End Sub

Private Sub BuildVendor()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Vendor tenders and ice"
    frm.RecordSource = "tblVendorInput"
    frm.DefaultView = 2
    DatasheetCombo frm.Name, "PayeeID", "Table/Query", "SELECT PayeeID, PayeeName FROM tblPayee WHERE PayeeType='SUB' AND Active<>0 ORDER BY PayeeName", 2, "0;3600", 2400
    DatasheetBox frm.Name, "CashTender", 1200
    DatasheetBox frm.Name, "CardTender", 1200
    DatasheetBox frm.Name, "CorporateVoucher", 1400
    DatasheetBox frm.Name, "LevyNfpVoucher", 1400
    DatasheetBox frm.Name, "UnmGuestVoucher", 1400
    DatasheetBox frm.Name, "IceBags", 1000
    DatasheetBox frm.Name, "OtherExpense", 1200
    SaveObject acForm, frm.Name, "fsubVendor"
End Sub

Private Sub BuildSupplement()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Supplements"
    frm.RecordSource = "tblSupplement"
    frm.DefaultView = 2
    DatasheetCombo frm.Name, "PayeeID", "Table/Query", "SELECT PayeeID, Abbreviation FROM tblPayee WHERE Active<>0 ORDER BY Abbreviation", 2, "0;1400", 1200
    DatasheetBox frm.Name, "Label", 2200
    DatasheetBox frm.Name, "Amount", 1200
    DatasheetBox frm.Name, "GLCode", 1000
    SaveObject acForm, frm.Name, "fsubSupplement"
End Sub

Private Sub BuildStaff()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Staff tips"
    frm.RecordSource = "tblStaffTip"
    frm.DefaultView = 2
    frm.AllowAdditions = False
    frm.AllowEdits = False
    frm.AllowDeletions = False
    DatasheetBox frm.Name, "LocationName", 2800
    DatasheetBox frm.Name, "SlotNo", 700
    DatasheetBox frm.Name, "PersonName", 1800
    DatasheetBox frm.Name, "Amount", 1100
    DatasheetBox frm.Name, "PoolAmount", 1200
    DatasheetBox frm.Name, "HeldBack", 1100
    SaveObject acForm, frm.Name, "fsubStaff"
End Sub

Private Sub BuildEmail()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Emails"
    frm.RecordSource = "tblPayeeEmail"
    frm.DefaultView = 2
    DatasheetBox frm.Name, "EmailAddress", 3200
    DatasheetBox frm.Name, "Label", 1200
    DatasheetBox frm.Name, "IsPrimary", 900
    SaveObject acForm, frm.Name, "fsubEmail"
End Sub

Private Sub BuildRate()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Rates"
    frm.RecordSource = "tblPayeeRate"
    frm.DefaultView = 2
    DatasheetBox frm.Name, "RateCode", 1800
    DatasheetBox frm.Name, "RateAmount", 1100
    DatasheetBox frm.Name, "EffectiveDate", 1400
    DatasheetBox frm.Name, "EndDate", 1400
    DatasheetBox frm.Name, "Notes", 2200
    SaveObject acForm, frm.Name, "fsubRate"
End Sub

Private Sub BuildStatement()
    Dim rpt As Report
    Dim temp As String
    Dim index As Long
    Dim names As Variant
    Dim captions As Variant
    Dim sources As Variant
    Set rpt = CreateReport()
    temp = rpt.Name
    rpt.Caption = "Settlement Statement"
    rpt.RecordSource = "SELECT * FROM tblPrintLine ORDER BY SortOrder"
    rpt.Width = PAGE_WIDTH
    EnsureReportChrome temp
    rpt.Section(acHeader).Height = 4400
    rpt.Section(acPageHeader).Height = 900
    rpt.Section(acDetail).Height = 340
    rpt.Section(acDetail).CanGrow = True
    rpt.Section(acDetail).KeepTogether = True
    rpt.Section(acDetail).OnFormat = "=EMSStatementDetail()"
    rpt.Section(acFooter).Height = 4200
    rpt.Section(acFooter).KeepTogether = True
    rpt.Section(acPageFooter).Height = 280
    rpt.OnOpen = "=EMSStatementOpen()"
    rpt.PageHeader = 1
    On Error Resume Next
    rpt.UseDefaultPrinter = False
    rpt.Printer.PaperSize = acPRPSLetter
    rpt.Printer.Orientation = acPRORPortrait
    rpt.Printer.LeftMargin = 720
    rpt.Printer.RightMargin = 720
    rpt.Printer.TopMargin = 600
    rpt.Printer.BottomMargin = 600
    On Error GoTo FailedBuild
    TitleLabel temp, "lblTitle", "SETTLEMENT STATEMENT", 40, 22
    TitleLabel temp, "lblBrand", "COMPASS GROUP USA | LEVY - UNM ATHLETICS", 460, 10
    Reports(temp).Controls("lblBrand").ForeColor = RGB(180, 32, 40)
    UnboundBox temp, acHeader, "txtStatus", 0, 720, PAGE_WIDTH, 240, 11, True, 2
    Reports(temp).Controls("txtStatus").ForeColor = RGB(180, 32, 40)
    Bar temp, acHeader, "lblRule", "", 0, 980, PAGE_WIDTH, 36
    UnboundBox temp, acHeader, "txtPayee", 0, 1060, 6600, 320, 16, True, 1
    UnboundBox temp, acHeader, "txtAddress", 0, 1400, 6600, 280, 9, False, 1
    CaptionLabel temp, acHeader, "lblDateCap", "Statement date", 7000, 1060, 1600, 200
    UnboundBox temp, acHeader, "txtDate", 8600, 1040, 2140, 240, 9, False, 3
    CaptionLabel temp, acHeader, "lblInvCap", "Invoice No.", 7000, 1320, 1600, 200
    UnboundBox temp, acHeader, "txtInv", 8200, 1300, 2540, 240, 9, False, 3
    CaptionLabel temp, acHeader, "lblDueCap", "AMOUNT DUE", 7000, 1600, 1700, 240
    Reports(temp).Controls("lblDueCap").FontBold = True
    UnboundBox temp, acHeader, "txtDue", 8000, 1840, 2740, 320, 14, True, 3
    captions = Array("EVENT DATE", "EVENT TYPE", "PAYEE / NPO", "SAP ID", "BSS ID")
    names = Array("txtEventDate", "txtEventType", "txtAbbr", "txtSAP", "txtBSS")
    For index = 0 To 4
        CaptionLabel temp, acHeader, "lblMeta" & index, CStr(captions(index)), index * 2148, 2240, 2050, 200
        UnboundBox temp, acHeader, CStr(names(index)), index * 2148, 2460, 2000, 280, 9, False, 1
    Next index
    Reports(temp).Controls("txtEventDate").BackStyle = 1
    Reports(temp).Controls("txtEventDate").BackColor = RGB(232, 236, 240)
    Reports(temp).Controls("txtAbbr").BackStyle = 1
    Reports(temp).Controls("txtAbbr").BackColor = RGB(255, 242, 204)
    Bar temp, acHeader, "lblSales", "EVENT SALES OVERVIEW", 0, 2840, PAGE_WIDTH, 280
    captions = Array("FOOD / NON-ALCOHOL NET", "ALCOHOL NET", "GROUP NET SALES", "VALIDATED GROSS SALES")
    names = Array("txtFoodNet", "txtAlcNet", "txtGroupNet", "txtGross")
    For index = 0 To 3
        CaptionLabel temp, acHeader, "lblSales" & index, CStr(captions(index)), index * 2685, 3160, 2500, 220
        UnboundBox temp, acHeader, CStr(names(index)), index * 2685, 3380, 2500, 300, 12, True, 1
    Next index
    Bar temp, acHeader, "lblLoc", "LOCATION SERVICE DETAIL", 0, 3720, PAGE_WIDTH, 280
    Bar temp, acPageHeader, "lblLocPage", "LOCATION SERVICE DETAIL", 0, 0, PAGE_WIDTH, 280
    names = Array("Service", "Net", "FoodNet", "FoodComm", "AlcNet", "AlcComm", "Comm", "Grat", "Adj", "Pay")
    captions = Array("SERVICE AREA", "NET SALES", "FOOD / NON-ALC NET", "10% COMM.", "ALCOHOL NET", "8% COMM.", "TOTAL COMM.", "GRATUITY", "BONUS / ADJ.", "LOCATION PAYOUT")
    sources = Array("ServiceArea", "NetSales", "FoodNet", "FoodCommission", "AlcoholNet", "AlcoholCommission", "Commission", "Gratuity", "Adjustments", "Payout")
    For index = 0 To 9
        ColumnHead temp, "r" & names(index), CStr(captions(index)), acHeader, 3920
        ColumnHead temp, "h" & names(index), CStr(captions(index)), acPageHeader, 300
        If index = 0 Then
            DetailBox temp, "txt" & names(index), CStr(sources(index)), 1
        Else
            DetailBox temp, "txt" & names(index), CStr(sources(index)), 3
        End If
    Next index
    Bar temp, acFooter, "lblNotesBar", "SETTLEMENT NOTES", 0, 40, 5200, 280
    Bar temp, acFooter, "lblSupBar", "SUPPLEMENTAL FINANCIALS", 5500, 40, 5240, 280
    UnboundBox temp, acFooter, "txtNotes", 0, 340, 5200, 1100, 9, False, 1
    Reports(temp).Controls("txtNotes").BackStyle = 1
    Reports(temp).Controls("txtNotes").BackColor = RGB(255, 249, 230)
    Reports(temp).Controls("txtNotes").BorderStyle = 1
    Reports(temp).Controls("txtNotes").BorderColor = RGB(230, 214, 150)
    For index = 1 To 8
        UnboundBox temp, acFooter, "txtSup" & index, 5500, 340 + (index - 1) * 180, 3000, 180, 8, False, 1
        UnboundBox temp, acFooter, "txtSupAmt" & index, 8600, 340 + (index - 1) * 180, 1900, 180, 8, False, 3
    Next index
    Bar temp, acFooter, "lblAuth", "INTERNAL AUTHORIZATION CHECK REQUEST", 0, 1900, 5200, 280
    Bar temp, acFooter, "lblGl", "INTERNAL - ACCOUNTING GL ENTRY - SUMMARY", 5500, 1900, 5240, 280
    UnboundBox temp, acFooter, "txtPrep", 0, 2240, 2600, 220, 9, True, 1
    UnboundBox temp, acFooter, "txtPrepTitle", 0, 2460, 2600, 200, 8, False, 1
    UnboundBox temp, acFooter, "txtPrepWhen", 2700, 2240, 2300, 220, 8, False, 3
    UnboundBox temp, acFooter, "txtPrepEvent", 2700, 2460, 2300, 200, 8, False, 3
    UnboundBox temp, acFooter, "txtRev", 0, 2760, 2600, 220, 9, True, 1
    UnboundBox temp, acFooter, "txtRevTitle", 0, 2980, 2600, 200, 8, False, 1
    UnboundBox temp, acFooter, "txtRevWhen", 2700, 2760, 2300, 220, 8, False, 3
    UnboundBox temp, acFooter, "txtRevEvent", 2700, 2980, 2300, 200, 8, False, 3
    UnboundBox temp, acFooter, "txtAppr", 0, 3280, 2600, 220, 9, True, 1
    UnboundBox temp, acFooter, "txtApprTitle", 0, 3500, 2600, 200, 8, False, 1
    UnboundBox temp, acFooter, "txtApprWhen", 2700, 3280, 2300, 220, 8, False, 3
    UnboundBox temp, acFooter, "txtApprEvent", 2700, 3500, 2300, 200, 8, False, 3
    For index = 1 To 8
        UnboundBox temp, acFooter, "txtGl" & index, 5500, 2240 + (index - 1) * 220, 2000, 200, 8, False, 1
        UnboundBox temp, acFooter, "txtGlCode" & index, 7500, 2240 + (index - 1) * 220, 2100, 200, 7, False, 1
        UnboundBox temp, acFooter, "txtGlAmt" & index, 9600, 2240 + (index - 1) * 220, 1100, 200, 8, False, 3
    Next index
    UnboundBox temp, acPageFooter, "txtFoot", 0, 40, 8000, 200, 8, False, 1
    Reports(temp).Controls("txtFoot").ControlSource = "=EMSPageFooter()"
    UnboundBox temp, acPageFooter, "txtPage", 8200, 40, 2540, 200, 8, False, 3
    Reports(temp).Controls("txtPage").ControlSource = "='Page ' & [Page] & ' of ' & [Pages]"
    SaveObject acReport, temp, "rptStatement"
    Exit Sub
FailedBuild:
    Err.Raise Err.Number, "BuildStatement", Err.Description
End Sub

Private Sub CaptionFields()
    FieldCaption "tblEvent", "EventDate", "Event date"
    FieldCaption "tblEvent", "EventName", "Event"
    FieldCaption "tblEvent", "EventType", "Type"
    FieldCaption "tblEvent", "StatusText", "Status"
    FieldCaption "tblEvent", "MinimumCommission", "Minimum commission"
    FieldCaption "tblEvent", "PreparedBy", "Prepared by"
    FieldCaption "tblEvent", "ReviewedBy", "Reviewed by"
    FieldCaption "tblEvent", "ApprovedBy", "Approved by"
    FieldCaption "tblEvent", "PreparedDate", "Prepared"
    FieldCaption "tblEvent", "ReviewedDate", "Reviewed"
    FieldCaption "tblEvent", "ApprovedDate", "Approved"
    FieldCaption "tblAssign", "LocationName", "Stand"
    FieldCaption "tblAssign", "DepartmentName", "Department"
    FieldCaption "tblAssign", "PosSales", "POS sales"
    FieldCaption "tblAssign", "FoodNet", "Food net"
    FieldCaption "tblAssign", "NaNet", "NA net"
    FieldCaption "tblAssign", "AlcoholNet", "Alcohol net"
    FieldCaption "tblAssign", "PosTips", "Tips"
    FieldCaption "tblAssign", "PayeeID", "Payee"
    FieldCaption "tblAssign", "Headcount", "People"
    FieldCaption "tblAssign", "StaffNames", "Staff names"
    FieldCaption "tblAssign", "CookFee", "Cook fee"
    FieldCaption "tblVendorInput", "PayeeID", "Vendor"
    FieldCaption "tblVendorInput", "CashTender", "Cash"
    FieldCaption "tblVendorInput", "CardTender", "Cards"
    FieldCaption "tblVendorInput", "CorporateVoucher", "Corporate voucher"
    FieldCaption "tblVendorInput", "LevyNfpVoucher", "Levy voucher"
    FieldCaption "tblVendorInput", "UnmGuestVoucher", "UNM voucher"
    FieldCaption "tblVendorInput", "IceBags", "Ice bags"
    FieldCaption "tblVendorInput", "OtherExpense", "Other"
    FieldCaption "tblSupplement", "PayeeID", "Payee"
    FieldCaption "tblSupplement", "GLCode", "GL"
    FieldCaption "tblStaffTip", "LocationName", "Stand"
    FieldCaption "tblStaffTip", "SlotNo", "#"
    FieldCaption "tblStaffTip", "PersonName", "Person"
    FieldCaption "tblStaffTip", "PoolAmount", "Pool"
    FieldCaption "tblStaffTip", "HeldBack", "Held"
    FieldCaption "tblPayee", "PayeeName", "Payee"
    FieldCaption "tblPayee", "PayeeType", "Type"
    FieldCaption "tblPayee", "Abbreviation", "Abbrev"
    FieldCaption "tblPayeeEmail", "EmailAddress", "Email"
    FieldCaption "tblPayeeEmail", "IsPrimary", "Primary"
    FieldCaption "tblPayeeRate", "RateCode", "Rate"
    FieldCaption "tblPayeeRate", "RateAmount", "Amount"
    FieldCaption "tblPayeeRate", "EffectiveDate", "Starts"
    FieldCaption "tblPayeeRate", "EndDate", "Ends"
End Sub

Private Function AddControl(ByVal objectName As String, ByVal objectKind As AcObjectType, ByVal controlKind As Long, ByVal section As Long, ByVal controlName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    Dim host As Object
    Dim before As Long
    Dim index As Long
    If objectKind = acForm Then
        Set host = Forms(objectName)
        before = host.Controls.Count
        Set AddControl = CreateControl(objectName, controlKind, section, "", "", x, y, width, height)
    Else
        Set host = Reports(objectName)
        before = host.Controls.Count
        Set AddControl = CreateReportControl(objectName, controlKind, section, "", "", x, y, width, height)
    End If
    AddControl.Name = controlName
    For index = host.Controls.Count - 1 To before Step -1
        If host.Controls(index).Name <> controlName Then DeleteControl objectName, host.Controls(index).Name
    Next index
End Function

Private Sub Paint(ByVal ctl As Control, ByVal size As Single, ByVal bold As Boolean)
    On Error Resume Next
    ctl.FontName = "Calibri"
    ctl.FontSize = size
    ctl.FontBold = bold
    ctl.ForeColor = RGB(24, 49, 78)
End Sub

Private Function ButtonForm(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal onClick As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    If height = 0 Then Exit Function
    Set ButtonForm = AddControl(formName, acForm, acCommandButton, acHeader, controlName, x, y, width, height)
    ButtonForm.Caption = caption
    ButtonForm.OnClick = onClick
    Paint ButtonForm, 9, False
End Function

Private Sub LabelForm(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal size As Single, ByVal bold As Boolean)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acLabel, acHeader, controlName, x, y, width, height)
    ctl.Caption = caption
    Paint ctl, size, bold
End Sub

Private Sub LabelDetail(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal size As Single, ByVal bold As Boolean)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acLabel, acDetail, controlName, x, y, width, height)
    ctl.Caption = caption
    Paint ctl, size, bold
End Sub

Private Sub TextForm(ByVal formName As String, ByVal controlName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal caption As String)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acTextBox, acHeader, controlName, x, y, width, height)
    Paint ctl, 9, False
    ctl.BorderStyle = 1
    LabelForm formName, "lbl" & controlName, caption, x, y - 220, width, 200, 8, False
End Sub

Private Sub BoundBox(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal shown As Boolean)
    Dim ctl As Control
    If shown And Len(caption) > 0 Then LabelDetail formName, "lbl" & fieldName, caption, x, y, width, 180, 8, False
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, fieldName, x, y + IIf(shown, 180, 0), width, height)
    ctl.ControlSource = fieldName
    ctl.Visible = shown
    Paint ctl, 9, False
    If shown Then ctl.BorderStyle = 1
End Sub

Private Sub LinkSubform(ByVal formName As String, ByVal controlName As String, ByVal source As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acSubform, acDetail, controlName, x, y, width, height)
    ctl.SourceObject = source
    ctl.LinkMasterFields = "PayeeID"
    ctl.LinkChildFields = "PayeeID"
    If Left$(source, 4) <> "fsub" Or source = "fsubAssign" Or source = "fsubVendor" Or source = "fsubSupplement" Or source = "fsubStaff" Then
        If source <> "fsubEmail" And source <> "fsubRate" Then
            ctl.LinkMasterFields = "EventID"
            ctl.LinkChildFields = "EventID"
        End If
    End If
End Sub

Private Sub DatasheetBox(ByVal formName As String, ByVal fieldName As String, ByVal width As Long)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, "txt" & fieldName, 0, 0, width, 300)
    ctl.ControlSource = fieldName
End Sub

Private Sub DatasheetCombo(ByVal formName As String, ByVal fieldName As String, ByVal sourceType As String, ByVal source As String, ByVal columns As Long, ByVal widths As String, ByVal width As Long)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acComboBox, acDetail, "cbo" & fieldName, 0, 0, width, 300)
    ctl.ControlSource = fieldName
    ctl.RowSourceType = sourceType
    ctl.RowSource = source
    ctl.ColumnCount = columns
    ctl.ColumnWidths = widths
    ctl.BoundColumn = 1
    ctl.LimitToList = True
End Sub

Private Sub TitleLabel(ByVal reportName As String, ByVal controlName As String, ByVal caption As String, ByVal y As Long, ByVal size As Single)
    Dim ctl As Control
    Set ctl = AddControl(reportName, acReport, acLabel, acHeader, controlName, 0, y, PAGE_WIDTH, 280)
    ctl.Caption = caption
    ctl.TextAlign = 2
    Paint ctl, size, True
End Sub

Private Sub CaptionLabel(ByVal reportName As String, ByVal section As Long, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Dim ctl As Control
    Set ctl = AddControl(reportName, acReport, acLabel, section, controlName, x, y, width, height)
    ctl.Caption = caption
    Paint ctl, 8, True
End Sub

Private Function Bar(ByVal reportName As String, ByVal section As Long, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    Set Bar = AddControl(reportName, acReport, acLabel, section, controlName, x, y, width, height)
    Bar.Caption = " " & caption
    Paint Bar, 9, True
    Bar.BackStyle = 1
    Bar.BackColor = RGB(24, 54, 90)
    Bar.ForeColor = vbWhite
End Function

Private Sub ColumnHead(ByVal reportName As String, ByVal controlName As String, ByVal caption As String, ByVal section As Long, ByVal y As Long)
    Dim ctl As Control
    Set ctl = AddControl(reportName, acReport, acLabel, section, controlName, 0, y, 800, 460)
    ctl.Caption = caption
    Paint ctl, 7, True
    ctl.BackStyle = 1
    ctl.BackColor = RGB(24, 54, 90)
    ctl.ForeColor = vbWhite
    ctl.TextAlign = 2
End Sub

Private Sub DetailBox(ByVal reportName As String, ByVal controlName As String, ByVal source As String, ByVal align As Long)
    Dim ctl As Control
    Set ctl = AddControl(reportName, acReport, acTextBox, acDetail, controlName, 0, 20, 800, 280)
    ctl.ControlSource = source
    ctl.TextAlign = align
    ctl.CanGrow = True
    Paint ctl, 8, False
End Sub

Private Sub UnboundBox(ByVal reportName As String, ByVal section As Long, ByVal controlName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal size As Single, ByVal bold As Boolean, ByVal align As Long)
    Dim ctl As Control
    Set ctl = AddControl(reportName, acReport, acTextBox, section, controlName, x, y, width, height)
    ctl.TextAlign = align
    Paint ctl, size, bold
End Sub

Private Sub EnsureFormHeader(ByVal formName As String)
    Dim sec As Section
    On Error Resume Next
    Set sec = Forms(formName).Section(acHeader)
    On Error GoTo 0
    If sec Is Nothing Then DoCmd.RunCommand acCmdFormHdrFtr
    Forms(formName).Section(acHeader).Height = 800
End Sub

Private Sub EnsureReportChrome(ByVal reportName As String)
    Dim sec As Section
    On Error Resume Next
    Set sec = Reports(reportName).Section(acHeader)
    On Error GoTo 0
    If sec Is Nothing Then DoCmd.RunCommand acCmdReportHdrFtr
    Reports(reportName).Section(acHeader).Visible = True
    Reports(reportName).Section(acFooter).Visible = True
End Sub

Private Sub SaveObject(ByVal kind As AcObjectType, ByVal currentName As String, ByVal finalName As String)
    DoCmd.Save kind, currentName
    DoCmd.Close kind, currentName, acSaveYes
    DoCmd.Rename finalName, kind, currentName
End Sub

Private Sub DropObject(ByVal kind As AcObjectType, ByVal objectName As String)
    On Error Resume Next
    DoCmd.Close kind, objectName, acSaveNo
    DoCmd.DeleteObject kind, objectName
    Err.Clear
End Sub

Private Sub SetTextProperty(ByVal propertyName As String, ByVal value As String)
    Dim prop As DAO.Property
    On Error Resume Next
    CurrentDb.Properties(propertyName) = value
    If Err.Number <> 0 Then
        Err.Clear
        Set prop = CurrentDb.CreateProperty(propertyName, dbText, value)
        CurrentDb.Properties.Append prop
    End If
End Sub
