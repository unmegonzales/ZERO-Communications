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
    CloseOpenDesign
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
    EMSEnsurePartners
    DoCmd.Echo True
    DoCmd.Hourglass False
    MsgBox "EMS is ready." & vbCrLf & "Close this database and open it again. The event list is the first screen." & vbCrLf & _
        "Use New event when there is no MyVenue file. Import replaces the sales on an event you already have open.", vbInformation
    Exit Sub
Failed:
    DoCmd.Echo True
    DoCmd.Hourglass False
    CloseOpenDesign
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
    frm.Section(acHeader).Height = 1320
    LabelForm frm.Name, "lblTitle", "Events", 200, 40, 4000, 320, 18, True
    LabelForm frm.Name, "lblHint", "The list is empty until you import the two MyVenue reports or choose New event. Double-click an event to open it.", 200, 380, 13400, 220, 9, False
    TextForm frm.Name, "txtFrom", 200, 900, 1500, 300, "From"
    TextForm frm.Name, "txtTo", 1900, 900, 1500, 300, "To"
    TextForm frm.Name, "txtFind", 3600, 900, 3400, 300, "Find"
    ButtonForm frm.Name, "cmdSearch", "Search", "=EMSHubSearch()", 7200, 880, 1300, 340
    ButtonForm frm.Name, "cmdNew", "New event", "=EMSNewEvent()", 8600, 880, 1500, 340
    ButtonForm frm.Name, "cmdImport", "Import reports", "=EMSImportNew()", 10200, 880, 1800, 340
    ButtonForm frm.Name, "cmdPayees", "Payees", "=EMSOpenPayees()", 12120, 880, 1400, 340
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstEvents", 200, 120, 13400, 6400)
    ctl.ColumnCount = 5
    ctl.ColumnHeads = True
    ctl.ColumnWidths = "0;1600;6200;2800;1800"
    ctl.BoundColumn = 1
    ctl.RowSource = "SELECT e.EventID, e.EventDate AS [Date], e.EventName AS [Event], e.EventType AS [Type], 0 AS [Amount due] FROM tblEvent AS e WHERE 1=0"
    ctl.OnDblClick = "=EMSOpenSelected()"
    frm.Section(acDetail).Height = 6800
    SaveObject acForm, frm.Name, "frmHub"
End Sub

Private Sub BuildEvent()
    Dim frm As Form
    Dim ctl As Control
    Set frm = CreateForm()
    frm.Caption = "Event assignments"
    frm.RecordSource = "tblEvent"
    frm.DefaultView = 0
    frm.Width = 16000
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.ScrollBars = 2
    frm.OnCurrent = "=EMSEventOpen()"
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 640
    frm.Section(acHeader).BackColor = RGB(232, 236, 240)
    frm.Section(acDetail).BackColor = vbWhite
    LabelForm frm.Name, "lblTitle", "Event assignments", 160, 120, 2800, 400, 18, True
    ButtonForm frm.Name, "cmdBack", "Events", "=EMSBackToEvents()", 3100, 120, 1100, 400
    ButtonForm frm.Name, "cmdImport", "Import", "=EMSImportCurrent()", 4320, 120, 1400, 400
    ButtonForm frm.Name, "cmdCalc", "Calculate", "=EMSCalculateCurrent()", 5840, 120, 1500, 400
    ButtonForm frm.Name, "cmdPreview", "Preview", "=EMSPreviewCurrent()", 7460, 120, 1300, 400
    ButtonForm frm.Name, "cmdPdf", "PDF", "=EMSExportCurrent()", 8880, 120, 1000, 400
    ButtonForm frm.Name, "cmdMail", "Email", "=EMSEmailCurrent()", 10000, 120, 1200, 400
    ButtonForm frm.Name, "cmdDelete", "Delete event", "=EMSDeleteEvent()", 13600, 120, 2000, 400
    BoundBox frm.Name, "EventID", "", 0, 0, 10, 10, False
    EventLine frm.Name, "EventDate", "Event date", 160, 80, 1800, False
    EventLine frm.Name, "EventName", "Event", 2120, 80, 6400, False
    EventLine frm.Name, "EventType", "Type", 8680, 80, 2800, False
    Set ctl = AddControl(frm.Name, acForm, acTextBox, acDetail, "txtDue", 11680, 300, 2400, 320)
    ctl.ControlSource = "=Nz(DSum('AmountDue','tblInvoice','EventID=' & Nz([EventID],0)),0)"
    ctl.Format = "Currency"
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(196, 202, 210)
    ctl.BackColor = vbWhite
    ctl.TextAlign = 3
    ctl.FontBold = True
    Paint ctl, 12, True
    LabelDetail frm.Name, "lblDue", "Amount due", 11680, 80, 2400, 200, 9, False
    EventLine frm.Name, "Venue", "Venue", 160, 720, 4200, False
    EventLine frm.Name, "MinimumCommission", "Minimum commission", 4520, 720, 2400, True
    EventLine frm.Name, "Notes", "Notes", 7080, 720, 7000, False
    LabelDetail frm.Name, "lblHow", "Import fills the stand nets. Choose a role and payee for each stand. A blank people count holds only that stand's tips.", 160, 1360, 15000, 240, 9, False
    LabelDetail frm.Name, "lblStands", "Stands", 160, 1660, 4000, 320, 16, True
    Set ctl = LinkSubform(frm.Name, "fsubAssign", "fsubAssign", 160, 2040, 15680, 4200)
    ctl.BorderStyle = 1
    LabelDetail frm.Name, "lblPick", "Selected stand", 160, 6340, 4000, 280, 14, True
    StandMoney frm.Name, "CookFee", "Cook fee", 160, 6680
    StandMoney frm.Name, "Bonus", "Bonus", 2100, 6680
    StandMoney frm.Name, "Shortage", "Shortage", 4040, 6680
    StandMoney frm.Name, "Cleaning", "Cleaning", 5980, 6680
    StandText frm.Name, "DepartmentName", "Department", 7920, 6680, 4200
    LabelDetail frm.Name, "lblVendor", "Vendor costs", 160, 7360, 4000, 280, 14, True
    Set ctl = LinkSubform(frm.Name, "fsubVendor", "fsubVendor", 160, 7700, 15680, 1100)
    ctl.BorderStyle = 1
    LabelDetail frm.Name, "lblSupp", "Adjustments", 160, 8920, 4000, 280, 14, True
    LabelDetail frm.Name, "lblStaff", "Staff tips", 8200, 8920, 4000, 280, 14, True
    Set ctl = LinkSubform(frm.Name, "fsubSupplement", "fsubSupplement", 160, 9260, 7800, 1100)
    ctl.BorderStyle = 1
    Set ctl = LinkSubform(frm.Name, "fsubStaff", "fsubStaff", 8200, 9260, 7640, 1100)
    ctl.BorderStyle = 1
    LabelDetail frm.Name, "lblInv", "Invoices", 160, 10480, 4000, 280, 14, True
    LabelDetail frm.Name, "lblExc", "Exceptions", 8200, 10480, 4000, 280, 14, True
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstInvoices", 160, 10820, 7800, 1100)
    ctl.RowSource = "SELECT i.InvoiceID, p.Abbreviation, i.InvoiceNumber, i.AmountDue FROM tblInvoice AS i INNER JOIN tblPayee AS p ON i.PayeeID=p.PayeeID WHERE i.EventID=[Forms]![frmEvent]![EventID] ORDER BY p.Abbreviation"
    ctl.ColumnCount = 4
    ctl.ColumnWidths = "0;1200;3200;1600"
    ctl.BoundColumn = 1
    ctl.BorderStyle = 1
    ctl.OnDblClick = "=EMSPreviewCurrent()"
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstExceptions", 8200, 10820, 7640, 1100)
    ctl.RowSource = "SELECT Kind, Detail FROM tblException WHERE EventID=[Forms]![frmEvent]![EventID] ORDER BY ExceptionID"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "1800;5600"
    ctl.BorderStyle = 1
    frm.Section(acDetail).Height = 12100
    SaveObject acForm, frm.Name, "frmEvent"
End Sub

Public Sub BuildEventScreen()
    On Error GoTo Failed
    Dim id As Variant
    DoCmd.Hourglass True
    EMSEnsurePayeeColumns
    CloseOpenDesign
    DropObject acForm, "frmEvent"
    DropObject acForm, "fsubAssign"
    DropObject acForm, "fsubVendor"
    DropObject acForm, "fsubSupplement"
    DropObject acForm, "fsubStaff"
    BuildAssign
    BuildVendor
    BuildSupplement
    BuildStaff
    BuildEvent
    DoCmd.Hourglass False
    id = Scalar("SELECT Max(EventID) FROM tblEvent")
    If IsNull(id) Then
        DoCmd.OpenForm "frmEvent"
    Else
        DoCmd.OpenForm "frmEvent", , , "EventID=" & id
    End If
    Exit Sub
Failed:
    DoCmd.Hourglass False
    CloseOpenDesign
    MsgBox "Build stopped." & vbCrLf & Err.Description, vbCritical
End Sub

' Removes the status and signature columns from the event table. The statement still prints the office names, dated with the event.
Public Sub TrimEventFields()
    On Error GoTo Failed
    DoCmd.Hourglass True
    CloseOpenDesign
    On Error Resume Next
    DoCmd.Close acTable, "tblEvent", acSaveNo
    DropEventColumn "StatusText"
    DropEventColumn "PreparedBy"
    DropEventColumn "PreparedTitle"
    DropEventColumn "PreparedDate"
    DropEventColumn "ReviewedBy"
    DropEventColumn "ReviewedTitle"
    DropEventColumn "ReviewedDate"
    DropEventColumn "ApprovedBy"
    DropEventColumn "ApprovedTitle"
    DropEventColumn "ApprovedDate"
    On Error GoTo Failed
    DropObject acForm, "frmEvent"
    BuildEvent
    DoCmd.Hourglass False
    MsgBox "The extra event columns are removed. Payout statements still print the office names, and those dates are the event date.", vbInformation
    Exit Sub
Failed:
    DoCmd.Hourglass False
    CloseOpenDesign
    MsgBox "Could not remove the event columns." & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub DropEventColumn(ByVal fieldName As String)
    On Error Resume Next
    CurrentDb.Execute "ALTER TABLE [tblEvent] DROP COLUMN [" & fieldName & "]"
    CurrentDb.TableDefs.Refresh
End Sub

Private Sub BuildPayee()
    Dim frm As Form
    Dim ctl As Control
    Set frm = CreateForm()
    frm.Caption = "Payees"
    frm.RecordSource = "tblPayee"
    frm.Width = 14800
    frm.RecordSelectors = False
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    frm.OnCurrent = "=EMSPayeeCurrent()"
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 640
    frm.Section(acHeader).BackColor = RGB(232, 236, 240)
    frm.Section(acDetail).BackColor = RGB(232, 236, 240)
    LabelForm frm.Name, "lblTitle", "Payees", 240, 140, 2200, 400, 20, True
    Set ctl = AddControl(frm.Name, acForm, acCheckBox, acHeader, "chkShowInactive", 2800, 200, 280, 280)
    ctl.DefaultValue = "False"
    ctl.OnClick = "=EMSPayeeFilter()"
    LabelForm frm.Name, "lblInactive", "Show inactive", 3160, 200, 2200, 280, 10, False
    ButtonForm frm.Name, "cmdHome", "Home", "=EMSBackToEvents()", 7800, 140, 1400, 400
    ButtonForm frm.Name, "cmdBulk", "Bulk edit vendors", "=EMSBulkPayees()", 9360, 140, 2500, 400
    ButtonForm frm.Name, "cmdNew", "New payee", "=GoToNewPayee()", 12020, 140, 1800, 400
    EMSEnsurePayeeColumns
    EMSDropPayeeQuery
    ListHead frm.Name, "lblHeadID", "ID", 240, 700
    ListHead frm.Name, "lblHeadPayee", "Payee", 940, 4600
    ListHead frm.Name, "lblHeadType", "Type", 5540, 1100
    ListHead frm.Name, "lblHeadAbbrev", "Abbrev", 6640, 1300
    ListHead frm.Name, "lblHeadSAP", "SAP ID", 7940, 1900
    ListHead frm.Name, "lblHeadBSS", "BSS ID", 9840, 1500
    ListHead frm.Name, "lblHeadStatus", "Status", 11340, 1400
    Set ctl = AddControl(frm.Name, acForm, acListBox, acDetail, "lstPayees", 240, 300, 14320, 2460)
    ctl.ColumnCount = 7
    ctl.ColumnHeads = False
    ctl.ColumnWidths = "700;4600;1100;1300;1900;1500;1400"
    ctl.BoundColumn = 1
    ctl.RowSourceType = "EMSPayeeList"
    ctl.RowSource = ""
    ctl.OnClick = "=EMSPayeePick()"
    ctl.OnDblClick = "=EMSPayeeEdit()"
    ctl.BackColor = vbWhite
    Paint ctl, 10, False
    ButtonFormDetail frm.Name, "cmdEdit", "Edit", "=EMSPayeeEdit()", 240, 3080, 1500, 400
    ButtonFormDetail frm.Name, "cmdSave", "Save", "=EMSPayeeSave()", 1860, 3080, 1500, 400
    ButtonFormDetail frm.Name, "cmdCancel", "Cancel", "=EMSPayeeCancel()", 3480, 3080, 1500, 400
    LabelDetail frm.Name, "lblSelected", "Selected payee", 240, 3600, 4000, 300, 14, True
    PayeeField frm.Name, "txtPayeeName", "PayeeName", "Payee name", 240, 3980, 6200, False
    PayeeCombo frm.Name, "cboType", "PayeeType", "Type", 6600, 3980, 1500, "NPO;SUB"
    PayeeField frm.Name, "txtAbbreviation", "Abbreviation", "Abbreviation", 8260, 3980, 2000, False
    Set ctl = AddControl(frm.Name, acForm, acCheckBox, acDetail, "chkActive", 10520, 4220, 300, 300)
    ctl.ControlSource = "Active"
    ctl.DefaultValue = "True"
    LabelDetail frm.Name, "lblActive", "Active", 10900, 4220, 1400, 300, 11, False
    PayeeField frm.Name, "txtSAPID", "SAPID", "SAP ID", 240, 4740, 2200, False
    PayeeField frm.Name, "txtBSSID", "BSSID", "BSS ID", 2600, 4740, 2200, False
    PayeeField frm.Name, "txtAddress1", "Address1", "Address", 4960, 4740, 4600, False
    PayeeField frm.Name, "txtAddress2", "Address2", "City, state, ZIP", 9720, 4740, 4840, False
    PayeeField frm.Name, "txtNotes", "Notes", "Notes", 240, 5520, 14320, True
    LabelDetail frm.Name, "lblEmails", "Emails", 240, 6560, 3000, 280, 12, True
    LabelDetail frm.Name, "lblRates", "Commission rates", 7600, 6560, 4000, 280, 12, True
    LabelDetail frm.Name, "lblRateHint", "Type 10 for 10 percent. Leave Ends blank while the rate is current.", 7600, 6860, 6800, 240, 9, False
    Set ctl = LinkSubform(frm.Name, "fsubEmail", "fsubEmail", 240, 7280, 7000, 2200)
    ctl.BorderStyle = 1
    Set ctl = LinkSubform(frm.Name, "fsubRate", "fsubRate", 7600, 7280, 6960, 2200)
    ctl.BorderStyle = 1
    LabelDetail frm.Name, "lblStatus", "Select a payee in the list. Edit unlocks the fields.", 240, 9600, 14320, 280, 10, False
    frm.Section(acDetail).Height = 10000
    SaveObject acForm, frm.Name, "frmPayee"
End Sub

Public Sub BuildPayeeScreen()
    On Error GoTo Failed
    DoCmd.Hourglass True
    CloseOpenDesign
    EMSEnsureIceColumn
    EMSEnsureEmailColumns
    DropObject acForm, "frmPayee"
    DropObject acForm, "fsubEmail"
    DropObject acForm, "fsubRate"
    DropObject acForm, "fsubVendor"
    BuildEmail
    BuildRate
    BuildVendor
    BuildPayee
    EMSEnsurePartners
    DoCmd.Hourglass False
    DoCmd.OpenForm "frmPayee"
    Exit Sub
Failed:
    DoCmd.Hourglass False
    CloseOpenDesign
    MsgBox "Build stopped." & vbCrLf & Err.Description, vbCritical
End Sub

Public Function GoToNewPayee()
    DoCmd.GoToRecord acDataForm, "frmPayee", acNewRec
    Forms!frmPayee!txtPayeeName.SetFocus
End Function

Private Sub StandMoney(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal rowTop As Long)
    StandBind formName, fieldName, caption, x, rowTop, 1800, True
End Sub

Private Sub StandText(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal rowTop As Long, ByVal width As Long)
    StandBind formName, fieldName, caption, x, rowTop, width, False
End Sub

Private Sub StandBind(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal rowTop As Long, ByVal width As Long, ByVal money As Boolean)
    Dim ctl As Control
    LabelDetail formName, "lbl" & fieldName, caption, x, rowTop, width, 200, 9, False
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, "txt" & fieldName, x, rowTop + 220, width, 320)
    ctl.ControlSource = "=[fsubAssign].[Form]![col" & fieldName & "]"
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(196, 202, 210)
    ctl.BackColor = vbWhite
    If money Then ctl.Format = "Currency"
    Paint ctl, 11, False
End Sub

Private Sub EventLine(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal rowTop As Long, ByVal width As Long, ByVal money As Boolean)
    Dim ctl As Control
    LabelDetail formName, "lbl" & fieldName, caption, x, rowTop, width, 200, 9, False
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, "txt" & fieldName, x, rowTop + 220, width, 320)
    ctl.ControlSource = fieldName
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(196, 202, 210)
    ctl.BackColor = vbWhite
    If fieldName = "EventDate" Then ctl.Format = "mm/dd/yyyy"
    If money Then ctl.Format = "Currency"
    Paint ctl, 11, False
End Sub

Private Sub GridHead(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal width As Long)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acLabel, acHeader, controlName, x, 40, width, 280)
    ctl.Caption = caption
    Paint ctl, 9, True
End Sub

Private Sub GridCell(ByVal formName As String, ByVal fieldName As String, ByVal x As Long, ByVal width As Long, ByVal money As Boolean)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, "col" & fieldName, x, 40, width, 320)
    ctl.ControlSource = fieldName
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(196, 202, 210)
    ctl.BackColor = vbWhite
    If money Then
        ctl.Format = "Currency"
        ctl.TextAlign = 3
    End If
    Paint ctl, 10, False
End Sub

Private Sub GridCombo(ByVal formName As String, ByVal fieldName As String, ByVal x As Long, ByVal width As Long, ByVal sourceType As String, ByVal source As String, ByVal columns As Long, ByVal widths As String)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acComboBox, acDetail, "col" & fieldName, x, 40, width, 320)
    ctl.ControlSource = fieldName
    ctl.RowSourceType = sourceType
    ctl.RowSource = source
    ctl.ColumnCount = columns
    ctl.ColumnWidths = widths
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(196, 202, 210)
    ctl.BackColor = vbWhite
    Paint ctl, 10, False
    ctl.ListRows = 16
End Sub

Private Function OpenGrid(ByVal caption As String, ByVal source As String, ByVal width As Long) As Form
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = caption
    frm.RecordSource = source
    frm.DefaultView = 1
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    frm.DividingLines = True
    frm.Width = width
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 360
    frm.Section(acHeader).BackColor = RGB(232, 236, 240)
    frm.Section(acDetail).Height = 400
    frm.Section(acDetail).BackColor = vbWhite
    On Error Resume Next
    frm.Section(acDetail).AlternateBackColor = RGB(245, 247, 250)
    On Error GoTo 0
    Set OpenGrid = frm
End Function

Private Sub HideKey(ByVal formName As String, ByVal fieldName As String)
    GridCell formName, fieldName, 0, 10, False
    Forms(formName).Controls("col" & fieldName).Visible = False
    Forms(formName).Controls("col" & fieldName).TabStop = False
End Sub

Private Sub BuildAssign()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Stands"
    frm.RecordSource = "tblAssign"
    frm.DefaultView = 1
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    frm.DividingLines = True
    frm.Width = 15680
    EnsureFormHeader frm.Name
    frm.Section(acHeader).Height = 360
    frm.Section(acHeader).BackColor = RGB(232, 236, 240)
    frm.Section(acDetail).Height = 400
    frm.Section(acDetail).BackColor = vbWhite
    On Error Resume Next
    frm.Section(acDetail).AlternateBackColor = RGB(245, 247, 250)
    On Error GoTo 0
    frm.Width = 15000
    GridHead frm.Name, "lblLocation", "Location", 40, 2600
    GridHead frm.Name, "lblRole", "Role", 2680, 1200
    GridHead frm.Name, "lblPayee", "Payee", 3920, 1600
    GridHead frm.Name, "lblFood", "Food net", 5560, 1200
    GridHead frm.Name, "lblNA", "NA net", 6800, 1100
    GridHead frm.Name, "lblAlcohol", "Alcohol net", 7940, 1300
    GridHead frm.Name, "lblTips", "Tips", 9280, 1100
    GridHead frm.Name, "lblPeople", "People", 10420, 800
    GridHead frm.Name, "lblNames", "Staff names", 11260, 2200
    GridCell frm.Name, "AssignID", 0, 10, False
    GridCell frm.Name, "EventID", 0, 10, False
    Forms(frm.Name).Controls("colAssignID").Visible = False
    Forms(frm.Name).Controls("colEventID").Visible = False
    Forms(frm.Name).Controls("colAssignID").TabStop = False
    Forms(frm.Name).Controls("colEventID").TabStop = False
    GridCell frm.Name, "LocationName", 40, 2600, False
    GridCombo frm.Name, "Role", 2680, 1200, "Value List", "NPO;SUB;STAFF;OPEN", 1, "1200"
    GridCombo frm.Name, "PayeeID", 3920, 1600, "Table/Query", "SELECT PayeeID, Abbreviation FROM tblPayee WHERE [Active]<>0 ORDER BY Abbreviation", 2, "0;1400"
    GridCell frm.Name, "FoodNet", 5560, 1200, True
    GridCell frm.Name, "NaNet", 6800, 1100, True
    GridCell frm.Name, "AlcoholNet", 7940, 1300, True
    GridCell frm.Name, "PosTips", 9280, 1100, True
    GridCell frm.Name, "Headcount", 10420, 800, False
    GridCell frm.Name, "StaffNames", 11260, 2200, False
    HideKey frm.Name, "CookFee"
    HideKey frm.Name, "Bonus"
    HideKey frm.Name, "Shortage"
    HideKey frm.Name, "Cleaning"
    HideKey frm.Name, "DepartmentName"
    ButtonFormDetail frm.Name, "cmdRemove", "Remove", "=EMSRemoveStand()", 13520, 40, 1300, 300
    SaveObject acForm, frm.Name, "fsubAssign"
End Sub

Private Sub BuildVendor()
    Dim frm As Form
    Set frm = OpenGrid("Vendor costs", "tblVendorInput", 13800)
    HideKey frm.Name, "EventID"
    GridHead frm.Name, "lblVendor", "Vendor", 40, 2200
    GridHead frm.Name, "lblCash", "Cash", 2280, 1400
    GridHead frm.Name, "lblCards", "Cards", 3720, 1400
    GridHead frm.Name, "lblCorp", "Corporate voucher", 5160, 1900
    GridHead frm.Name, "lblLevy", "Levy voucher", 7100, 1700
    GridHead frm.Name, "lblUnm", "UNM voucher", 8840, 1700
    GridHead frm.Name, "lblIce", "Ice", 10580, 1400
    GridHead frm.Name, "lblOther", "Other", 12020, 1400
    GridCombo frm.Name, "PayeeID", 40, 2200, "Table/Query", "SELECT PayeeID, Abbreviation FROM tblPayee WHERE PayeeType='SUB' AND [Active]<>0 ORDER BY Abbreviation", 2, "0;2000"
    GridCell frm.Name, "CashTender", 2280, 1400, True
    GridCell frm.Name, "CardTender", 3720, 1400, True
    GridCell frm.Name, "CorporateVoucher", 5160, 1900, True
    GridCell frm.Name, "LevyNfpVoucher", 7100, 1700, True
    GridCell frm.Name, "UnmGuestVoucher", 8840, 1700, True
    GridCell frm.Name, "IceAmount", 10580, 1400, True
    GridCell frm.Name, "OtherExpense", 12020, 1400, True
    SaveObject acForm, frm.Name, "fsubVendor"
End Sub

Private Sub BuildSupplement()
    Dim frm As Form
    Set frm = OpenGrid("Adjustments", "tblSupplement", 7600)
    HideKey frm.Name, "EventID"
    GridHead frm.Name, "lblPayee", "Payee", 40, 1600
    GridHead frm.Name, "lblLabel", "Label", 1680, 2800
    GridHead frm.Name, "lblAmount", "Amount", 4520, 1500
    GridHead frm.Name, "lblGL", "GL", 6060, 1400
    GridCombo frm.Name, "PayeeID", 40, 1600, "Table/Query", "SELECT PayeeID, Abbreviation FROM tblPayee WHERE [Active]<>0 ORDER BY Abbreviation", 2, "0;1400"
    GridCell frm.Name, "Label", 1680, 2800, False
    GridCell frm.Name, "Amount", 4520, 1500, True
    GridCell frm.Name, "GLCode", 6060, 1400, False
    SaveObject acForm, frm.Name, "fsubSupplement"
End Sub

Private Sub BuildStaff()
    Dim frm As Form
    Set frm = OpenGrid("Staff tips", "tblStaffTip", 7500)
    frm.AllowAdditions = False
    frm.AllowEdits = False
    frm.AllowDeletions = False
    HideKey frm.Name, "EventID"
    GridHead frm.Name, "lblStand", "Stand", 40, 2000
    GridHead frm.Name, "lblSlot", "#", 2080, 500
    GridHead frm.Name, "lblPerson", "Person", 2620, 1600
    GridHead frm.Name, "lblAmount", "Amount", 4260, 1000
    GridHead frm.Name, "lblPool", "Pool", 5300, 1000
    GridHead frm.Name, "lblHeld", "Held", 6340, 1000
    GridCell frm.Name, "LocationName", 40, 2000, False
    GridCell frm.Name, "SlotNo", 2080, 500, False
    GridCell frm.Name, "PersonName", 2620, 1600, False
    GridCell frm.Name, "Amount", 4260, 1000, True
    GridCell frm.Name, "PoolAmount", 5300, 1000, True
    GridCell frm.Name, "HeldBack", 6340, 1000, True
    SaveObject acForm, frm.Name, "fsubStaff"
End Sub

Private Sub BuildEmail()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Emails"
    frm.RecordSource = "tblPayeeEmail"
    UseDatasheet frm
    DatasheetBox frm.Name, "EmailAddress", "Email", 4200
    DatasheetBox frm.Name, "Label", "Label", 1600
    DatasheetCheck frm.Name, "IsPrimary", "Primary", 1000
    SaveObject acForm, frm.Name, "fsubEmail"
End Sub

Private Sub BuildRate()
    Dim frm As Form
    Set frm = CreateForm()
    frm.Caption = "Rates"
    frm.RecordSource = "tblPayeeRate"
    frm.Filter = "RateCode<>'ICE_PER_BAG'"
    frm.FilterOn = True
    UseDatasheet frm
    DatasheetBox frm.Name, "RateCode", "Rate", 2400
    DatasheetBox frm.Name, "RateAmount", "Amount", 1200
    DatasheetBox frm.Name, "EffectiveDate", "Starts", 1400
    DatasheetBox frm.Name, "EndDate", "Ends", 1400
    DatasheetBox frm.Name, "Notes", "Notes", 2200
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
    names = Array("txtSumFood", "txtSumAlc", "txtSumGroup", "txtSumGross")
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
    FieldCaption "tblEvent", "MinimumCommission", "Minimum commission"
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
    FieldCaption "tblVendorInput", "IceAmount", "Ice"
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
    Set ButtonForm = PlaceButton(formName, controlName, caption, onClick, acHeader, x, y, width, height)
End Function

Private Function ButtonFormDetail(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal onClick As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    Set ButtonFormDetail = PlaceButton(formName, controlName, caption, onClick, acDetail, x, y, width, height)
End Function

Private Function PlaceButton(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal onClick As String, ByVal section As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    Set PlaceButton = AddControl(formName, acForm, acCommandButton, section, controlName, x, y, width, height)
    PlaceButton.Caption = caption
    PlaceButton.OnClick = onClick
    Paint PlaceButton, 10, False
End Function

Private Sub PayeeField(ByVal formName As String, ByVal controlName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal tall As Boolean)
    Dim ctl As Control
    Dim boxHeight As Long
    boxHeight = IIf(tall, 560, 340)
    LabelDetail formName, "lbl" & controlName, caption, x, y, width, 220, 9, False
    Set ctl = AddControl(formName, acForm, acTextBox, acDetail, controlName, x, y + 240, width, boxHeight)
    ctl.ControlSource = fieldName
    ctl.BorderStyle = 1
    ctl.BackColor = vbWhite
    If tall Then ctl.ScrollBars = 2
    Paint ctl, 11, False
End Sub

Private Sub PayeeCombo(ByVal formName As String, ByVal controlName As String, ByVal fieldName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal list As String)
    Dim ctl As Control
    LabelDetail formName, "lbl" & controlName, caption, x, y, width, 220, 9, False
    Set ctl = AddControl(formName, acForm, acComboBox, acDetail, controlName, x, y + 240, width, 340)
    ctl.ControlSource = fieldName
    ctl.RowSourceType = "Value List"
    ctl.RowSource = list
    ctl.LimitToList = True
    ctl.BorderStyle = 1
    ctl.BackColor = vbWhite
    Paint ctl, 11, False
End Sub

Private Sub LabelForm(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal size As Single, ByVal bold As Boolean)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acLabel, acHeader, controlName, x, y, width, height)
    ctl.Caption = caption
    Paint ctl, size, bold
End Sub

Private Sub ListHead(ByVal formName As String, ByVal controlName As String, ByVal caption As String, ByVal x As Long, ByVal width As Long)
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acLabel, acDetail, controlName, x, 40, width, 240)
    ctl.Caption = caption
    Paint ctl, 9, True
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

Private Function LinkSubform(ByVal formName As String, ByVal controlName As String, ByVal source As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Control
    Dim ctl As Control
    Set ctl = AddControl(formName, acForm, acSubform, acDetail, controlName, x, y, width, height)
    Set LinkSubform = ctl
    ctl.SourceObject = source
    ctl.LinkMasterFields = "PayeeID"
    ctl.LinkChildFields = "PayeeID"
    If Left$(source, 4) <> "fsub" Or source = "fsubAssign" Or source = "fsubVendor" Or source = "fsubSupplement" Or source = "fsubStaff" Then
        If source <> "fsubEmail" And source <> "fsubRate" Then
            ctl.LinkMasterFields = "EventID"
            ctl.LinkChildFields = "EventID"
        End If
    End If
End Function

Private Sub DatasheetBox(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal width As Long)
    DatasheetControl formName, acTextBox, fieldName, caption, width, "", "", 1, ""
End Sub

Private Sub DatasheetCheck(ByVal formName As String, ByVal fieldName As String, ByVal caption As String, ByVal width As Long)
    DatasheetControl formName, acCheckBox, fieldName, caption, width, "", "", 1, ""
End Sub

Private Sub DatasheetCombo(ByVal formName As String, ByVal fieldName As String, ByVal sourceType As String, ByVal source As String, ByVal columns As Long, ByVal widths As String, ByVal caption As String, ByVal width As Long)
    DatasheetControl formName, acComboBox, fieldName, caption, width, sourceType, source, columns, widths
End Sub

' The column title in a datasheet is the attached label. AddControl removes that label,
' which is why the grid was showing names like txtEmailAddress.
Private Sub DatasheetControl(ByVal formName As String, ByVal controlKind As Long, ByVal fieldName As String, ByVal caption As String, ByVal width As Long, ByVal sourceType As String, ByVal source As String, ByVal columns As Long, ByVal widths As String)
    Dim host As Object
    Dim ctl As Control
    Dim lbl As Control
    Dim before As Long
    Dim index As Long
    Dim kept As String
    Dim found As Boolean
    Set host = Forms(formName)
    before = host.Controls.Count
    Set ctl = CreateControl(formName, controlKind, acDetail, "", fieldName, 0, 0, width, 300)
    kept = "col" & fieldName
    ctl.Name = kept
    ctl.ControlSource = fieldName
    On Error Resume Next
    ctl.ColumnWidth = width
    ctl.ColumnHidden = False
    On Error GoTo 0
    If controlKind = acComboBox Then
        ctl.RowSourceType = sourceType
        ctl.RowSource = source
        ctl.ColumnCount = columns
        ctl.ColumnWidths = widths
        ctl.BoundColumn = 1
        ctl.LimitToList = True
    End If
    If InStr(fieldName, "Date") > 0 Then ctl.Format = "Short Date"
    If fieldName = "RateAmount" Then
        ctl.Format = "0%"
        ctl.AfterUpdate = "=EMSRateTyped()"
    ElseIf IsMoneyField(fieldName) Then
        ctl.Format = "Currency"
    End If
    found = False
    For index = host.Controls.Count - 1 To before Step -1
        If host.Controls(index).Name <> kept Then
            host.Controls(index).Caption = caption
            host.Controls(index).Name = "lbl" & fieldName
            found = True
            Exit For
        End If
    Next index
    If found Then Exit Sub
    Set lbl = CreateControl(formName, acLabel, acDetail, ctl.Name, "", 0, 0, width, 200)
    lbl.Caption = caption
    lbl.Name = "lbl" & fieldName
End Sub

Private Function IsMoneyField(ByVal fieldName As String) As Boolean
    Select Case fieldName
        Case "CashTender", "CardTender", "CorporateVoucher", "LevyNfpVoucher", "UnmGuestVoucher", "IceAmount", "OtherExpense", "Amount", "PoolAmount", "HeldBack", "CookFee", "Bonus", "Shortage", "Cleaning", "PosSales", "FoodNet", "NaNet", "AlcoholNet", "PosTips"
            IsMoneyField = True
    End Select
End Function

Private Function ColumnToken(ByVal caption As String, ByVal fieldName As String) As String
    Dim result As String
    Dim index As Long
    Dim ch As String
    result = ""
    For index = 1 To Len(caption)
        ch = Mid$(caption, index, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then result = result & ch
    Next index
    If Len(result) = 0 Then result = "col" & fieldName
    ColumnToken = result
End Function

Private Sub UseDatasheet(ByVal frm As Form)
    frm.DefaultView = 2
    frm.DatasheetFontName = "Calibri"
    frm.DatasheetFontHeight = 10
    On Error Resume Next
    frm.DatasheetBackColor = vbWhite
    frm.DatasheetForeColor = RGB(24, 49, 78)
    frm.DatasheetGridlinesColor = RGB(196, 202, 210)
    frm.DatasheetAlternateBackColor = RGB(245, 247, 250)
    On Error GoTo 0
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

Private Sub CloseOpenDesign()
    Dim index As Long
    On Error Resume Next
    For index = Reports.Count - 1 To 0 Step -1
        DoCmd.Close acReport, Reports(index).Name, acSaveNo
    Next index
    For index = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(index).Name, acSaveNo
    Next index
End Sub

Private Sub DropObject(ByVal kind As AcObjectType, ByVal objectName As String)
    On Error Resume Next
    DoCmd.Close kind, objectName, acSaveNo
    DoCmd.DeleteObject kind, objectName
    Err.Clear
End Sub

Private Sub SetTextProperty(ByVal propertyName As String, ByVal value As String)
    Dim prop As Object
    On Error Resume Next
    CurrentDb.Properties(propertyName) = value
    If Err.Number <> 0 Then
        Err.Clear
        Set prop = CurrentDb.CreateProperty(propertyName, EMS_TEXT, value)
        CurrentDb.Properties.Append prop
    End If
End Sub
