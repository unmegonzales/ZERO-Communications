Attribute VB_Name = "EMS_Calculate"
Option Compare Database
Option Explicit

' A blank staff headcount holds only that stand's tips. Donation and vendor
' checks still calculate. A new event can be typed with no MyVenue file;
' stands added by hand use the nets on the row.

Public Sub EMSCalculateCurrent()
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Sub
    If Forms!frmEvent.Dirty Then Forms!frmEvent.Dirty = False
    EMSCalculateEvent CLng(Nz(Forms!frmEvent!EventID, 0))
    EMSEventRefresh
End Sub

Public Sub EMSCalculateEvent(ByVal eventId As Long)
    On Error GoTo Failed
    Dim eventDay As Date
    Dim minimum As Currency
    Dim rs As Object
    Dim groups As Object
    Dim key As Variant
    If eventId = 0 Then
        MsgBox "Save the event first.", vbInformation
        Exit Sub
    End If
    If IsNull(Scalar("SELECT EventDate FROM tblEvent WHERE EventID=" & eventId)) Then
        MsgBox "Enter the event date before calculating.", vbInformation
        Exit Sub
    End If
    eventDay = CDate(Scalar("SELECT EventDate FROM tblEvent WHERE EventID=" & eventId))
    minimum = CCur(Nz(Scalar("SELECT MinimumCommission FROM tblEvent WHERE EventID=" & eventId), 200))
    Exec "DELETE FROM tblInvoiceLine WHERE InvoiceID In (SELECT InvoiceID FROM tblInvoice WHERE EventID=" & eventId & ")"
    Exec "DELETE FROM tblInvoice WHERE EventID=" & eventId
    Exec "DELETE FROM tblStaffTip WHERE EventID=" & eventId
    Exec "DELETE FROM tblException WHERE EventID=" & eventId & " AND Kind In ('Headcount','Unassigned','No payee','Staff names','Alcohol rate')"
    Exec "DELETE FROM tblSupplement WHERE EventID=" & eventId & " AND Label In ('CREDIT CARD FEES','ICE','SHORTAGE','OTHER EXPENSE')"
    Set groups = CreateObject("Scripting.Dictionary")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblAssign WHERE EventID=" & eventId & " ORDER BY LocationName", EMS_SNAPSHOT)
    Do While Not rs.EOF
        Accumulate rs, eventId, eventDay, minimum, groups
        rs.MoveNext
    Loop
    rs.Close
    For Each key In groups.Keys
        WriteInvoice eventId, eventDay, CLng(key), groups(key)
    Next key
    MsgBox "Calculation finished. Open stands and blank staff counts are listed under exceptions and were left unpaid. Everything else was calculated.", vbInformation
    Exit Sub
Failed:
    MsgBox "Calculate stopped. Run Calculate again and it will rebuild this event." & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub Accumulate(ByVal rs As Object, ByVal eventId As Long, ByVal eventDay As Date, ByVal minimum As Currency, ByVal groups As Object)
    Dim role As String
    Dim tips As Currency
    Dim sales As Currency
    role = RoleOf(CStr(Nz(rs!role, "")))
    tips = CCur(Nz(rs!PosTips, 0))
    sales = CCur(Nz(rs!PosSales, 0))
    If role = "STAFF" Then
        SplitStaff rs, eventId, tips
        Exit Sub
    End If
    If role = "OPEN" Then
        If tips <> 0 Or sales <> 0 Or CCur(Nz(rs!FoodNet, 0)) <> 0 Or CCur(Nz(rs!AlcoholNet, 0)) <> 0 Then
            LogException eventId, 0, "Unassigned", CStr(rs!LocationName) & " still has " & MoneyText(tips) & " in tips and was left unpaid. Choose NPO, SUB, or STAFF."
        End If
        Exit Sub
    End If
    If IsNull(rs!PayeeID) Then
        LogException eventId, 0, "No payee", CStr(rs!LocationName) & " is marked " & role & " and has no payee. That stand was skipped."
        Exit Sub
    End If
    AddLine rs, eventId, eventDay, minimum, role, CLng(rs!PayeeID), groups
End Sub

Private Sub SplitStaff(ByVal rs As Object, ByVal eventId As Long, ByVal tips As Currency)
    Dim people As Long
    Dim names As Collection
    Dim cents As Variant
    Dim eachPay As Currency
    Dim held As Currency
    Dim index As Long
    If tips = 0 Then Exit Sub
    people = CLng(Nz(rs!Headcount, 0))
    If people < 1 Then
        LogException eventId, 0, "Headcount", CStr(rs!LocationName) & " has " & MoneyText(tips) & " in tips and the people count is blank. Those tips stay on this stand until you enter the count and calculate again. No other stand is held."
        Exit Sub
    End If
    Set names = StaffList(CStr(Nz(rs!StaffNames, "")), people, eventId, CStr(rs!LocationName))
    cents = Int(CDec(tips) * CDec(100))
    eachPay = CCur(Int(cents / CDec(people)) / 100)
    held = tips - (eachPay * people)
    For index = 1 To people
        Exec "INSERT INTO tblStaffTip (EventID, LocationName, SlotNo, PersonName, Amount, PoolAmount, HeldBack) VALUES (" & _
            eventId & "," & SqlText(Left$(CStr(rs!LocationName), 150)) & "," & index & "," & SqlText(Left$(CStr(names(index)), 80)) & "," & _
            eachPay & "," & tips & "," & IIf(index = 1, held, 0) & ")"
    Next index
End Sub

Private Function StaffList(ByVal raw As String, ByVal people As Long, ByVal eventId As Long, ByVal locationName As String) As Collection
    Dim lines As Variant
    Dim index As Long
    Dim name As String
    Dim typed As New Collection
    Dim result As New Collection
    raw = Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf)
    If Len(Trim$(raw)) > 0 Then
        lines = Split(raw, vbLf)
        For index = LBound(lines) To UBound(lines)
            name = Trim$(CStr(lines(index)))
            If Len(name) > 0 Then typed.Add name
        Next index
    End If
    If typed.Count = 0 Then
        For index = 1 To people
            result.Add "Associate " & index
        Next index
    Else
        If typed.Count <> people Then
            LogException eventId, 0, "Staff names", locationName & " has " & typed.Count & " name(s) and a people count of " & people & ". The split uses the people count."
        End If
        For index = 1 To people
            If index <= typed.Count Then result.Add typed(index) Else result.Add "Associate " & index
        Next index
    End If
    Set StaffList = result
End Function

Private Sub AddLine(ByVal rs As Object, ByVal eventId As Long, ByVal eventDay As Date, ByVal minimum As Currency, ByVal role As String, ByVal payeeId As Long, ByVal groups As Object)
    Dim food As Currency, na As Currency, beer As Currency, liquor As Currency
    Dim fromFile As Boolean
    Dim foodRate As Double, alcoholRate As Double
    Dim foodComm As Currency, alcoholComm As Currency, commission As Currency
    Dim lineFood As Currency, lineAlcohol As Currency
    Dim topUp As Currency, tips As Currency, payable As Currency
    Dim adjustment As Currency
    Dim line As Object
    Dim key As String
    StandNets rs, food, na, beer, liquor, fromFile
    If role = "NPO" Then
        foodRate = RateOn(payeeId, "DONATION_FOOD", eventDay, 0.1)
        alcoholRate = RateOn(payeeId, "DONATION_ALCOHOL", eventDay, 0.08)
        foodComm = MoneyProduct(food + na, foodRate)
        alcoholComm = MoneyProduct(beer + liquor, alcoholRate)
        lineFood = food + na
        lineAlcohol = beer + liquor
    Else
        foodRate = RateOn(payeeId, "FOOD", eventDay, 0)
        foodComm = MoneyProduct(food, foodRate)
        alcoholComm = MoneyProduct(na, RateOn(payeeId, "NA", eventDay, 0)) + SubAlcoholCommission(payeeId, eventDay, eventId, CStr(rs!LocationName), beer, liquor, fromFile, alcoholRate)
        lineFood = food
        lineAlcohol = na + beer + liquor
    End If
    commission = foodComm + alcoholComm
    topUp = 0
    If role = "NPO" And (food + na + beer + liquor) > 0 And commission < minimum Then topUp = minimum - commission
    tips = CCur(Nz(rs!PosTips, 0))
    payable = FloorProduct(tips, 1 - RateOn(payeeId, "TIP_WITHHOLD", eventDay, 0.1))
    adjustment = CCur(Nz(rs!CookFee, 0)) + CCur(Nz(rs!Bonus, 0)) + topUp - CCur(Nz(rs!Shortage, 0)) - CCur(Nz(rs!Cleaning, 0))
    Set line = CreateObject("Scripting.Dictionary")
    line.Add "Location", CStr(rs!LocationName)
    line.Add "Net", food + na + beer + liquor
    line.Add "FoodNet", lineFood
    line.Add "FoodCommission", foodComm
    line.Add "AlcoholNet", lineAlcohol
    line.Add "AlcoholCommission", alcoholComm
    line.Add "Commission", commission
    line.Add "Gratuity", payable
    line.Add "Adjustments", adjustment
    line.Add "Payout", commission + payable + adjustment
    line.Add "Kind", role
    Dim bucket As Collection
    key = CStr(payeeId)
    If Not groups.Exists(key) Then groups.Add key, New Collection
    Set bucket = groups(key)
    bucket.Add line
End Sub

Private Function SubAlcoholCommission(ByVal payeeId As Long, ByVal eventDay As Date, ByVal eventId As Long, ByVal locationName As String, ByVal beer As Currency, ByVal liquor As Currency, ByVal fromFile As Boolean, ByRef labelRate As Double) As Currency
    Dim beerRate As Double
    Dim liquorRate As Double
    beerRate = RateOn(payeeId, "BEER", eventDay, 0)
    liquorRate = RateOn(payeeId, "LIQUOR", eventDay, 0)
    If fromFile Then
        labelRate = beerRate
        SubAlcoholCommission = MoneyProduct(beer, beerRate) + MoneyProduct(liquor, liquorRate)
        If Abs(beerRate - liquorRate) > 0.0000001 And beer <> 0 And liquor <> 0 Then labelRate = -1
    ElseIf Abs(beerRate - liquorRate) <= 0.0000001 Then
        labelRate = beerRate
        SubAlcoholCommission = MoneyProduct(liquor, beerRate)
    Else
        labelRate = beerRate
        SubAlcoholCommission = MoneyProduct(liquor, beerRate)
        If liquor <> 0 Then LogException eventId, 0, "Alcohol rate", locationName & " has no price report, and the beer and liquor rates differ. The typed alcohol net used the beer rate."
    End If
End Function

Private Sub StandNets(ByVal rs As Object, ByRef food As Currency, ByRef na As Currency, ByRef beer As Currency, ByRef liquor As Currency, ByRef fromFile As Boolean)
    Dim sales As Object
    food = 0
    na = 0
    beer = 0
    liquor = 0
    Set sales = CurrentDb.OpenRecordset("SELECT Bucket, Sum(NetSales) AS NetAmount FROM tblSale WHERE EventID=" & rs!EventID & " AND LocationName=" & SqlText(CStr(rs!LocationName)) & " GROUP BY Bucket", EMS_SNAPSHOT)
    fromFile = Not sales.EOF
    Do While Not sales.EOF
        Select Case UCase$(CStr(Nz(sales!Bucket, "")))
            Case "FOOD", "OTHER"
                food = food + CCur(Nz(sales!NetAmount, 0))
            Case "NA"
                na = na + CCur(Nz(sales!NetAmount, 0))
            Case "BEER"
                beer = beer + CCur(Nz(sales!NetAmount, 0))
            Case "LIQUOR"
                liquor = liquor + CCur(Nz(sales!NetAmount, 0))
        End Select
        sales.MoveNext
    Loop
    sales.Close
    If fromFile Then
        Exec "UPDATE tblAssign SET FoodNet=" & food & ", NaNet=" & na & ", AlcoholNet=" & (beer + liquor) & " WHERE AssignID=" & rs!AssignID
    Else
        food = CCur(Nz(rs!FoodNet, 0))
        na = CCur(Nz(rs!NaNet, 0))
        liquor = CCur(Nz(rs!AlcoholNet, 0))
    End If
End Sub

Private Sub WriteInvoice(ByVal eventId As Long, ByVal eventDay As Date, ByVal payeeId As Long, ByVal lines As Collection)
    Dim line As Object
    Dim kind As String
    Dim commission As Currency, gratuity As Currency, adjustments As Currency
    Dim foodNet As Currency, alcoholNet As Currency, foodComm As Currency, alcoholComm As Currency, gross As Currency
    Dim cardFee As Currency, ice As Currency, shortage As Currency, other As Currency
    Dim supplements As Currency
    Dim due As Currency
    Dim invoiceId As Long
    Dim sortIndex As Long
    Dim payee As Object
    kind = "NPO"
    For Each line In lines
        kind = CStr(line("Kind"))
        commission = commission + CCur(line("Commission"))
        gratuity = gratuity + CCur(line("Gratuity"))
        adjustments = adjustments + CCur(line("Adjustments"))
        foodNet = foodNet + CCur(line("FoodNet"))
        alcoholNet = alcoholNet + CCur(line("AlcoholNet"))
        foodComm = foodComm + CCur(line("FoodCommission"))
        alcoholComm = alcoholComm + CCur(line("AlcoholCommission"))
    Next line
    gross = CCur(Nz(Scalar("SELECT Sum(PosSales) FROM tblAssign WHERE EventID=" & eventId & " AND PayeeID=" & payeeId), 0))
    If kind = "SUB" Then VendorCosts eventId, payeeId, eventDay, foodNet + alcoholNet, gross, cardFee, ice, shortage, other
    supplements = CCur(Nz(Scalar("SELECT Sum(Amount) FROM tblSupplement WHERE EventID=" & eventId & " AND PayeeID=" & payeeId), 0))
    due = commission + gratuity + adjustments + supplements - cardFee - ice - shortage - other
    Set payee = CurrentDb.OpenRecordset("SELECT Abbreviation FROM tblPayee WHERE PayeeID=" & payeeId, EMS_SNAPSHOT)
    Exec "INSERT INTO tblInvoice (EventID, PayeeID, InvoiceNumber, Kind, GrossSales, FoodNet, AlcoholNet, GroupNet, FoodCommission, AlcoholCommission, Commission, Gratuity, LocationAdjustments, Supplements, AmountDue, CardFee, IceAmount, Shortage, OtherExpense, Notes) VALUES (" & _
        eventId & "," & payeeId & "," & SqlText(InvoiceNumber(eventDay, Nz(payee!Abbreviation, ""), payeeId)) & "," & SqlText(kind) & "," & _
        gross & "," & foodNet & "," & alcoholNet & "," & (foodNet + alcoholNet) & "," & foodComm & "," & alcoholComm & "," & commission & "," & _
        gratuity & "," & adjustments & "," & supplements & "," & due & "," & cardFee & "," & ice & "," & shortage & "," & other & ",'')"
    payee.Close
    invoiceId = CLng(Scalar("SELECT Max(InvoiceID) FROM tblInvoice WHERE EventID=" & eventId & " AND PayeeID=" & payeeId))
    sortIndex = 1
    For Each line In lines
        Exec "INSERT INTO tblInvoiceLine (InvoiceID, LocationName, SortOrder, NetSales, FoodNet, FoodCommission, AlcoholNet, AlcoholCommission, Commission, Gratuity, Adjustments, Payout) VALUES (" & _
            invoiceId & "," & SqlText(Left$(CStr(line("Location")), 150)) & "," & sortIndex & "," & CCur(line("Net")) & "," & CCur(line("FoodNet")) & "," & _
            CCur(line("FoodCommission")) & "," & CCur(line("AlcoholNet")) & "," & CCur(line("AlcoholCommission")) & "," & CCur(line("Commission")) & "," & _
            CCur(line("Gratuity")) & "," & CCur(line("Adjustments")) & "," & CCur(line("Payout")) & ")"
        sortIndex = sortIndex + 1
    Next line
    If cardFee <> 0 Then AddCalculatedSupplement eventId, payeeId, "CREDIT CARD FEES", -cardFee, "694000"
    If ice <> 0 Then AddCalculatedSupplement eventId, payeeId, "ICE", -ice, "411054"
    If shortage <> 0 Then AddCalculatedSupplement eventId, payeeId, "SHORTAGE", -shortage, ""
    If other <> 0 Then AddCalculatedSupplement eventId, payeeId, "OTHER EXPENSE", -other, "673000"
End Sub

Private Sub VendorCosts(ByVal eventId As Long, ByVal payeeId As Long, ByVal eventDay As Date, ByVal net As Currency, ByVal gross As Currency, ByRef cardFee As Currency, ByRef ice As Currency, ByRef shortage As Currency, ByRef other As Currency)
    Dim rs As Object
    Dim tenders As Currency
    cardFee = 0
    ice = 0
    shortage = 0
    other = 0
    Set rs = CurrentDb.OpenRecordset("SELECT Nz(Sum(CashTender),0) AS CashTender, Nz(Sum(CardTender),0) AS CardTender, Nz(Sum(CorporateVoucher),0) AS CorporateVoucher, Nz(Sum(LevyNfpVoucher),0) AS LevyNfpVoucher, Nz(Sum(UnmGuestVoucher),0) AS UnmGuestVoucher, Nz(Sum(IceBags),0) AS IceBags, Nz(Sum(OtherExpense),0) AS OtherExpense FROM tblVendorInput WHERE EventID=" & eventId & " AND PayeeID=" & payeeId, EMS_SNAPSHOT)
    If rs.EOF Then
        rs.Close
        cardFee = MoneyProduct(net, RateOn(payeeId, "CARD_FEE", eventDay, 0.03))
        Exit Sub
    End If
    tenders = CCur(rs!CashTender) + CCur(rs!CardTender) + CCur(rs!CorporateVoucher) + CCur(rs!LevyNfpVoucher) + CCur(rs!UnmGuestVoucher)
    If tenders <> 0 And tenders < gross Then shortage = gross - tenders
    ice = MoneyProduct(CCur(rs!IceBags), RateOn(payeeId, "ICE_PER_BAG", eventDay, 3))
    other = CCur(rs!OtherExpense)
    cardFee = MoneyProduct(net, RateOn(payeeId, "CARD_FEE", eventDay, 0.03))
    rs.Close
End Sub

Private Sub AddCalculatedSupplement(ByVal eventId As Long, ByVal payeeId As Long, ByVal label As String, ByVal amount As Currency, ByVal glCode As String)
    Exec "INSERT INTO tblSupplement (EventID, PayeeID, Label, Amount, GLCode) VALUES (" & eventId & "," & payeeId & "," & SqlText(label) & "," & amount & "," & SqlText(glCode) & ")"
End Sub

Private Function RoleOf(ByVal text As String) As String
    Select Case UCase$(Trim$(text))
        Case "NPO", "NONPROFIT", "NON-PROFIT", "DONATION"
            RoleOf = "NPO"
        Case "SUB", "VENDOR", "SUBCONTRACTOR"
            RoleOf = "SUB"
        Case "STAFF", "HOURLY", "LEVY"
            RoleOf = "STAFF"
        Case Else
            RoleOf = "OPEN"
    End Select
End Function

Public Sub EMSNewEvent()
    If CurrentProject.AllForms("frmEvent").IsLoaded Then DoCmd.Close acForm, "frmEvent"
    DoCmd.OpenForm "frmEvent", , , , acFormAdd
End Sub

Public Function EMSEventOpen()
    On Error Resume Next
    If Nz(Forms!frmEvent!EventID, 0) = 0 Then
        ApplyEventDefaults
    Else
        EMSEventRefresh
    End If
End Function

Private Sub ApplyEventDefaults()
    Dim office As Object
    With Forms!frmEvent
        If IsNull(!EventDate) Then !EventDate = Date
        If Len(Nz(!EventType, "")) = 0 Then !EventType = "Event"
        If Len(Nz(!Venue, "")) = 0 Then !Venue = "University of New Mexico"
        If Nz(!TaxRate, 0) = 0 Then !TaxRate = 0.07625
        If Nz(!MinimumCommission, 0) = 0 Then !MinimumCommission = 200
        If Len(Nz(!StatusText, "")) = 0 Then !StatusText = "DRAFT"
    End With
    Set office = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM tblOffice", EMS_SNAPSHOT)
    If Not office.EOF Then
        With Forms!frmEvent
            If Len(Nz(!PreparedBy, "")) = 0 Then !PreparedBy = office!PreparedBy
            If Len(Nz(!PreparedTitle, "")) = 0 Then !PreparedTitle = office!PreparedTitle
            If Len(Nz(!ReviewedBy, "")) = 0 Then !ReviewedBy = office!ReviewedBy
            If Len(Nz(!ReviewedTitle, "")) = 0 Then !ReviewedTitle = office!ReviewedTitle
            If Len(Nz(!ApprovedBy, "")) = 0 Then !ApprovedBy = office!ApprovedBy
            If Len(Nz(!ApprovedTitle, "")) = 0 Then !ApprovedTitle = office!ApprovedTitle
        End With
    End If
    office.Close
End Sub

Public Sub EMSEventRefresh()
    On Error Resume Next
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Sub
    Forms!frmEvent!fsubAssign.Requery
    Forms!frmEvent!fsubVendor.Requery
    Forms!frmEvent!fsubSupplement.Requery
    Forms!frmEvent!fsubStaff.Requery
    Forms!frmEvent!lstExceptions.Requery
    Forms!frmEvent!txtDue.Requery
End Sub

Public Sub EMSDeleteEvent()
    Dim eventId As Long
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Sub
    eventId = CLng(Nz(Forms!frmEvent!EventID, 0))
    If eventId = 0 Then
        DoCmd.Close acForm, "frmEvent"
        Exit Sub
    End If
    If MsgBox("Delete this event and its imported sales, assignments, and statements?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Exec "DELETE FROM tblInvoiceLine WHERE InvoiceID In (SELECT InvoiceID FROM tblInvoice WHERE EventID=" & eventId & ")"
    Exec "DELETE FROM tblInvoice WHERE EventID=" & eventId
    Exec "DELETE FROM tblStaffTip WHERE EventID=" & eventId
    Exec "DELETE FROM tblException WHERE EventID=" & eventId
    Exec "DELETE FROM tblSupplement WHERE EventID=" & eventId
    Exec "DELETE FROM tblVendorInput WHERE EventID=" & eventId
    Exec "DELETE FROM tblAssign WHERE EventID=" & eventId
    Exec "DELETE FROM tblSale WHERE EventID=" & eventId
    Exec "DELETE FROM tblTip WHERE EventID=" & eventId
    Exec "DELETE FROM tblImport WHERE EventID=" & eventId
    Exec "DELETE FROM tblEvent WHERE EventID=" & eventId
    DoCmd.Close acForm, "frmEvent"
    If CurrentProject.AllForms("frmHub").IsLoaded Then EMSHubSearch
End Sub

Public Function EMSHubSearch()
    Dim sql As String
    Dim findText As String
    sql = "SELECT e.EventID, e.EventDate AS [Date], e.EventName AS [Event], e.EventType AS [Type], (SELECT Nz(Sum(i.AmountDue),0) FROM tblInvoice AS i WHERE i.EventID=e.EventID) AS [Amount due] FROM tblEvent AS e WHERE 1=1"
    If IsDate(Nz(Forms!frmHub!txtFrom, "")) Then sql = sql & " AND e.EventDate>=#" & Format$(CDate(Forms!frmHub!txtFrom), "mm/dd/yyyy") & "#"
    If IsDate(Nz(Forms!frmHub!txtTo, "")) Then sql = sql & " AND e.EventDate<=#" & Format$(CDate(Forms!frmHub!txtTo), "mm/dd/yyyy") & "#"
    findText = Trim$(Nz(Forms!frmHub!txtFind, ""))
    If Len(findText) > 0 Then
        sql = sql & " AND (e.EventName Like " & SqlText("%" & findText & "%") & " OR e.EventType Like " & SqlText("%" & findText & "%") & _
            " OR e.StatusText Like " & SqlText("%" & findText & "%") & _
            " OR e.EventID In (SELECT a.EventID FROM tblAssign AS a WHERE a.LocationName Like " & SqlText("%" & findText & "%") & ")" & _
            " OR e.EventID In (SELECT i.EventID FROM tblInvoice AS i WHERE i.InvoiceNumber Like " & SqlText("%" & findText & "%") & ")" & _
            " OR e.EventID In (SELECT a.EventID FROM tblAssign AS a INNER JOIN tblPayee AS p ON a.PayeeID=p.PayeeID WHERE p.PayeeName Like " & SqlText("%" & findText & "%") & " OR p.Abbreviation Like " & SqlText("%" & findText & "%") & "))"
    End If
    sql = sql & " ORDER BY e.EventDate DESC, e.EventID DESC"
    Forms!frmHub!lstEvents.RowSource = sql
End Function

Public Function EMSOpenSelected()
    Dim eventId As Long
    eventId = CLng(Nz(Forms!frmHub!lstEvents, 0))
    If eventId = 0 Then Exit Function
    DoCmd.OpenForm "frmEvent", , , "EventID=" & eventId
End Function

Public Sub EMSBackToEvents()
    DoCmd.Close acForm, "frmEvent"
    If Not CurrentProject.AllForms("frmHub").IsLoaded Then DoCmd.OpenForm "frmHub"
    EMSHubSearch
End Sub

Public Sub EMSOpenPayees()
    DoCmd.OpenForm "frmPayee"
End Sub
