Attribute VB_Name = "EMS_Calculate"
Option Compare Database
Option Explicit

Private payeeRows As Variant
Private payeeCount As Long
Private payeeSelecting As Boolean
Private pendingPayeeId As Long
Private emailsRepaired As Boolean

' A blank staff headcount holds only that stand's tips. Donation and vendor
' checks still calculate. A new event can be typed with no MyVenue file;
' stands added by hand use the nets on the row.

Public Function EMSCalculateCurrent()
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Function
    If Forms!frmEvent.Dirty Then Forms!frmEvent.Dirty = False
    EMSCalculateEvent CLng(Nz(Forms!frmEvent!EventID, 0))
    EMSEventRefresh
End Function

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
    EMSEnsureIceColumn
    Set rs = CurrentDb.OpenRecordset("SELECT Nz(Sum(CashTender),0) AS CashTender, Nz(Sum(CardTender),0) AS CardTender, Nz(Sum(CorporateVoucher),0) AS CorporateVoucher, Nz(Sum(LevyNfpVoucher),0) AS LevyNfpVoucher, Nz(Sum(UnmGuestVoucher),0) AS UnmGuestVoucher, Nz(Sum(IceAmount),0) AS IceAmount, Nz(Sum(OtherExpense),0) AS OtherExpense FROM tblVendorInput WHERE EventID=" & eventId & " AND PayeeID=" & payeeId, EMS_SNAPSHOT)
    If rs.EOF Then
        rs.Close
        cardFee = MoneyProduct(net, RateOn(payeeId, "CARD_FEE", eventDay, 0.03))
        Exit Sub
    End If
    tenders = CCur(rs!CashTender) + CCur(rs!CardTender) + CCur(rs!CorporateVoucher) + CCur(rs!LevyNfpVoucher) + CCur(rs!UnmGuestVoucher)
    If tenders <> 0 And tenders < gross Then shortage = gross - tenders
    ice = CCur(rs!IceAmount)
    other = CCur(rs!OtherExpense)
    cardFee = MoneyProduct(net, RateOn(payeeId, "CARD_FEE", eventDay, 0.03))
    rs.Close
End Sub

Public Function EMSEnsureIceColumn()
    Dim rs As Object
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset("SELECT IceAmount FROM tblVendorInput WHERE 1=0", EMS_SNAPSHOT)
    If Err.Number <> 0 Then
        Err.Clear
        CurrentDb.Execute "ALTER TABLE tblVendorInput ADD COLUMN IceAmount CURRENCY"
    Else
        rs.Close
    End If
    On Error GoTo 0
End Function

Public Function EMSRateTyped()
    Dim amount As Double
    If Not IsNumeric(Screen.ActiveControl.Value) Then Exit Function
    amount = CDbl(Screen.ActiveControl.Value)
    If amount > 1 Then Screen.ActiveControl.Value = amount / 100
End Function

Public Function EMSBulkPayees()
    If CurrentProject.AllForms("frmPayee").IsLoaded Then
        If Forms!frmPayee.Dirty Then Forms!frmPayee.Dirty = False
    End If
    DoCmd.OpenTable "tblPayee", acViewNormal, acEdit
End Function

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

Public Function EMSRemoveStand()
    Dim child As Object
    On Error GoTo Failed
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Function
    Set child = Forms!frmEvent!fsubAssign.Form
    If child.NewRecord Then Exit Function
    If MsgBox("Remove " & Nz(child!LocationName, "this stand") & " from the event?", vbYesNo + vbQuestion) <> vbYes Then Exit Function
    child.AllowDeletions = True
    If child.Dirty Then child.Dirty = False
    child.Recordset.Delete
    child.Requery
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function EMSNewEvent()
    If CurrentProject.AllForms("frmEvent").IsLoaded Then DoCmd.Close acForm, "frmEvent"
    DoCmd.OpenForm "frmEvent", , , , acFormAdd
End Function

Public Function EMSEventOpen()
    On Error Resume Next
    If Nz(Forms!frmEvent!EventID, 0) = 0 Then
        ApplyEventDefaults
    Else
        EMSEventRefresh
    End If
End Function

Private Sub ApplyEventDefaults()
    With Forms!frmEvent
        If IsNull(!EventDate) Then !EventDate = Date
        If Len(Nz(!EventType, "")) = 0 Then !EventType = "Event"
        If Len(Nz(!Venue, "")) = 0 Then !Venue = "University of New Mexico"
        If Nz(!TaxRate, 0) = 0 Then !TaxRate = 0.07625
        If Nz(!MinimumCommission, 0) = 0 Then !MinimumCommission = 200
    End With
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

Public Function EMSDeleteEvent()
    Dim eventId As Long
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then Exit Function
    eventId = CLng(Nz(Forms!frmEvent!EventID, 0))
    If eventId = 0 Then
        DoCmd.Close acForm, "frmEvent"
        Exit Function
    End If
    If MsgBox("Delete this event and its imported sales, assignments, and statements?", vbYesNo + vbQuestion) <> vbYes Then Exit Function
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
End Function

Public Function EMSHubSearch()
    EMSEnsurePartners
    Dim sql As String
    Dim findText As String
    sql = "SELECT e.EventID, e.EventDate AS [Date], e.EventName AS [Event], e.EventType AS [Type], (SELECT Nz(Sum(i.AmountDue),0) FROM tblInvoice AS i WHERE i.EventID=e.EventID) AS [Amount due] FROM tblEvent AS e WHERE 1=1"
    If IsDate(Nz(Forms!frmHub!txtFrom, "")) Then sql = sql & " AND e.EventDate>=#" & Format$(CDate(Forms!frmHub!txtFrom), "mm/dd/yyyy") & "#"
    If IsDate(Nz(Forms!frmHub!txtTo, "")) Then sql = sql & " AND e.EventDate<=#" & Format$(CDate(Forms!frmHub!txtTo), "mm/dd/yyyy") & "#"
    findText = Trim$(Nz(Forms!frmHub!txtFind, ""))
    If Len(findText) > 0 Then
        sql = sql & " AND (e.EventName Like " & SqlText("%" & findText & "%") & " OR e.EventType Like " & SqlText("%" & findText & "%") & _
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

Public Function EMSBackToEvents()
    On Error Resume Next
    DoCmd.Close acForm, "frmPayee"
    DoCmd.Close acForm, "frmEvent"
    On Error GoTo 0
    If Not CurrentProject.AllForms("frmHub").IsLoaded Then DoCmd.OpenForm "frmHub"
    EMSHubSearch
End Function

Public Function EMSPayeeCurrent()
    Dim frm As Object
    EMSRepairPayeeEmails
    Set frm = Forms!frmPayee
    If IsNull(frm!PayeeID) Then
        If Not payeeSelecting Then frm!lstPayees = Null
        EMSPayeeLock False
        frm!lblStatus.Caption = "New payee. Enter the name, then Save."
    Else
        ' Assigning the list here, while the click is still finishing, clears the highlight.
        If Not payeeSelecting Then EMSPayeeShowRow CLng(frm!PayeeID), False
        EMSPayeeLock True
        frm!lblStatus.Caption = "Viewing " & Nz(frm!txtPayeeName, "this payee") & ". Choose Edit to change it."
    End If
End Function

Public Function EMSPayeePick()
    Dim frm As Object
    Dim id As Long
    Set frm = Forms!frmPayee
    id = CLng(Nz(frm!lstPayees, 0))
    If id = 0 Then Exit Function
    If frm.NewRecord And Len(Trim$(Nz(frm!txtPayeeName, ""))) = 0 Then
        frm.Undo
    ElseIf Len(Trim$(Nz(frm!txtPayeeName, ""))) = 0 Then
        MsgBox "Enter the payee name before leaving this record.", vbExclamation
        Exit Function
    ElseIf frm.Dirty Then
        frm.Dirty = False
    End If
    pendingPayeeId = id
    payeeSelecting = True
    DoCmd.SearchForRecord acDataForm, "frmPayee", acFirst, "PayeeID=" & id
    payeeSelecting = False
    ' The record change moves the cursor off the list, so the row highlight vanishes.
    ' Put it back after Access finishes the click.
    frm.OnTimer = "=EMSPayeeListFocus()"
    frm.TimerInterval = 100
End Function

Public Function EMSPayeeListFocus()
    Dim frm As Object
    Set frm = Forms!frmPayee
    frm.TimerInterval = 0
    If pendingPayeeId = 0 Then Exit Function
    EMSPayeeShowRow pendingPayeeId, True
End Function

Private Sub EMSPayeeShowRow(ByVal id As Long, ByVal takeFocus As Boolean)
    Dim frm As Object
    Dim row As Long
    Set frm = Forms!frmPayee
    On Error Resume Next
    If CLng(Nz(frm!lstPayees, 0)) <> id Then frm!lstPayees = id
    For row = 0 To frm!lstPayees.ListCount - 1
        If CLng(Nz(frm!lstPayees.Column(0, row), 0)) = id Then
            frm!lstPayees.Selected(row) = True
            Exit For
        End If
    Next row
    If takeFocus Then frm!lstPayees.SetFocus
    On Error GoTo 0
End Sub

Public Function EMSPayeeEdit()
    Dim frm As Object
    Set frm = Forms!frmPayee
    If IsNull(frm!PayeeID) Then Exit Function
    frm.TimerInterval = 0
    EMSPayeeLock False
    frm!txtPayeeName.SetFocus
    frm!lblStatus.Caption = "Editing " & frm!txtPayeeName & ". Save keeps the changes."
End Function

Public Function EMSPayeeSave()
    Dim frm As Object
    Set frm = Forms!frmPayee
    If Len(Trim$(Nz(frm!txtPayeeName, ""))) = 0 Then
        MsgBox "Enter the payee name.", vbExclamation
        Exit Function
    End If
    If frm.Dirty Then frm.Dirty = False
    If Not Nz(frm!chkActive, False) Then frm!chkShowInactive = True
    EMSPayeeFilter
    EMSPayeeLock True
    frm!lblStatus.Caption = "Saved " & frm!txtPayeeName & "."
End Function

Public Function EMSPayeeCancel()
    Dim frm As Object
    Set frm = Forms!frmPayee
    On Error Resume Next
    frm.Undo
    On Error GoTo 0
    If frm.NewRecord Then
        EMSPayeeLock False
        frm!lblStatus.Caption = "New payee. Enter the name, then Save."
    Else
        EMSPayeeLock True
        frm!lblStatus.Caption = "Changes discarded."
    End If
End Function

Public Function EMSPayeeFilter()
    Dim frm As Object
    Set frm = Forms!frmPayee
    frm!lstPayees.Requery
    If Not IsNull(frm!PayeeID) Then EMSPayeeShowRow CLng(frm!PayeeID), False
End Function

' Adds any payee columns this screen expects. A missing name is what made Access ask for SAPID.
Public Sub EMSEnsurePayeeColumns()
    EnsureColumn "tblPayee", "SAPID", "TEXT(255)", False
    EnsureColumn "tblPayee", "BSSID", "TEXT(255)", False
    EnsureColumn "tblPayee", "Address1", "TEXT(255)", False
    EnsureColumn "tblPayee", "Address2", "TEXT(255)", False
    EnsureColumn "tblPayee", "Phone", "TEXT(255)", False
    EnsureColumn "tblPayee", "Notes", "MEMO", False
    EnsureColumn "tblPayee", "Abbreviation", "TEXT(255)", False
    EnsureColumn "tblPayee", "PayeeType", "TEXT(255)", False
    ' The open file has no Active column. That is the "unknown field" and the list error.
    EnsureColumn "tblPayee", "Active", "YESNO", True
End Sub

' The email grid was bound to EmailAddress, which this file did not have, so the cell would not show or accept an address.
Public Sub EMSEnsureEmailColumns()
    Dim db As Object
    Dim td As Object
    Dim fld As Object
    Dim prior As String
    EnsureColumn "tblPayeeEmail", "PayeeID", "LONG", False
    prior = ""
    On Error Resume Next
    Set db = CurrentDb
    Set td = db.TableDefs("tblPayeeEmail")
    For Each fld In td.Fields
        If InStr(1, UCase$(CStr(fld.Name)), "EMAIL", vbTextCompare) > 0 Then
            If StrComp(CStr(fld.Name), "EmailID", vbTextCompare) <> 0 And StrComp(CStr(fld.Name), "EmailAddress", vbTextCompare) <> 0 Then
                prior = CStr(fld.Name)
                Exit For
            End If
        End If
    Next fld
    On Error GoTo 0
    EnsureColumn "tblPayeeEmail", "EmailAddress", "TEXT(255)", False
    EnsureColumn "tblPayeeEmail", "Label", "TEXT(255)", False
    EnsureColumn "tblPayeeEmail", "IsPrimary", "YESNO", False
    If Not ColumnExists("tblPayeeEmail", "EmailID") Then
        On Error Resume Next
        CurrentDb.Execute "ALTER TABLE [tblPayeeEmail] ADD COLUMN [EmailID] COUNTER"
        If Err.Number <> 0 Then
            Err.Clear
            CurrentDb.Execute "ALTER TABLE [tblPayeeEmail] ADD COLUMN [EmailID] AUTOINCREMENT"
        End If
        CurrentDb.TableDefs.Refresh
        On Error GoTo 0
    End If
    If Len(prior) > 0 Then
        On Error Resume Next
        CurrentDb.Execute "UPDATE tblPayeeEmail SET [EmailAddress]=[" & prior & "] WHERE [EmailAddress] Is Null OR [EmailAddress]=''"
        On Error GoTo 0
    End If
End Sub

Public Sub EMSRepairPayeeEmails()
    Dim host As Object
    Dim child As Object
    Dim ctl As Object
    If emailsRepaired Then Exit Sub
    On Error Resume Next
    EMSEnsureEmailColumns
    EMSEnsurePartners
    If Not CurrentProject.AllForms("frmPayee").IsLoaded Then
        emailsRepaired = True
        Exit Sub
    End If
    Set host = Forms!frmPayee
    Set child = host!fsubEmail.Form
    child.RecordSource = "tblPayeeEmail"
    For Each ctl In child.Controls
        If StrComp(CStr(ctl.Name), "colEmailAddress", vbTextCompare) = 0 Or StrComp(CStr(ctl.ControlSource), "EmailAddress", vbTextCompare) = 0 Then ctl.ControlSource = "EmailAddress"
        If StrComp(CStr(ctl.Name), "colLabel", vbTextCompare) = 0 Or StrComp(CStr(ctl.ControlSource), "Label", vbTextCompare) = 0 Then ctl.ControlSource = "Label"
        If StrComp(CStr(ctl.Name), "colIsPrimary", vbTextCompare) = 0 Or StrComp(CStr(ctl.ControlSource), "IsPrimary", vbTextCompare) = 0 Then ctl.ControlSource = "IsPrimary"
    Next ctl
    child.Requery
    emailsRepaired = True
    On Error GoTo 0
End Sub

Public Sub EMSDropPayeeQuery()
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryPayeeList"
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ColumnExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim td As Object
    Dim fld As Object
    On Error GoTo Missing
    Set td = CurrentDb.TableDefs(tableName)
    For Each fld In td.Fields
        If StrComp(CStr(fld.Name), fieldName, vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next fld
    Exit Function
Missing:
    ColumnExists = False
End Function

Private Sub EnsureColumn(ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As String, ByVal fillYes As Boolean)
    Dim db As Object
    If ColumnExists(tableName, fieldName) Then Exit Sub
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] " & fieldType
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    db.TableDefs.Refresh
    If fillYes Then db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=-1"
    On Error GoTo 0
End Sub

' Fills the payee list without a query on the form. That query was the SAPID prompt.
Public Function EMSPayeeList(fld As Control, id As Variant, row As Variant, col As Variant, code As Variant) As Variant
    Select Case code
        Case acLBInitialize
            EMSPayeeList = True
        Case acLBOpen
            LoadPayeeRows
            EMSPayeeList = 1
        Case acLBGetRowCount
            EMSPayeeList = payeeCount
        Case acLBGetColumnCount
            EMSPayeeList = 7
        Case acLBGetColumnWidth
            EMSPayeeList = -1
        Case acLBGetValue
            If row >= 0 And row < payeeCount Then EMSPayeeList = payeeRows(row, col)
    End Select
End Function

Private Sub LoadPayeeRows()
    Dim rs As Object
    Dim sql As String
    Dim showAll As Boolean
    Dim index As Long
    EMSEnsurePayeeColumns
    showAll = False
    If CurrentProject.AllForms("frmPayee").IsLoaded Then showAll = CBool(Nz(Forms!frmPayee!chkShowInactive, False))
    sql = "SELECT PayeeID, PayeeName, PayeeType, Abbreviation, SAPID, BSSID, [Active] FROM tblPayee"
    If Not showAll Then sql = sql & " WHERE [Active]<>0"
    sql = sql & " ORDER BY PayeeName"
    Set rs = CurrentDb.OpenRecordset(sql, EMS_SNAPSHOT)
    payeeCount = 0
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    rs.MoveLast
    payeeCount = rs.RecordCount
    ReDim payeeRows(0 To payeeCount - 1, 0 To 6)
    rs.MoveFirst
    index = 0
    Do While Not rs.EOF
        payeeRows(index, 0) = rs!PayeeID
        payeeRows(index, 1) = Nz(rs!PayeeName, "")
        payeeRows(index, 2) = Nz(rs!PayeeType, "")
        payeeRows(index, 3) = Nz(rs!Abbreviation, "")
        payeeRows(index, 4) = Nz(rs!SAPID, "")
        payeeRows(index, 5) = Nz(rs!BSSID, "")
        payeeRows(index, 6) = IIf(Nz(rs!Active, False), "Active", "Inactive")
        index = index + 1
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub EMSPayeeLock(ByVal locked As Boolean)
    Dim frm As Object
    Dim name As Variant
    Set frm = Forms!frmPayee
    For Each name In Array("txtPayeeName", "cboType", "txtAbbreviation", "chkActive", "txtSAPID", "txtBSSID", "txtAddress1", "txtAddress2", "txtNotes")
        frm.Controls(CStr(name)).Locked = locked
    Next name
    frm!cmdEdit.Enabled = locked And Not frm.NewRecord
    frm!cmdSave.Enabled = Not locked
    frm!cmdCancel.Enabled = Not locked
    On Error Resume Next
    frm!fsubEmail.Form.AllowEdits = Not locked
    frm!fsubEmail.Form.AllowAdditions = Not locked
    frm!fsubEmail.Form.AllowDeletions = Not locked
    frm!fsubRate.Form.AllowEdits = Not locked
    frm!fsubRate.Form.AllowAdditions = Not locked
    frm!fsubRate.Form.AllowDeletions = Not locked
    On Error GoTo 0
End Sub

Public Function EMSOpenPayees()
    Dim added As Long
    added = EMSEnsurePartners()
    DoCmd.OpenForm "frmPayee"
    If added > 0 Then MsgBox added & " payee(s) from the prior database were added.", vbInformation
End Function

' Inserts the nonprofit and vendor records from the prior payout file.
' An abbreviation that is already here is left in place. Returns how many were added.
Public Function EMSEnsurePartners() As Long
    Dim added As Long
    Dim payeeId As Long
    EMSEnsurePayeeColumns
    payeeId = EnsurePayee("Albuquerque Rebels Volleyball Club", "NPO", "ARVC", "7800 LAS LOMITAS AVENUE NE", "Albuquerque, NM 87113-2645", "1013051", "5487", "Senobia Szuch", True, added)
    EnsureEmail payeeId, "gabe.padilla07@yahoo.com", "Primary", True
    EnsureEmail payeeId, "jillianshane@sbcglobal.net", "Additional", False
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("Duke City Volleyball Athletics Inc", "NPO", "DCVA", "8717 DEMOCRACY NE", "Albuquerque, NM 87109-5183", "1026871", "10468", "Shannon Mick", True, added)
    EnsureEmail payeeId, "fundraising@dcvanm.com", "Primary", True
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("APS - Rio Grande High School", "NPO", "RGHS", "2300 ARENAL RD SW", "Albuquerque, NM 87105-4180", "1029380", "11037", "Chris Smalls", True, added)
    EnsureEmail payeeId, "christopher.smalls@aps.edu", "Primary", True
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("58 SOW TOP 3 ASSOCIATION", "NPO", "USAF", "5249 HERCULES WY SE", "Albuquerque, NM 87117-0001", "1031810", "11773", "Daniel Gonzalez", True, added)
    EnsureEmail payeeId, "daniel.gonzalez.23@us.af.mil", "Primary", True
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("APS - Del Norte High School Cheer", "NPO", "DNHS", "", "", "1027902", "10743", "", True, added)
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("DJCF INC. (DBA: E3VB)", "NPO", "E3VB", "6919 LUZ DEL SOL PL NW", "Albuquerque, NM 87114-3702", "1031901", "11793", "Sarah Pulling", True, added)
    EnsureEmail payeeId, "smith5411@gmail.com", "Primary", True
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("VICTORY OUTREACH INC., ROSWELL", "NPO", "VOUTREACH", "", "", "1029375", "11040", "", True, added)
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("UNITED FIT", "NPO", "UF", "", "", "982165", "5190", "", True, added)
    EnsureNpoRates payeeId
    payeeId = EnsurePayee("Fuego Enterprises LLC", "SUB", "LAG", "7770 JEFFERSON ST NE STE 302", "Albuquerque, NM 87109-4386", "992855", "5177", "FAlvarez", True, added)
    EnsureEmail payeeId, "FAlvarez@fuegonm.com", "Primary", True
    RepairLagunaRates payeeId
    payeeId = EnsurePayee("NXTLVL HOLDINGS LLC SUGAR SHIVERS", "SUB", "SUGAR", "533 61ST ST NW", "Albuquerque, NM 87105-1407", "1037451", "13470", "Victor Martiez", True, added)
    EnsureEmail payeeId, "nxtlevelholdingsllc@gmail.com", "Primary", True
    EnsureSubRates payeeId, 0.8, 0.1, 0, 0, #1/1/2026#, Null
    payeeId = EnsurePayee("BNW Enterprises LLC", "SUB", "NOOD", "", "", "", "", "", True, added)
    EnsureSubRates payeeId, 0.9, 0.1, 0, 0, #1/1/2026#, Null
    payeeId = EnsurePayee("OHANA FOODS LLC", "SUB", "OHANA", "130 DEER PARK DR", "Alto, NM 88312-8029", "1030237", "", "Mike Walker", True, added)
    EnsureEmail payeeId, "wok1004@hotmail.com", "Primary", True
    EnsureSubRates payeeId, 0.9, 0.1, 0, 0, #1/1/2026#, Null
    payeeId = EnsurePayee("NXTLVL HOLDINGS LLC", "SUB", "SUG", "", "", "1037451", "10809", "", True, added)
    EnsureSubRates payeeId, 0.8, 0.1, 0, 0, #9/1/2026#, Null
    EMSEnsurePartners = added
End Function

Private Function EnsurePayee(ByVal payeeName As String, ByVal payeeType As String, ByVal abbreviation As String, ByVal address1 As String, ByVal address2 As String, ByVal sapId As String, ByVal bssId As String, ByVal contact As String, ByVal isActive As Boolean, ByRef added As Long) As Long
    Dim id As Variant
    Dim noteText As String
    If abbreviation = "VOUTREACH" Then
        id = PayeeIdFor("VICT")
        If Not IsNull(id) Then Exec "UPDATE tblPayee SET Abbreviation='VOUTREACH' WHERE PayeeID=" & id
    End If
    id = PayeeIdFor(abbreviation)
    noteText = ""
    If Len(contact) > 0 Then noteText = "Primary contact: " & contact
    If abbreviation = "DNHS" Or abbreviation = "VOUTREACH" Then
        If Len(noteText) > 0 Then noteText = noteText & ". "
        noteText = noteText & "Inactive in the prior database."
    End If
    If abbreviation = "SUG" Then noteText = "Earlier Sugar Shivers contract. Current payee abbreviation is SUGAR."
    If abbreviation = "OHANA" Then noteText = "Primary contact: Mike Walker. Prior database also used abbreviation NOOD."
    If IsNull(id) Then
        Exec "INSERT INTO tblPayee (PayeeName, PayeeType, Abbreviation, Address1, Address2, SAPID, BSSID, Phone, Notes, [Active]) VALUES (" & _
            SqlText(payeeName) & "," & SqlText(payeeType) & "," & SqlText(abbreviation) & "," & SqlText(address1) & "," & SqlText(address2) & "," & _
            SqlText(sapId) & "," & SqlText(bssId) & ",''," & SqlText(noteText) & "," & IIf(isActive, "-1", "0") & ")"
        id = PayeeIdFor(abbreviation)
        added = added + 1
    Else
        FillBlank "tblPayee", "Address1", address1, CLng(id)
        FillBlank "tblPayee", "Address2", address2, CLng(id)
        FillBlank "tblPayee", "SAPID", sapId, CLng(id)
        FillBlank "tblPayee", "BSSID", bssId, CLng(id)
        FillBlank "tblPayee", "Notes", noteText, CLng(id)
    End If
    EnsurePayee = CLng(id)
End Function

Private Sub FillBlank(ByVal tableName As String, ByVal fieldName As String, ByVal value As String, ByVal payeeId As Long)
    If Len(value) = 0 Then Exit Sub
    Exec "UPDATE " & tableName & " SET " & fieldName & "=" & SqlText(value) & " WHERE PayeeID=" & payeeId & " AND (" & fieldName & " Is Null OR " & fieldName & "='')"
End Sub

Private Function PayeeIdFor(ByVal abbreviation As String) As Variant
    Dim rs As Object
    PayeeIdFor = Null
    On Error GoTo Done
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tblPayee]", EMS_SNAPSHOT)
    Do Until rs.EOF
        If StrComp(CStr(Nz(rs!Abbreviation, "")), abbreviation, vbTextCompare) = 0 Then
            PayeeIdFor = rs!PayeeID
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
Done:
End Function

Private Sub EnsureEmail(ByVal payeeId As Long, ByVal address As String, ByVal label As String, ByVal isPrimary As Boolean)
    Dim rs As Object
    Dim cols As String
    Dim vals As String
    If Len(address) = 0 Or payeeId = 0 Then Exit Sub
    EMSEnsureEmailColumns
    On Error GoTo Done
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tblPayeeEmail]", EMS_SNAPSHOT)
    If Not RecordHas(rs, "PayeeID") Or Not RecordHas(rs, "EmailAddress") Then
        rs.Close
        Exit Sub
    End If
    Do Until rs.EOF
        If CLng(Nz(rs!PayeeID, 0)) = payeeId Then
            If StrComp(CStr(Nz(rs!EmailAddress, "")), address, vbTextCompare) = 0 Then
                rs.Close
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    cols = "[PayeeID], [EmailAddress]"
    vals = payeeId & "," & SqlText(address)
    If ColumnExists("tblPayeeEmail", "Label") Then
        cols = cols & ", [Label]"
        vals = vals & "," & SqlText(label)
    End If
    If ColumnExists("tblPayeeEmail", "IsPrimary") Then
        cols = cols & ", [IsPrimary]"
        vals = vals & "," & IIf(isPrimary, "-1", "0")
    End If
    Exec "INSERT INTO [tblPayeeEmail] (" & cols & ") VALUES (" & vals & ")"
Done:
End Sub

Private Function RecordHas(ByVal rs As Object, ByVal fieldName As String) As Boolean
    Dim fld As Object
    For Each fld In rs.Fields
        If StrComp(CStr(fld.Name), fieldName, vbTextCompare) = 0 Then
            RecordHas = True
            Exit Function
        End If
    Next fld
End Function

Private Sub EnsureNpoRates(ByVal payeeId As Long)
    EnsureRate payeeId, "DONATION_FOOD", 0.1, #1/1/2026#, Null
    EnsureRate payeeId, "DONATION_ALCOHOL", 0.08, #1/1/2026#, Null
    EnsureRate payeeId, "TIP_WITHHOLD", 0.1, #1/1/2026#, Null
End Sub

Private Sub EnsureSubRates(ByVal payeeId As Long, ByVal foodRate As Double, ByVal naRate As Double, ByVal beerRate As Double, ByVal liquorRate As Double, ByVal startOn As Date, ByVal endOn As Variant)
    EnsureRate payeeId, "FOOD", foodRate, startOn, endOn
    EnsureRate payeeId, "NA", naRate, startOn, endOn
    EnsureRate payeeId, "BEER", beerRate, startOn, endOn
    EnsureRate payeeId, "LIQUOR", liquorRate, startOn, endOn
    EnsureRate payeeId, "TIP_WITHHOLD", 0.1, startOn, endOn
    EnsureRate payeeId, "CARD_FEE", 0.03, startOn, endOn
End Sub

Private Sub RepairLagunaRates(ByVal payeeId As Long)
    Dim foodId As Variant
    Dim tipId As Variant
    foodId = RateIdFor(payeeId, "FOOD", #1/1/2026#, 0.7)
    If Not IsNull(foodId) Then
        Exec "UPDATE tblPayeeRate SET RateAmount=0.8, EndDate=#10/17/2026# WHERE RateID=" & foodId
    End If
    tipId = RateIdFor(payeeId, "TIP_WITHHOLD", #1/1/2026#, 0.1)
    If Not IsNull(tipId) Then
        Exec "UPDATE tblPayeeRate SET RateAmount=0, EndDate=#10/17/2026# WHERE RateID=" & tipId
    End If
    EnsureRate payeeId, "FOOD", 0.8, #1/1/2026#, #10/17/2026#
    EnsureRate payeeId, "FOOD", 0.7, #10/18/2026#, Null
    EnsureRate payeeId, "NA", 0.05, #1/1/2026#, Null
    EnsureRate payeeId, "BEER", 0, #1/1/2026#, Null
    EnsureRate payeeId, "LIQUOR", 0, #1/1/2026#, Null
    EnsureRate payeeId, "TIP_WITHHOLD", 0, #1/1/2026#, #10/17/2026#
    EnsureRate payeeId, "TIP_WITHHOLD", 0.1, #10/18/2026#, Null
    EnsureRate payeeId, "CARD_FEE", 0.03, #1/1/2026#, Null
End Sub

Private Sub EnsureRate(ByVal payeeId As Long, ByVal code As String, ByVal amount As Double, ByVal startOn As Date, ByVal endOn As Variant)
    Dim endSql As String
    If RateExists(payeeId, code, startOn) Then Exit Sub
    If IsNull(endOn) Then endSql = "Null" Else endSql = SqlHash(CDate(endOn))
    Exec "INSERT INTO tblPayeeRate (PayeeID, RateCode, RateAmount, EffectiveDate, EndDate, Notes) VALUES (" & _
        payeeId & "," & SqlText(code) & "," & SqlRate(amount) & "," & SqlHash(startOn) & "," & endSql & ",'')"
End Sub

Private Function RateExists(ByVal payeeId As Long, ByVal code As String, ByVal startOn As Date) As Boolean
    Dim rs As Object
    On Error GoTo Done
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tblPayeeRate]", EMS_SNAPSHOT)
    Do Until rs.EOF
        If CLng(Nz(rs!PayeeID, 0)) = payeeId Then
            If StrComp(CStr(Nz(rs!RateCode, "")), code, vbTextCompare) = 0 Then
                If IsDate(rs!EffectiveDate) Then
                    If CDate(rs!EffectiveDate) = startOn Then
                        RateExists = True
                        Exit Do
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
Done:
End Function

Private Function RateIdFor(ByVal payeeId As Long, ByVal code As String, ByVal startOn As Date, ByVal amountMatch As Double) As Variant
    Dim rs As Object
    RateIdFor = Null
    On Error GoTo Done
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tblPayeeRate]", EMS_SNAPSHOT)
    If Not RecordHas(rs, "RateID") Then
        rs.Close
        Exit Function
    End If
    Do Until rs.EOF
        If CLng(Nz(rs!PayeeID, 0)) = payeeId Then
            If StrComp(CStr(Nz(rs!RateCode, "")), code, vbTextCompare) = 0 Then
                If IsDate(rs!EffectiveDate) Then
                    If CDate(rs!EffectiveDate) = startOn Then
                        If Abs(CDbl(Nz(rs!RateAmount, 0)) - amountMatch) < 0.000001 Then
                            If IsNull(rs!EndDate) Then
                                RateIdFor = rs!RateID
                                Exit Do
                            End If
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
Done:
End Function

Private Function SqlHash(ByVal day As Date) As String
    SqlHash = "#" & Format$(day, "mm/dd/yyyy") & "#"
End Function

Private Function SqlRate(ByVal amount As Double) As String
    SqlRate = Replace$(Format$(amount, "0.000000"), ",", ".")
End Function
