Attribute VB_Name = "EMS_Import"
Option Compare Database
Option Explicit

' Reads the two raw MyVenue exports. A new pair creates the event.
' Importing onto an open event replaces the sales and tips and keeps assignments.

Public Function EMSImportNew()
    EMSImportPair 0
End Function

Public Function EMSImportCurrent()
    If Not CurrentProject.AllForms("frmEvent").IsLoaded Then
        MsgBox "Open an event first.", vbInformation
        Exit Function
    End If
    If Forms!frmEvent.Dirty Then Forms!frmEvent.Dirty = False
    EMSImportPair CLng(Nz(Forms!frmEvent!EventID, 0))
    EMSEventRefresh
End Function

Public Sub EMSImportPair(ByVal eventId As Long)
    On Error GoTo Failed
    Dim pricePath As String
    Dim tipsPath As String
    Dim priceRows As Collection
    Dim tipRows As Collection
    Dim priceName As String
    Dim tipsName As String
    Dim priceDay As Date
    Dim tipsDay As Date
    Dim priceFound As Boolean
    Dim tipsFound As Boolean

    pricePath = PickFile("Choose the Price Level category report")
    If Len(pricePath) = 0 Then Exit Sub
    tipsPath = PickFile("Choose the Attendant Tips report")
    If Len(tipsPath) = 0 Then Exit Sub

    Set priceRows = ParseCsv(ReadTextFile(pricePath))
    Set tipRows = ParseCsv(ReadTextFile(tipsPath))
    ReadHeader priceRows, priceName, priceDay, priceFound
    ReadHeader tipRows, tipsName, tipsDay, tipsFound
    If Not priceFound Or Not tipsFound Then
        MsgBox "One of the files is missing the Event and Start Date header.", vbExclamation
        Exit Sub
    End If
    If EventKey(priceName) <> EventKey(tipsName) Or priceDay <> tipsDay Then
        MsgBox "These two reports are not the same event." & vbCrLf & _
            priceName & "  " & Format$(priceDay, "mm/dd/yyyy") & vbCrLf & _
            tipsName & "  " & Format$(tipsDay, "mm/dd/yyyy"), vbExclamation
        Exit Sub
    End If

    Dim created As Boolean
    Dim importId As Long
    eventId = ResolveEvent(eventId, priceName, priceDay, created)
    importId = ReplaceReports(eventId, pricePath, tipsPath, priceRows, tipRows)

    Dim count As Variant
    count = Scalar("SELECT Count(*) FROM tblException WHERE EventID=" & eventId)
    MsgBox IIf(created, "Event created. ", "Reports replaced. Assignments were kept. ") & _
        "Review " & Nz(count, 0) & " exception(s), assign any open stands, then calculate.", vbInformation
    If CurrentProject.AllForms("frmHub").IsLoaded Then EMSHubSearch
    If created Then DoCmd.OpenForm "frmEvent", , , "EventID=" & eventId
    Exit Sub
Failed:
    MsgBox "Import stopped. Run the import again if this event was only partly loaded." & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub EnsureImportColumns()
    EnsureImportColumn "tblImport", "PriceFile", "TEXT(255)"
    EnsureImportColumn "tblImport", "TipsFile", "TEXT(255)"
    EnsureImportColumn "tblImport", "ImportedAt", "DATETIME"
    EnsureImportColumn "tblImport", "IsCurrent", "YESNO"
    EnsureImportColumn "tblLocation", "Active", "YESNO"
End Sub

Private Sub EnsureImportColumn(ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As String)
    Dim td As Object
    Dim fld As Object
    Dim found As Boolean
    On Error GoTo Missing
    Set td = CurrentDb.TableDefs(tableName)
    For Each fld In td.Fields
        If StrComp(CStr(fld.Name), fieldName, vbTextCompare) = 0 Then found = True
    Next fld
    If found Then Exit Sub
    CurrentDb.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] " & fieldType
    CurrentDb.TableDefs.Refresh
Missing:
End Sub

Private Function ResolveEvent(ByVal eventId As Long, ByVal eventName As String, ByVal eventDay As Date, ByRef created As Boolean) As Long
    Dim rs As Object
    created = False
    If eventId > 0 Then
        ResolveEvent = eventId
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT EventID, EventName FROM tblEvent WHERE EventDate=#" & Format$(eventDay, "mm/dd/yyyy") & "#", EMS_SNAPSHOT)
    Do While Not rs.EOF
        If EventKey(CStr(rs!EventName)) = EventKey(eventName) Then
            ResolveEvent = CLng(rs!EventID)
            rs.Close
            Exit Function
        End If
        rs.MoveNext
    Loop
    rs.Close
    created = True
    ResolveEvent = InsertEvent(eventName, eventDay)
End Function

Private Function InsertEvent(ByVal eventName As String, ByVal eventDay As Date) As Long
    Exec "INSERT INTO tblEvent (EventDate, EventName, EventType, Venue, TaxRate, MinimumCommission, Notes, CreatedAt) VALUES (" & _
        "#" & Format$(eventDay, "mm/dd/yyyy") & "#," & SqlText(Left$(eventName, 200)) & "," & SqlText(GuessType(eventName)) & "," & _
        SqlText("University of New Mexico") & ",0.07625,200,''," & _
        "#" & Format$(Now, "mm/dd/yyyy hh:nn:ss") & "#)"
    InsertEvent = CLng(Scalar("SELECT Max(EventID) FROM tblEvent"))
End Function

Private Function GuessType(ByVal eventName As String) As String
    Dim text As String
    text = UCase$(eventName)
    If InStr(text, "FOOTBALL") > 0 Then
        GuessType = "Football"
    ElseIf InStr(text, "BASKETBALL") > 0 And InStr(text, "WOMEN") > 0 Then
        GuessType = "Women's Basketball"
    ElseIf InStr(text, "BASKETBALL") > 0 Then
        GuessType = "Men's Basketball"
    ElseIf InStr(text, "SOCCER") > 0 Then
        GuessType = "Soccer"
    ElseIf InStr(text, "BASEBALL") > 0 Then
        GuessType = "Baseball"
    ElseIf InStr(text, "VOLLEYBALL") > 0 Then
        GuessType = "Volleyball"
    Else
        GuessType = "Event"
    End If
End Function

Private Function ReplaceReports(ByVal eventId As Long, ByVal pricePath As String, ByVal tipsPath As String, ByVal priceRows As Collection, ByVal tipRows As Collection) As Long
    Dim importId As Long
    EnsureImportColumns
    Exec "DELETE FROM tblSale WHERE EventID=" & eventId
    Exec "DELETE FROM tblTip WHERE EventID=" & eventId
    Exec "DELETE FROM tblInvoiceLine WHERE InvoiceID In (SELECT InvoiceID FROM tblInvoice WHERE EventID=" & eventId & ")"
    Exec "DELETE FROM tblInvoice WHERE EventID=" & eventId
    Exec "DELETE FROM tblStaffTip WHERE EventID=" & eventId
    Exec "DELETE FROM tblException WHERE EventID=" & eventId
    Exec "UPDATE [tblImport] SET [IsCurrent]=0 WHERE [EventID]=" & eventId
    Exec "INSERT INTO tblImport (EventID, PriceFile, TipsFile, ImportedAt, IsCurrent) VALUES (" & _
        eventId & "," & SqlText(Left$(pricePath, 255)) & "," & SqlText(Left$(tipsPath, 255)) & "," & _
        "#" & Format$(Now, "mm/dd/yyyy hh:nn:ss") & "#,-1)"
    importId = CLng(Scalar("SELECT Max(ImportID) FROM tblImport WHERE EventID=" & eventId))
    LoadPrice eventId, importId, priceRows
    LoadTips eventId, importId, tipRows
    Reconcile eventId, importId
    SyncAssignments eventId, importId
    ReplaceReports = importId
End Function

Private Sub LoadPrice(ByVal eventId As Long, ByVal importId As Long, ByVal rows As Collection)
    Dim row As Collection
    Dim locationName As String
    Dim label As String
    Dim code As String
    Dim categoryName As String
    Dim bucket As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    locationName = ""
    For Each row In rows
        If FieldAt(row, 2) Like "Location:*" Then
            locationName = Trim$(Mid$(FieldAt(row, 2), 10))
        ElseIf Len(locationName) > 0 And Right$(UCase$(FieldAt(row, 3)), 6) = "TOTAL:" Then
            label = Trim$(Left$(FieldAt(row, 3), Len(FieldAt(row, 3)) - 6))
            If InStr(label, " - ") > 0 Then
                code = Trim$(Left$(label, InStr(label, " - ") - 1))
                categoryName = Trim$(Mid$(label, InStr(label, " - ") + 3))
                bucket = BucketOf(code)
                If bucket = "OTHER" And Len(code) > 0 Then
                    If Not seen.Exists(code) Then
                        seen(code) = True
                        LogException eventId, importId, "Unknown category", code & " " & categoryName & " was treated as food / non-alcoholic."
                    End If
                End If
                Exec "INSERT INTO tblSale (ImportID, EventID, LocationName, CategoryCode, CategoryName, Bucket, GrossSales, NetSales) VALUES (" & _
                    importId & "," & eventId & "," & SqlText(Left$(locationName, 150)) & "," & SqlText(Left$(code, 30)) & "," & _
                    SqlText(Left$(categoryName, 120)) & "," & SqlText(bucket) & "," & MoneyValue(FieldAt(row, 16)) & "," & MoneyValue(FieldAt(row, 19)) & ")"
            End If
        End If
    Next row
End Sub

Private Sub LoadTips(ByVal eventId As Long, ByVal importId As Long, ByVal rows As Collection)
    Dim row As Collection
    Dim departmentCode As String
    Dim departmentName As String
    Dim text As String
    Dim locationName As String
    departmentCode = ""
    departmentName = ""
    For Each row In rows
        text = FieldAt(row, 1)
        If Left$(UCase$(text), 11) = "DEPARTMENT:" Then
            text = Trim$(Mid$(text, 12))
            If InStr(text, " - ") > 0 Then
                departmentCode = Trim$(Left$(text, InStr(text, " - ") - 1))
                departmentName = Trim$(Mid$(text, InStr(text, " - ") + 3))
            End If
        ElseIf Right$(UCase$(FieldAt(row, 2)), 6) = "TOTAL:" And InStr(FieldAt(row, 2), " - ") > 0 Then
            locationName = Trim$(Left$(FieldAt(row, 2), Len(FieldAt(row, 2)) - 6))
            Exec "INSERT INTO tblTip (ImportID, EventID, DepartmentCode, DepartmentName, LocationName, SalesAmount, TipsAmount, AutoTips, TotalTips) VALUES (" & _
                importId & "," & eventId & "," & SqlText(Left$(departmentCode, 20)) & "," & SqlText(Left$(departmentName, 80)) & "," & _
                SqlText(Left$(locationName, 150)) & "," & MoneyValue(FieldAt(row, 5)) & "," & MoneyValue(FieldAt(row, 8)) & "," & _
                MoneyValue(FieldAt(row, 10)) & "," & MoneyValue(FieldAt(row, 13)) & ")"
        End If
    Next row
End Sub

Private Sub Reconcile(ByVal eventId As Long, ByVal importId As Long)
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT t.LocationName, t.SalesAmount, Nz((SELECT Sum(s.GrossSales) FROM tblSale AS s WHERE s.ImportID=t.ImportID AND s.LocationName=t.LocationName),0) AS PriceGross " & _
        "FROM tblTip AS t WHERE t.ImportID=" & importId, EMS_SNAPSHOT)
    Do While Not rs.EOF
        If Abs(CCur(Nz(rs!SalesAmount, 0)) - CCur(Nz(rs!PriceGross, 0))) > 0.01 Then
            LogException eventId, importId, "Sales mismatch", CStr(rs!LocationName) & " tips report " & MoneyText(CCur(rs!SalesAmount)) & ", price report " & MoneyText(CCur(rs!PriceGross)) & ". Both amounts were kept."
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub SyncAssignments(ByVal eventId As Long, ByVal importId As Long)
    Dim rs As Object
    Dim locationName As String
    Dim gross As Currency
    Dim tips As Currency
    Dim foodNet As Currency
    Dim naNet As Currency
    Dim alcoholNet As Currency
    Dim departmentName As String
    Dim existing As Variant
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT LocationName FROM tblSale WHERE ImportID=" & importId & _
        " UNION SELECT LocationName FROM tblTip WHERE ImportID=" & importId, EMS_SNAPSHOT)
    Do While Not rs.EOF
        locationName = CStr(rs!LocationName)
        gross = CCur(Nz(Scalar("SELECT Sum(GrossSales) FROM tblSale WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName)), 0))
        tips = CCur(Nz(Scalar("SELECT Sum(TotalTips) FROM tblTip WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName)), 0))
        departmentName = Nz(Scalar("SELECT Top 1 DepartmentName FROM tblTip WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName)), "")
        ReadNets importId, locationName, foodNet, naNet, alcoholNet
        EnsureLocation eventId, importId, locationName, departmentName
        existing = Scalar("SELECT AssignID FROM tblAssign WHERE EventID=" & eventId & " AND LocationName=" & SqlText(locationName))
        If IsNull(existing) Then
            InsertAssignment eventId, importId, locationName, departmentName, gross, tips, foodNet, naNet, alcoholNet
        Else
            Exec "UPDATE tblAssign SET PosSales=" & gross & ", PosTips=" & tips & _
                ", FoodNet=" & foodNet & ", NaNet=" & naNet & ", AlcoholNet=" & alcoholNet & _
                ", DepartmentName=" & SqlText(Left$(departmentName, 80)) & " WHERE AssignID=" & existing
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT AssignID, LocationName FROM tblAssign WHERE EventID=" & eventId, EMS_SNAPSHOT)
    Do While Not rs.EOF
        If IsNull(Scalar("SELECT Top 1 SaleID FROM tblSale WHERE ImportID=" & importId & " AND LocationName=" & SqlText(CStr(rs!LocationName)))) And _
           IsNull(Scalar("SELECT Top 1 TipID FROM tblTip WHERE ImportID=" & importId & " AND LocationName=" & SqlText(CStr(rs!LocationName)))) Then
            Exec "UPDATE tblAssign SET PosSales=0, PosTips=0, FoodNet=0, NaNet=0, AlcoholNet=0 WHERE AssignID=" & rs!AssignID
            LogException eventId, importId, "Missing stand", CStr(rs!LocationName) & " was not in the new reports. Its assignment was kept and the POS amounts were set to zero."
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub EnsureLocation(ByVal eventId As Long, ByVal importId As Long, ByVal locationName As String, ByVal departmentName As String)
    Dim code As String
    If Not IsNull(Scalar("SELECT LocationID FROM tblLocation WHERE LocationName=" & SqlText(locationName))) Then Exit Sub
    code = Nz(Scalar("SELECT Top 1 DepartmentCode FROM tblTip WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName)), "")
    Exec "INSERT INTO tblLocation (LocationName, DepartmentCode, DepartmentName, Active) VALUES (" & _
        SqlText(Left$(locationName, 150)) & "," & SqlText(Left$(code, 20)) & "," & SqlText(Left$(departmentName, 80)) & ",-1)"
    LogException eventId, importId, "New location", locationName & " was not on the stand list. It was added so it can be assigned."
End Sub

Private Sub ReadNets(ByVal importId As Long, ByVal locationName As String, ByRef foodNet As Currency, ByRef naNet As Currency, ByRef alcoholNet As Currency)
    Dim rs As Object
    foodNet = 0
    naNet = 0
    alcoholNet = 0
    Set rs = CurrentDb.OpenRecordset("SELECT Bucket, Sum(NetSales) AS NetAmount FROM tblSale WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName) & " GROUP BY Bucket", EMS_SNAPSHOT)
    Do While Not rs.EOF
        Select Case UCase$(CStr(Nz(rs!Bucket, "")))
            Case "FOOD", "OTHER"
                foodNet = foodNet + CCur(Nz(rs!NetAmount, 0))
            Case "NA"
                naNet = naNet + CCur(Nz(rs!NetAmount, 0))
            Case "BEER", "LIQUOR"
                alcoholNet = alcoholNet + CCur(Nz(rs!NetAmount, 0))
        End Select
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub InsertAssignment(ByVal eventId As Long, ByVal importId As Long, ByVal locationName As String, ByVal departmentName As String, ByVal gross As Currency, ByVal tips As Currency, ByVal foodNet As Currency, ByVal naNet As Currency, ByVal alcoholNet As Currency)
    Dim prior As Object
    Dim role As String
    Dim payeeSql As String
    Dim headcount As Long
    Dim names As String
    Dim code As String
    role = "OPEN"
    payeeSql = "Null"
    headcount = 0
    names = ""
    Set prior = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 Role, PayeeID, Headcount, StaffNames FROM tblAssign WHERE LocationName=" & SqlText(locationName) & _
        " AND EventID<>" & eventId & " AND Role<>'OPEN' ORDER BY EventID DESC", EMS_SNAPSHOT)
    If Not prior.EOF Then
        role = CStr(Nz(prior!role, "OPEN"))
        If Not IsNull(prior!PayeeID) Then payeeSql = CStr(prior!PayeeID)
        headcount = CLng(Nz(prior!Headcount, 0))
        names = CStr(Nz(prior!StaffNames, ""))
    End If
    prior.Close
    Exec "INSERT INTO tblAssign (EventID, LocationName, DepartmentName, PosSales, FoodNet, NaNet, AlcoholNet, PosTips, Role, PayeeID, Headcount, StaffNames, CookFee, Bonus, Shortage, Cleaning) VALUES (" & _
        eventId & "," & SqlText(Left$(locationName, 150)) & "," & SqlText(Left$(departmentName, 80)) & "," & gross & "," & foodNet & "," & naNet & "," & alcoholNet & "," & tips & "," & _
        SqlText(role) & "," & payeeSql & "," & headcount & "," & SqlText(names) & ",0,0,0,0)"
    code = Nz(Scalar("SELECT Top 1 DepartmentCode FROM tblTip WHERE ImportID=" & importId & " AND LocationName=" & SqlText(locationName)), "")
    If role = "OPEN" And (code = "36637" Or InStr(1, locationName, " SUB", vbTextCompare) > 0) Then
        LogException eventId, importId, "Subcontractor?", locationName & " looks like a subcontracted stand. It was left open so the payee can be chosen."
    End If
End Sub

Private Sub ReadHeader(ByVal rows As Collection, ByRef eventName As String, ByRef eventDay As Date, ByRef found As Boolean)
    Dim row As Collection
    Dim item As Variant
    Dim line As Variant
    Dim text As String
    found = False
    eventName = ""
    For Each row In rows
        For Each item In row
            text = Replace(CStr(item), vbCr, "")
            If InStr(1, text, "Event:", vbTextCompare) > 0 And InStr(1, text, "Start Date:", vbTextCompare) > 0 Then
                For Each line In Split(text, vbLf)
                    If Left$(UCase$(Trim$(CStr(line))), 6) = "EVENT:" Then eventName = Trim$(Mid$(Trim$(CStr(line)), 7))
                    If Left$(UCase$(Trim$(CStr(line))), 11) = "START DATE:" Then eventDay = ParseUsDate(Trim$(Mid$(Trim$(CStr(line)), 12)))
                Next line
                found = Len(eventName) > 0 And eventDay > 0
                If found Then Exit Sub
            End If
        Next item
    Next row
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadTextFile = stream.ReadText
    stream.Close
    If Left$(ReadTextFile, 1) = ChrW(65279) Then ReadTextFile = Mid$(ReadTextFile, 2)
End Function

Private Function ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection
    Dim row As New Collection
    Dim field As String
    Dim index As Long
    Dim char As String
    Dim quoted As Boolean
    index = 1
    Do While index <= Len(text)
        char = Mid$(text, index, 1)
        If quoted Then
            If char = """" Then
                If Mid$(text, index + 1, 1) = """" Then
                    field = field & """"
                    index = index + 1
                Else
                    quoted = False
                End If
            Else
                field = field & char
            End If
        ElseIf char = """" Then
            quoted = True
        ElseIf char = "," Then
            row.Add field
            field = ""
        ElseIf char = vbCr Then
            ' The line feed closes the row.
        ElseIf char = vbLf Then
            row.Add field
            field = ""
            If RowHasText(row) Then rows.Add row
            Set row = New Collection
        Else
            field = field & char
        End If
        index = index + 1
    Loop
    If Len(field) > 0 Or row.Count > 0 Then
        row.Add field
        If RowHasText(row) Then rows.Add row
    End If
    Set ParseCsv = rows
End Function

Private Function RowHasText(ByVal row As Collection) As Boolean
    Dim item As Variant
    For Each item In row
        If Len(Trim$(CStr(item))) > 0 Then
            RowHasText = True
            Exit Function
        End If
    Next item
End Function

Private Function FieldAt(ByVal row As Collection, ByVal index As Long) As String
    If index >= 1 And index <= row.Count Then FieldAt = CStr(row(index))
End Function
