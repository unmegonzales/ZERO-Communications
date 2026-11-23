Attribute VB_Name = "EMS_Common"
Option Compare Database
Option Explicit

' Shared helpers. Money that is split (gratuity and staff) is cut off at the
' cent and the leftover stays with Levy. Commission lines round half up.

Public gPageFooter As String

Public Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(Nz(value, ""), "'", "''") & "'"
End Function

Public Sub Exec(ByVal sql As String)
    On Error GoTo Failed
    CurrentDb.Execute sql, dbFailOnError
    Exit Sub
Failed:
    Err.Raise Err.Number, "EMS", Err.Description & vbCrLf & Left$(sql, 240)
End Sub

Public Function Scalar(ByVal sql As String) As Variant
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then Scalar = Null Else Scalar = rs.Fields(0).Value
    rs.Close
End Function

Public Function MoneyRound(ByVal value As Variant) As Currency
    Dim amount As Variant
    amount = CDec(Nz(value, 0))
    If amount >= 0 Then
        MoneyRound = CCur(Int(amount * CDec(100) + CDec("0.5")) / 100)
    Else
        MoneyRound = CCur(-Int((-amount) * CDec(100) + CDec("0.5")) / 100)
    End If
End Function

Public Function FloorMoney(ByVal value As Variant) As Currency
    Dim amount As Variant
    amount = CDec(Nz(value, 0))
    If amount >= 0 Then
        FloorMoney = CCur(Int(amount * CDec(100)) / 100)
    Else
        FloorMoney = CCur(-Int((-amount) * CDec(100)) / 100)
    End If
End Function

' Half-up product of a currency amount and a rate. Used for commission and fees.
Public Function MoneyProduct(ByVal amount As Currency, ByVal rate As Double) As Currency
    Dim cents As Variant
    cents = CDec(amount) * CDec(100) * RateDecimal(rate)
    If cents >= 0 Then
        MoneyProduct = CCur(Int(cents + CDec("0.5")) / 100)
    Else
        MoneyProduct = CCur(-Int((-cents) + CDec("0.5")) / 100)
    End If
End Function

' Cuts the product off at the cent. Used for the payable share of a tip pool.
Public Function FloorProduct(ByVal amount As Currency, ByVal rate As Double) As Currency
    Dim cents As Variant
    cents = CDec(amount) * CDec(100) * RateDecimal(rate)
    If cents >= 0 Then
        FloorProduct = CCur(Int(cents) / 100)
    Else
        FloorProduct = CCur(-Int(-cents) / 100)
    End If
End Function

Private Function RateDecimal(ByVal rate As Double) As Variant
    Dim text As String
    text = Format$(rate, "0.000000")
    If Right$(text, 5) = "99999" Then text = Format$(rate + 0.0000005, "0.000000")
    RateDecimal = CDec(text)
End Function

Public Function MoneyText(ByVal amount As Currency, Optional ByVal blankZero As Boolean = False) As String
    If amount = 0 And blankZero Then
        MoneyText = "$ -"
    Else
        MoneyText = Format$(amount, "$#,##0.00;($#,##0.00)")
    End If
End Function

Public Function RateOn(ByVal payeeId As Long, ByVal code As String, ByVal eventDay As Date, ByVal fallback As Double) As Double
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 RateAmount FROM tblPayeeRate WHERE PayeeID=" & payeeId & _
        " AND RateCode=" & SqlText(code) & _
        " AND EffectiveDate<=#" & Format$(eventDay, "mm/dd/yyyy") & "#" & _
        " AND (EndDate Is Null OR EndDate>=#" & Format$(eventDay, "mm/dd/yyyy") & "#)" & _
        " ORDER BY EffectiveDate DESC", dbOpenSnapshot)
    If rs.EOF Then RateOn = fallback Else RateOn = CDbl(Nz(rs!RateAmount, fallback))
    rs.Close
End Function

Public Function EventKey(ByVal eventName As String) As String
    Dim text As String
    Dim paren As Long
    text = UCase$(Trim$(eventName))
    paren = InStrRev(text, "(")
    If paren > 1 And Right$(text, 1) = ")" Then
        If IsNumeric(Trim$(Mid$(text, paren + 1, Len(text) - paren - 1))) Then
            text = Trim$(Left$(text, paren - 1))
        End If
    End If
    EventKey = text
End Function

Public Function ParseUsDate(ByVal text As String) As Date
    Dim parts As Variant
    Dim yearNumber As Long
    text = Trim$(Split(Trim$(text), " ")(0))
    parts = Split(text, "/")
    If UBound(parts) < 2 Then Err.Raise vbObjectError + 1, "EMS", "Could not read the date " & text
    yearNumber = CLng(parts(2))
    If yearNumber < 100 Then yearNumber = 2000 + yearNumber
    ParseUsDate = DateSerial(yearNumber, CLng(parts(0)), CLng(parts(1)))
End Function

Public Function BucketOf(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "313000", "FF-313076", "353008"
            BucketOf = "FOOD"
        Case "312115", "FF-312114"
            BucketOf = "NA"
        Case "313082"
            BucketOf = "BEER"
        Case "313004", "313096"
            BucketOf = "LIQUOR"
        Case Else
            BucketOf = "OTHER"
    End Select
End Function

Public Function MoneyValue(ByVal text As String) As Currency
    text = Trim$(Replace(Replace(Replace(text, "$", ""), ",", ""), " ", ""))
    If Len(text) = 0 Or text = "-" Then
        MoneyValue = 0
    Else
        MoneyValue = CCur(text)
    End If
End Function

Public Sub LogException(ByVal eventId As Long, ByVal importId As Long, ByVal kind As String, ByVal detail As String)
    Dim importSql As String
    If importId = 0 Then importSql = "Null" Else importSql = CStr(importId)
    Exec "INSERT INTO tblException (EventID, ImportID, Kind, Detail) VALUES (" & _
        eventId & "," & importSql & "," & SqlText(Left$(kind, 40)) & "," & SqlText(Left$(detail, 4000)) & ")"
End Sub

Public Function EMSFolder() As String
    Dim folder As String
    folder = CurrentProject.Path
    If Len(folder) = 0 Or Len(folder) > 180 Then folder = Environ$("USERPROFILE") & "\Documents"
    folder = folder & "\EMS_Statements"
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    EMSFolder = folder
End Function

Public Function PickFile(ByVal title As String) As String
    On Error GoTo Typed
    Dim dialog As Object
    Set dialog = Application.FileDialog(3)
    dialog.title = title
    dialog.AllowMultiSelect = False
    dialog.Filters.Clear
    dialog.Filters.Add "CSV reports", "*.csv"
    If dialog.Show = -1 Then PickFile = dialog.SelectedItems(1)
    Exit Function
Typed:
    PickFile = Trim$(InputBox(title & vbCrLf & "Type the full path to the CSV."))
End Function

Public Function InvoiceNumber(ByVal eventDay As Date, ByVal abbreviation As String, ByVal payeeId As Long) As String
    Dim code As String
    code = UCase$(Trim$(abbreviation))
    If Len(code) = 0 Then code = "P" & payeeId
    InvoiceNumber = "UNM-" & Format$(eventDay, "mmddyy") & "-" & code
End Function

Public Sub FieldCaption(ByVal tableName As String, ByVal fieldName As String, ByVal caption As String)
    Dim fld As DAO.Field
    Dim prop As DAO.Property
    Set fld = CurrentDb.TableDefs(tableName).Fields(fieldName)
    On Error Resume Next
    fld.Properties("Caption") = caption
    If Err.Number <> 0 Then
        Err.Clear
        Set prop = CurrentDb.CreateProperty("Caption", dbText, caption)
        fld.Properties.Append prop
    End If
    On Error GoTo 0
End Sub
