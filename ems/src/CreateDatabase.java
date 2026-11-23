import com.healthmarketscience.jackcess.ColumnBuilder;
import com.healthmarketscience.jackcess.DataType;
import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.DateTimeType;
import com.healthmarketscience.jackcess.Table;
import com.healthmarketscience.jackcess.TableBuilder;

import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

/** Builds the empty EMS payout database and seeds payees, rates, and stands. */
public class CreateDatabase {
    public static void main(String[] args) throws Exception {
        Path file = Path.of(args[0]);
        Path locations = Path.of(args[1]);
        if (Files.exists(file)) Files.delete(file);
        Database db = DatabaseBuilder.create(Database.FileFormat.V2016, file.toFile());
        db.setDateTimeType(DateTimeType.LOCAL_DATE_TIME);

        table(db, "tblPayee",
                auto("PayeeID"),
                text("PayeeName", 200),
                text("PayeeType", 20),
                text("Abbreviation", 20),
                text("Address1", 255),
                text("Address2", 255),
                text("SAPID", 30),
                text("BSSID", 30),
                text("Phone", 40),
                memo("Notes"),
                flag("Active"));
        table(db, "tblPayeeEmail",
                auto("EmailID"),
                num("PayeeID"),
                text("EmailAddress", 255),
                text("Label", 40),
                flag("IsPrimary"));
        table(db, "tblPayeeRate",
                auto("RateID"),
                num("PayeeID"),
                text("RateCode", 80),
                dbl("RateAmount"),
                date("EffectiveDate"),
                date("EndDate"),
                memo("Notes"));
        table(db, "tblLocation",
                auto("LocationID"),
                text("LocationName", 150),
                text("DepartmentCode", 20),
                text("DepartmentName", 80),
                flag("Active"));
        table(db, "tblOffice",
                auto("OfficeID"),
                text("BrandLine", 120),
                text("PreparedBy", 80),
                text("PreparedTitle", 80),
                text("ReviewedBy", 80),
                text("ReviewedTitle", 80),
                text("ApprovedBy", 80),
                text("ApprovedTitle", 80),
                text("CompanyCode", 20),
                text("CostCenter", 20));
        table(db, "tblEvent",
                auto("EventID"),
                date("EventDate"),
                text("EventName", 200),
                text("EventType", 80),
                text("Venue", 120),
                dbl("TaxRate"),
                money("MinimumCommission"),
                text("StatusText", 30),
                text("PreparedBy", 80),
                text("PreparedTitle", 80),
                date("PreparedDate"),
                text("ReviewedBy", 80),
                text("ReviewedTitle", 80),
                date("ReviewedDate"),
                text("ApprovedBy", 80),
                text("ApprovedTitle", 80),
                date("ApprovedDate"),
                memo("Notes"),
                date("CreatedAt"));
        table(db, "tblImport",
                auto("ImportID"),
                num("EventID"),
                text("PriceFile", 255),
                text("TipsFile", 255),
                date("ImportedAt"),
                flag("IsCurrent"));
        table(db, "tblSale",
                auto("SaleID"),
                num("ImportID"),
                num("EventID"),
                text("LocationName", 150),
                text("CategoryCode", 30),
                text("CategoryName", 120),
                text("Bucket", 20),
                money("GrossSales"),
                money("NetSales"));
        table(db, "tblTip",
                auto("TipID"),
                num("ImportID"),
                num("EventID"),
                text("DepartmentCode", 20),
                text("DepartmentName", 80),
                text("LocationName", 150),
                money("SalesAmount"),
                money("TipsAmount"),
                money("AutoTips"),
                money("TotalTips"));
        table(db, "tblAssign",
                auto("AssignID"),
                num("EventID"),
                text("LocationName", 150),
                text("DepartmentName", 80),
                money("PosSales"),
                money("FoodNet"),
                money("NaNet"),
                money("AlcoholNet"),
                money("PosTips"),
                text("Role", 20),
                num("PayeeID"),
                num("Headcount"),
                memo("StaffNames"),
                money("CookFee"),
                money("Bonus"),
                money("Shortage"),
                money("Cleaning"));
        table(db, "tblVendorInput",
                auto("InputID"),
                num("EventID"),
                num("PayeeID"),
                money("CashTender"),
                money("CardTender"),
                money("CorporateVoucher"),
                money("LevyNfpVoucher"),
                money("UnmGuestVoucher"),
                dbl("IceBags"),
                money("OtherExpense"));
        table(db, "tblInvoice",
                auto("InvoiceID"),
                num("EventID"),
                num("PayeeID"),
                text("InvoiceNumber", 40),
                text("Kind", 16),
                money("GrossSales"),
                money("FoodNet"),
                money("AlcoholNet"),
                money("GroupNet"),
                money("FoodCommission"),
                money("AlcoholCommission"),
                money("Commission"),
                money("Gratuity"),
                money("LocationAdjustments"),
                money("Supplements"),
                money("AmountDue"),
                money("CardFee"),
                money("IceAmount"),
                money("Shortage"),
                money("OtherExpense"),
                memo("Notes"));
        table(db, "tblInvoiceLine",
                auto("LineID"),
                num("InvoiceID"),
                text("LocationName", 150),
                num("SortOrder"),
                money("NetSales"),
                money("FoodNet"),
                money("FoodCommission"),
                money("AlcoholNet"),
                money("AlcoholCommission"),
                money("Commission"),
                money("Gratuity"),
                money("Adjustments"),
                money("Payout"));
        table(db, "tblSupplement",
                auto("SupplementID"),
                num("EventID"),
                num("PayeeID"),
                text("Label", 80),
                money("Amount"),
                text("GLCode", 20));
        table(db, "tblStaffTip",
                auto("StaffTipID"),
                num("EventID"),
                text("LocationName", 150),
                num("SlotNo"),
                text("PersonName", 80),
                money("Amount"),
                money("PoolAmount"),
                money("HeldBack"));
        table(db, "tblException",
                auto("ExceptionID"),
                num("EventID"),
                num("ImportID"),
                text("Kind", 40),
                memo("Detail"));
        table(db, "tblPrintHead",
                auto("PrintID"),
                text("Kind", 16),
                text("PayeeName", 200),
                text("AddressBlock", 255),
                text("StatementDate", 20),
                text("InvoiceNumber", 40),
                text("AmountDue", 24),
                text("EventWhen", 40),
                text("EventType", 80),
                text("Abbreviation", 20),
                text("SAPID", 30),
                text("BSSID", 30),
                text("StatusText", 30),
                text("FoodNet", 24),
                text("AlcoholNet", 24),
                text("GroupNet", 24),
                text("GrossSales", 24),
                memo("Notes"),
                text("PreparedBy", 80),
                text("PreparedTitle", 80),
                text("PreparedWhen", 20),
                text("ReviewedBy", 80),
                text("ReviewedTitle", 80),
                text("ReviewedWhen", 20),
                text("ApprovedBy", 80),
                text("ApprovedTitle", 80),
                text("ApprovedWhen", 20),
                text("EventWhenShort", 20),
                text("FoodRateLabel", 24),
                text("AlcoholRateLabel", 24),
                flag("ShowAdjustments"),
                memo("Extra"));
        table(db, "tblPrintLine",
                auto("PrintLineID"),
                num("SortOrder"),
                text("ServiceArea", 150),
                text("NetSales", 24),
                text("FoodNet", 24),
                text("FoodCommission", 24),
                text("AlcoholNet", 24),
                text("AlcoholCommission", 24),
                text("Commission", 24),
                text("Gratuity", 24),
                text("Adjustments", 24),
                text("Payout", 24));
        table(db, "tblPrintExtra",
                auto("PrintExtraID"),
                num("SortOrder"),
                text("Label", 80),
                text("Amount", 24),
                text("GLText", 80));

        Table office = db.getTable("tblOffice");
        office.addRow(null, "COMPASS GROUP USA | LEVY - UNM ATHLETICS",
                "Abigail Tenorio", "Concessions Manager",
                "Ray Velasquez", "Director of Operations",
                "Evan E. Gonzales", "Finance Manager",
                "241", "36127");

        int arvc = payee(db, "Albuquerque Rebels Volleyball Club", "NPO", "ARVC",
                "7800 LAS LOMITAS AVENUE NE", "Albuquerque, NM 87113-2645", "1013051", "5487");
        email(db, arvc, "gabe.padilla07@yahoo.com", "Primary", true);
        email(db, arvc, "jillianshane@sbcglobal.net", "Additional", false);
        npoRates(db, arvc);

        int dcva = payee(db, "Duke City Volleyball Athletics Inc", "NPO", "DCVA",
                "8717 DEMOCRACY NE", "Albuquerque, NM 87109-5183", "1026871", "10468");
        email(db, dcva, "fundraising@dcvanm.com", "Primary", true);
        npoRates(db, dcva);

        int rghs = payee(db, "APS - Rio Grande High School", "NPO", "RGHS",
                "2300 ARENAL RD SW", "Albuquerque, NM 87105-4180", "1029380", "11037");
        email(db, rghs, "christopher.smalls@aps.edu", "Primary", true);
        npoRates(db, rghs);

        int usaf = payee(db, "58 SOW TOP 3 ASSOCIATION", "NPO", "USAF",
                "5249 HERCULES WY SE", "Albuquerque, NM 87117-0001", "1031810", "11773");
        email(db, usaf, "daniel.gonzalez.23@us.af.mil", "Primary", true);
        npoRates(db, usaf);

        int dnhs = payee(db, "APS - Del Norte High School Cheer", "NPO", "DNHS",
                "", "", "1027902", "10743");
        npoRates(db, dnhs);

        int e3 = payee(db, "DJCF INC. (DBA: E3VB)", "NPO", "E3VB",
                "6919 LUZ DEL SOL PL NW", "Albuquerque, NM 87114-3702", "1031901", "11793");
        email(db, e3, "smith5411@gmail.com", "Primary", true);
        npoRates(db, e3);

        int vict = payee(db, "VICTORY OUTREACH INC., ROSWELL", "NPO", "VICT",
                "", "", "1029375", "11040");
        npoRates(db, vict);

        int uf = payee(db, "UNITED FIT", "NPO", "UF", "", "", "982165", "5190");
        npoRates(db, uf);

        int laguna = payee(db, "Fuego Enterprises LLC", "SUB", "LAG",
                "7770 JEFFERSON ST NE STE 302", "Albuquerque, NM 87109-4386", "992855", "5177");
        email(db, laguna, "FAlvarez@fuegonm.com", "Primary", true);
        subRates(db, laguna, "0.70", "0.05", "0", "0");

        int sugar = payee(db, "NXTLVL HOLDINGS LLC SUGAR SHIVERS", "SUB", "SUGAR",
                "533 61ST ST NW", "Albuquerque, NM 87105-1407", "1037451", "13470");
        email(db, sugar, "nxtlevelholdingsllc@gmail.com", "Primary", true);
        subRates(db, sugar, "0.80", "0.10", "0", "0");

        int nood = payee(db, "BNW Enterprises LLC", "SUB", "NOOD", "", "", "", "");
        subRates(db, nood, "0.90", "0.10", "0", "0");

        Table locs = db.getTable("tblLocation");
        for (String line : Files.readAllLines(locations)) {
            if (line.isBlank()) continue;
            String[] p = line.split("\t", -1);
            locs.addRow(null, p[0], p[1], p[2], "1".equals(p[3]));
        }
        db.close();
        System.out.println("created " + file);
    }

    private static int payee(Database db, String name, String type, String abbr,
                             String a1, String a2, String sap, String bss) throws Exception {
        Table t = db.getTable("tblPayee");
        t.addRow(null, name, type, abbr, a1, a2, sap, bss, "", "", true);
        for (var row : t) {
            if (abbr.equals(row.get("Abbreviation")) && name.equals(row.get("PayeeName"))) {
                return ((Number) row.get("PayeeID")).intValue();
            }
        }
        throw new IllegalStateException("payee not found " + abbr);
    }

    private static void email(Database db, int payeeId, String address, String label, boolean primary) throws Exception {
        db.getTable("tblPayeeEmail").addRow(null, payeeId, address, label, primary);
    }

    private static void npoRates(Database db, int payeeId) throws Exception {
        rate(db, payeeId, "DONATION_FOOD", "0.10");
        rate(db, payeeId, "DONATION_ALCOHOL", "0.08");
        rate(db, payeeId, "TIP_WITHHOLD", "0.10");
    }

    private static void subRates(Database db, int payeeId, String food, String na, String beer, String liquor) throws Exception {
        rate(db, payeeId, "FOOD", food);
        rate(db, payeeId, "NA", na);
        rate(db, payeeId, "BEER", beer);
        rate(db, payeeId, "LIQUOR", liquor);
        rate(db, payeeId, "TIP_WITHHOLD", "0.10");
        rate(db, payeeId, "CARD_FEE", "0.03");
        rate(db, payeeId, "ICE_PER_BAG", "3");
    }

    private static void rate(Database db, int payeeId, String code, String amount) throws Exception {
        db.getTable("tblPayeeRate").addRow(null, payeeId, code, Double.valueOf(amount),
                local(LocalDate.of(2026, 1, 1)), null, "");
    }

    private static ColumnBuilder auto(String name) {
        return new ColumnBuilder(name, DataType.LONG).setAutoNumber(true);
    }

    private static ColumnBuilder num(String name) {
        return new ColumnBuilder(name, DataType.LONG);
    }

    private static ColumnBuilder text(String name, int len) {
        return new ColumnBuilder(name, DataType.TEXT).setLength(len);
    }

    private static ColumnBuilder memo(String name) {
        return new ColumnBuilder(name, DataType.MEMO);
    }

    private static ColumnBuilder flag(String name) {
        return new ColumnBuilder(name, DataType.BOOLEAN);
    }

    private static ColumnBuilder dbl(String name) {
        return new ColumnBuilder(name, DataType.DOUBLE);
    }

    private static ColumnBuilder money(String name) {
        return new ColumnBuilder(name, DataType.MONEY);
    }

    private static ColumnBuilder date(String name) {
        return new ColumnBuilder(name, DataType.SHORT_DATE_TIME);
    }

    private static void table(Database db, String name, ColumnBuilder... cols) throws Exception {
        TableBuilder b = new TableBuilder(name);
        String key = cols[0].getName();
        for (ColumnBuilder c : cols) b.addColumn(c);
        b.setPrimaryKey(key);
        b.toTable(db);
    }

    private static LocalDateTime local(LocalDate day) {
        return day.atStartOfDay();
    }

    @SuppressWarnings("unused")
    private static Date unused(BigDecimal v) {
        return Date.from(LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant());
    }
}
