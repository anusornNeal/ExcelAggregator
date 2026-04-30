import data.reader.InvoiceExcelReader
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import kotlin.io.path.createTempFile
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class InvoiceExcelReaderTest {

    @Test
    fun readsInvoiceFromPivotSummaryWhenMainPriceCellsContainErrors() {
        val file = createTempFile(prefix = "pivot-summary-error-price-", suffix = ".xlsx").toFile()
        XSSFWorkbook().use { workbook ->
            val sheet = workbook.createSheet("ST44064")
            sheet.row(0).text(2, "Invoice - ST44064")
            sheet.row(1).text(2, "Date").text(3, "22/4/2026")

            sheet.row(7)
                .text(0, "No.")
                .text(1, "Code / List")
                .text(2, "SKU")
                .text(3, "Price")
                .text(4, "Amount (Pieces)")
                .text(5, "*Remark")
            sheet.row(8).text(4, "Shop")
            sheet.row(9)
                .number(0, 1.0)
                .text(2, "2503TPN04487-cream")
                .error(3)
                .number(4, 3.0)

            sheet.row(9)
                .text(8, "Row Labels")
                .text(9, "Sum of Amount (Pieces)")
                .text(10, "Sum of มูลค่า")
            sheet.row(10).number(8, 490.0).number(9, 6.0).number(10, 2940.0)
            sheet.row(11).number(8, 550.0).number(9, 5.0).number(10, 2750.0)
            sheet.row(12).number(8, 590.0).number(9, 8.0).number(10, 4720.0)
            sheet.row(13).number(8, 690.0).number(9, 5.0).number(10, 3450.0)
            sheet.row(14).text(8, "Grand Total").number(9, 24.0).number(10, 13860.0)

            FileOutputStream(file).use(workbook::write)
        }

        val sheets = InvoiceExcelReader.readAll(file)

        assertEquals(1, sheets.size)
        assertEquals("ST44064", sheets.single().branch)
        assertEquals("ST44064", sheets.single().sourceSheetName)
        assertEquals(5, sheets.single().rows.size)
        assertEquals(490.0, sheets.single().rows.first().price)
        assertEquals(6.0, sheets.single().rows.first().qty)
        assertEquals(2940.0, sheets.single().rows.first().value)
        assertTrue(sheets.single().rows.last().isTotal)
        assertEquals(24.0, sheets.single().rows.last().qty)
        assertEquals(13860.0, sheets.single().rows.last().value)
    }

    @Test
    fun calculatesPivotSummaryValueWhenOnlyShopColumnExists() {
        val file = createTempFile(prefix = "pivot-summary-shop-", suffix = ".xlsx").toFile()
        XSSFWorkbook().use { workbook ->
            val sheet = workbook.createSheet("ST43135")
            sheet.row(0).text(2, "Invoice - ST43135")
            sheet.row(1).text(2, "Branch").text(3, "Counter W")

            sheet.row(6).text(7, "Row Labels").text(8, "Sum of Amount (Pieces)").text(9, "Sum of Shop")
            sheet.row(7).number(7, 450.0).number(8, 30.0).number(9, 13500.0)
            sheet.row(8).number(7, 490.0).number(8, 26.0).number(9, 12740.0)
            sheet.row(9).text(7, "Grand Total").number(8, 56.0).number(9, 26240.0)

            FileOutputStream(file).use(workbook::write)
        }

        val sheet = InvoiceExcelReader.readAll(file).single()

        assertEquals("Counter W", sheet.branch)
        assertEquals(3, sheet.rows.size)
        assertEquals(450.0, sheet.rows.first().price)
        assertEquals(30.0, sheet.rows.first().qty)
        assertEquals(13500.0, sheet.rows.first().value)
        assertEquals(26240.0, sheet.rows.last().value)
    }

    @Test
    fun skipsSheetWhenOnlyMainTableExistsWithoutPivotSummary() {
        val file = createTempFile(prefix = "main-table-only-", suffix = ".xlsx").toFile()
        XSSFWorkbook().use { workbook ->
            val sheet = workbook.createSheet("MainOnly")
            sheet.row(0).text(2, "Invoice - ST00001")
            sheet.row(1).text(2, "Branch").text(3, "Counter Main")
            sheet.row(5)
                .text(0, "No.")
                .text(1, "Code / List")
                .text(2, "SKU")
                .text(3, "Price")
                .text(4, "Amount (Pieces)")
                .text(5, "มูลค่า")
            sheet.row(6)
                .number(0, 1.0)
                .text(2, "SKU-1")
                .number(3, 490.0)
                .number(4, 2.0)
                .number(5, 980.0)
            sheet.row(7)
                .text(3, "Total")
                .number(4, 2.0)
                .number(5, 980.0)

            FileOutputStream(file).use(workbook::write)
        }

        assertTrue(InvoiceExcelReader.readAll(file).isEmpty())
    }

    @Test
    fun readsThaiSummaryTableAndGpRow() {
        val file = createTempFile(prefix = "thai-summary-gp-", suffix = ".xlsx").toFile()
        XSSFWorkbook().use { workbook ->
            val sheet = workbook.createSheet("ST44763")
            sheet.row(0).text(2, "ใบส่งสินค้า - ST44763")
            sheet.row(1).text(2, "สาขา").text(3, "Depart Counter NaXWW เซ็นทรัลลำปาง")

            sheet.row(7).text(10, "ราคา").text(11, "จำนวน").text(12, "มูลค่า")
            sheet.row(8).number(10, 350.0).number(11, 3.0).number(12, 1050.0)
            sheet.row(9).number(10, 490.0).number(11, 16.0).number(12, 7840.0)
            sheet.row(10).number(10, 550.0).number(11, 3.0).number(12, 1650.0)
            sheet.row(11).text(10, "Grand Total").number(11, 22.0).number(12, 10540.0)
            sheet.row(12).text(10, "GP").text(11, "16%").number(12, 1686.4)

            FileOutputStream(file).use(workbook::write)
        }

        val sheet = InvoiceExcelReader.readAll(file).single()

        assertEquals("Depart Counter NaXWW เซ็นทรัลลำปาง", sheet.branch)
        assertEquals(5, sheet.rows.size)
        assertEquals(350.0, sheet.rows.first().price)
        assertEquals(3.0, sheet.rows.first().qty)
        assertEquals(1050.0, sheet.rows.first().value)
        assertTrue(sheet.rows[3].isTotal)
        assertEquals(22.0, sheet.rows[3].qty)
        assertEquals(10540.0, sheet.rows[3].value)
        assertTrue(sheet.rows[4].isGP)
        assertEquals(16.0, sheet.rows[4].gpPercent)
        assertEquals(1686.4, sheet.rows[4].value)
    }

    private fun org.apache.poi.ss.usermodel.Sheet.row(index: Int): org.apache.poi.ss.usermodel.Row =
        getRow(index) ?: createRow(index)

    private fun org.apache.poi.ss.usermodel.Row.text(column: Int, value: String): org.apache.poi.ss.usermodel.Row =
        apply { createCell(column).setCellValue(value) }

    private fun org.apache.poi.ss.usermodel.Row.number(column: Int, value: Double): org.apache.poi.ss.usermodel.Row =
        apply { createCell(column).setCellValue(value) }

    private fun org.apache.poi.ss.usermodel.Row.error(column: Int): org.apache.poi.ss.usermodel.Row =
        apply { createCell(column, CellType.ERROR) }
}
