import data.writer.AggregatedExcelWriter
import domain.model.PriceRow
import domain.model.SheetData
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import kotlin.io.path.createTempFile
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertFalse
import kotlin.test.assertTrue

class AggregatedExcelWriterTest {

    @Test
    fun writesSummaryAndDetailTabsWithFileAndSheetNamesButNoDateHeaders() {
        val outputFile = createTempFile(prefix = "aggregated-writer-", suffix = ".xlsx").toFile()
        AggregatedExcelWriter.write(sampleSheets(), outputFile)

        XSSFWorkbook(outputFile.inputStream()).use { workbook ->
            assertEquals(2, workbook.numberOfSheets)
            assertEquals("สรุปรวม (Summary)", workbook.getSheetAt(0).sheetName)
            assertEquals("รายละเอียดราคา (Price Detail)", workbook.getSheetAt(1).sheetName)

            val summary = workbook.getSheet("สรุปรวม (Summary)")
            val summaryHeaders = rowValues(summary.getRow(3), 10)
            assertTrue("File Name" in summaryHeaders)
            assertTrue("Sheet Name" in summaryHeaders)
            assertFalse(summaryHeaders.any { it.contains("Date", ignoreCase = true) || it.contains("วันที่") })

            val detail = workbook.getSheet("รายละเอียดราคา (Price Detail)")
            val detailHeaders = rowValues(detail.getRow(2), 8)
            assertTrue("File Name" in detailHeaders)
            assertTrue("Sheet Name" in detailHeaders)
            assertFalse(detailHeaders.any { it.contains("Date", ignoreCase = true) || it.contains("วันที่") })
            assertEquals("file-a.xlsx", detail.getRow(3).getCell(0).stringCellValue)
            assertEquals("file-a.xlsx", detail.getRow(4).getCell(0).stringCellValue)
            assertEquals("Branch A", detail.getRow(4).getCell(1).stringCellValue)
            assertEquals("file-a.xlsx", detail.getRow(6).getCell(0).stringCellValue)
            assertEquals("Branch A Round 2", detail.getRow(6).getCell(1).stringCellValue)
            assertEquals("file-b.xlsx", detail.getRow(8).getCell(0).stringCellValue)

            val firstFileFill = detail.getRow(4).getCell(0).cellStyle.fillForegroundColor
            assertEquals(firstFileFill, detail.getRow(6).getCell(0).cellStyle.fillForegroundColor)
            assertFalse(firstFileFill == detail.getRow(3).getCell(0).cellStyle.fillForegroundColor)
            assertTrue(detail.getRow(7).getCell(0).stringCellValue.contains("Branch A Round 2"))
            assertEquals("", detail.getRow(6).getCell(6).stringCellValue)
            assertEquals("16%", detail.getRow(7).getCell(6).stringCellValue)
        }
    }

    private fun sampleSheets(): List<SheetData> = listOf(
        SheetData(
            date = "15/03/2026",
            branch = "Counter A",
            rows = listOf(
                PriceRow(price = 190.0, qty = 2.0, value = 380.0),
                PriceRow(price = null, qty = 2.0, value = 380.0, isTotal = true),
                PriceRow(price = null, qty = null, value = 60.8, isGP = true, gpPercent = 0.16)
            ),
            sourceFileName = "file-a.xlsx",
            sourceSheetName = "Branch A"
        ),
        SheetData(
            date = "",
            branch = "Counter A2",
            rows = listOf(
                PriceRow(price = 350.0, qty = 1.0, value = 350.0),
                PriceRow(price = null, qty = 1.0, value = 350.0, isTotal = true),
                PriceRow(price = null, qty = null, value = 56.0, isGP = true, gpPercent = 0.16)
            ),
            sourceFileName = "file-a.xlsx",
            sourceSheetName = "Branch A Round 2"
        ),
        SheetData(
            date = "",
            branch = "Counter B",
            rows = listOf(
                PriceRow(price = 350.0, qty = 3.0, value = 1050.0),
                PriceRow(price = null, qty = 3.0, value = 1050.0, isTotal = true),
                PriceRow(price = null, qty = null, value = 210.0, isGP = true, gpPercent = 0.20)
            ),
            sourceFileName = "file-b.xlsx",
            sourceSheetName = "Branch B"
        )
    )

    private fun rowValues(row: org.apache.poi.ss.usermodel.Row, count: Int): List<String> =
        (0 until count).map { row.getCell(it)?.stringCellValue.orEmpty() }
}
