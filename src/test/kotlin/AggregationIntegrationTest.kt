import data.reader.InvoiceExcelReader
import data.writer.AggregatedExcelWriter
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Assumptions.assumeTrue
import kotlin.io.path.createTempFile
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class AggregationIntegrationTest {

    @Test
    fun aggregatesExperimentalExamplesIntoAccountantWorkbook() {
        val experimentalDir = java.io.File("C:/Users/tatar/Downloads/Bo's payslip/Aggregator/files/experimental")
        assumeTrue(experimentalDir.exists(), "experimental example folder is not available on this machine")

        val inputFiles = experimentalDir
            .listFiles { file -> file.isFile && file.extension.lowercase() == "xlsx" }
            ?.sortedBy { it.name }
            .orEmpty()

        assumeTrue(inputFiles.isNotEmpty(), "no example workbooks found in experimental folder")

        val parsedSheets = inputFiles.flatMap { InvoiceExcelReader.readAll(it) }
        assumeTrue(parsedSheets.isNotEmpty(), "no example workbook contains a supported summary table")

        val outputFile = createTempFile(prefix = "aggregation-integration-", suffix = ".xlsx").toFile()
        AggregatedExcelWriter.write(parsedSheets, outputFile)

        assertTrue(outputFile.exists(), "expected output workbook to be created")

        XSSFWorkbook(outputFile.inputStream()).use { workbook ->
            assertEquals(2, workbook.numberOfSheets)
            assertEquals("สรุปรวม (Summary)", workbook.getSheetAt(0).sheetName)
            assertEquals("รายละเอียดราคา (Price Detail)", workbook.getSheetAt(1).sheetName)

            val summary = workbook.getSheet("สรุปรวม (Summary)")
            assertTrue(summary.lastRowNum >= 5, "summary should contain one row per parsed sheet")
            assertEquals("สรุปรวมใบส่งสินค้า Depart Counter - ทุกสาขา", summary.getRow(0).getCell(0).stringCellValue)
        }
    }
}
