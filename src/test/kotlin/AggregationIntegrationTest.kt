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
        assertTrue(parsedSheets.isNotEmpty(), "expected at least one visible sheet to be parsed")

        val outputFile = createTempFile(prefix = "aggregation-integration-", suffix = ".xlsx").toFile()
        AggregatedExcelWriter.write(parsedSheets, outputFile)

        assertTrue(outputFile.exists(), "expected output workbook to be created")

        XSSFWorkbook(outputFile.inputStream()).use { workbook ->
            assertEquals(1, workbook.numberOfSheets)
            assertEquals("Accountant View", workbook.getSheetAt(0).sheetName)

            val accountantView = workbook.getSheet("Accountant View")
            assertTrue(accountantView.lastRowNum > 5, "accountant view should contain summary and grouped sections")
            assertEquals("Combined Invoice Summary", accountantView.getRow(0).getCell(0).stringCellValue)
        }
    }
}
