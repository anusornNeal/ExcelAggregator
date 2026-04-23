import domain.model.PriceRow
import domain.model.SheetData
import domain.usecase.AggregateInvoicesUseCase
import domain.usecase.BuildStatusMessageUseCase
import domain.usecase.InvoiceReader
import domain.usecase.InvoiceWriter
import java.io.File
import kotlin.io.path.createTempDirectory
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertFalse
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

class AggregateInvoicesUseCaseTest {

    @Test
    fun deletesOnlyFilesThatWereAggregated() {
        val dir = createTempDirectory(prefix = "aggregate-cleanup-").toFile()
        val processedFile = File(dir, "processed.xlsx").apply { writeText("processed") }
        val skippedFile = File(dir, "skipped.xlsx").apply { writeText("skipped") }

        val useCase = AggregateInvoicesUseCase(
            reader = object : InvoiceReader {
                override fun readAll(file: File): List<SheetData> =
                    if (file == processedFile) listOf(sampleSheet(file.name)) else emptyList()

                override fun diagnose(file: File): String = "no invoice data"
            },
            writer = InvoiceWriter { _, outputFile -> outputFile.writeText("aggregated") }
        )

        val result = useCase.execute(listOf(processedFile, skippedFile))

        assertNotNull(result.outFile)
        assertEquals(1, result.processedCount)
        assertEquals(1, result.deletedCount)
        assertTrue(result.deleteFailures.isEmpty())
        assertEquals(listOf(skippedFile), result.skipped.map { it.first })
        assertFalse(processedFile.exists())
        assertTrue(skippedFile.exists())
        assertTrue(result.outFile.exists())
    }

    @Test
    fun statusMessageDoesNotGenerateSkippedReport() {
        val dir = createTempDirectory(prefix = "aggregate-status-").toFile()
        val outFile = File(dir, "aggregated.xlsx").apply { writeText("aggregated") }
        val skippedFile = File(dir, "skipped.xlsx").apply { writeText("skipped") }
        val result = AggregateInvoicesUseCase.Result(
            processedCount = 1,
            skipped = listOf(skippedFile to "no invoice data"),
            outFile = outFile,
            deletedCount = 1
        )

        val message = BuildStatusMessageUseCase().execute(result)

        assertTrue(message.contains("skipped.xlsx"))
        assertTrue(dir.listFiles { file -> file.name.startsWith("skipped_report_") }.isNullOrEmpty())
    }

    @Test
    fun doesNotDeleteProcessedFilesWhenWriterFails() {
        val dir = createTempDirectory(prefix = "aggregate-writer-fails-").toFile()
        val processedFile = File(dir, "processed.xlsx").apply { writeText("processed") }
        val useCase = AggregateInvoicesUseCase(
            reader = object : InvoiceReader {
                override fun readAll(file: File): List<SheetData> = listOf(sampleSheet(file.name))
                override fun diagnose(file: File): String = "not used"
            },
            writer = InvoiceWriter { _, _ -> error("write failed") }
        )

        val failed = runCatching { useCase.execute(listOf(processedFile)) }

        assertTrue(failed.isFailure)
        assertTrue(processedFile.exists())
    }

    private fun sampleSheet(fileName: String): SheetData =
        SheetData(
            date = "",
            branch = "Branch",
            rows = listOf(
                PriceRow(price = 190.0, qty = 2.0, value = 380.0),
                PriceRow(price = null, qty = 2.0, value = 380.0, isTotal = true)
            ),
            sourceFileName = fileName,
            sourceSheetName = "Sheet1"
        )
}
