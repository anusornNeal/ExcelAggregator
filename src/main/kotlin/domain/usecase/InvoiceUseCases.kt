package domain.usecase

import data.reader.InvoiceExcelReader
import domain.model.SheetData
import java.io.File
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter

/** Port: anything that can write aggregated sheets to a file. Implemented in data layer. */
fun interface InvoiceWriter {
    fun write(sheets: List<SheetData>, outputFile: File)
}

/** Port: anything that can read invoice sheets from a file. Implemented in data layer. */
interface InvoiceReader {
    fun readAll(file: File): List<SheetData>
    fun diagnose(file: File): String
}

/** Orchestrates reading, aggregating, writing, and cleanup for invoice Excel files. */
class AggregateInvoicesUseCase(
    private val reader: InvoiceReader = object : InvoiceReader {
        override fun readAll(file: File): List<SheetData> = InvoiceExcelReader.readAll(file)
        override fun diagnose(file: File): String = InvoiceExcelReader.diagnose(file)
    },
    private val writer: InvoiceWriter = InvoiceWriter { sheets, file -> data.writer.AggregatedExcelWriter.write(sheets, file) }
) {
    data class Result(
        val processedCount: Int,
        val skipped: List<Pair<File, String>>,
        val outFile: File?,
        val deletedCount: Int = 0,
        val deleteFailures: List<Pair<File, String>> = emptyList()
    )

    fun execute(selectedFiles: List<File>): Result {
        val (sheets, processedFiles, skipped) = parseFiles(selectedFiles)

        if (sheets.isEmpty()) return Result(processedFiles.size, skipped, null)

        val outFile = resolveOutputFile(selectedFiles)
        writer.write(sheets, outFile)
        val deleteFailures = deleteProcessedFiles(processedFiles, outFile)

        return Result(
            processedCount = processedFiles.size,
            skipped = skipped,
            outFile = outFile,
            deletedCount = processedFiles.size - deleteFailures.size,
            deleteFailures = deleteFailures
        )
    }

    private fun parseFiles(files: List<File>): Triple<List<SheetData>, List<File>, List<Pair<File, String>>> {
        val sheets = mutableListOf<SheetData>()
        val processed = mutableListOf<File>()
        val skipped = mutableListOf<Pair<File, String>>()

        for (file in files) {
            runCatching { reader.readAll(file) }
                .onSuccess { fileSheets ->
                    if (fileSheets.isNotEmpty()) {
                        sheets.addAll(fileSheets)
                        processed.add(file)
                    } else {
                        skipped.add(file to runCatching { reader.diagnose(file) }.getOrElse { "diagnose failed: ${it.message}" })
                    }
                }
                .onFailure { skipped.add(file to "อ่านไฟล์ไม่สำเร็จ: ${it.message}") }
        }

        return Triple(sheets, processed, skipped)
    }

    private fun deleteProcessedFiles(files: List<File>, outFile: File): List<Pair<File, String>> =
        files
            .filterNot { it.absoluteFile == outFile.absoluteFile }
            .mapNotNull { file ->
                runCatching {
                    if (!file.exists() || file.delete()) null else file to "delete returned false"
                }.getOrElse { file to (it.message ?: it::class.simpleName.orEmpty()) }
            }

    private fun resolveOutputFile(files: List<File>): File {
        val baseDir = files.firstOrNull()?.parentFile ?: File(".")
        val time = ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"))
        return File(baseDir, "aggregated_invoices_$time.xlsx")
    }
}

/** Generates a diagnostic report for a single file and writes it to disk. */
class DiagnoseFileUseCase(
    private val reader: InvoiceExcelReader = InvoiceExcelReader
) {
    data class Result(val reportPath: String, val content: String)

    fun execute(file: File): Result {
        val content = runCatching { reader.diagnose(file) }.getOrElse { "Failed to diagnose: ${it.message}" }
        val time = ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"))
        val report = File(File(".").absoluteFile, "diagnose_${file.name}_$time.txt")
        report.writeText(content)
        return Result(report.absolutePath, content)
    }
}

/** Builds the human-readable status message shown in the UI after processing. */
class BuildStatusMessageUseCase {
    fun execute(result: AggregateInvoicesUseCase.Result): String {
        val outFile = result.outFile ?: return buildNoDataMessage(result.skipped)
        return buildString {
            appendLine("✅ รวม ${result.processedCount} ไฟล์สำเร็จ → ${outFile.name}")
            appendLine("ลบไฟล์ต้นฉบับ ${result.deletedCount} ไฟล์แล้ว")
            if (result.deleteFailures.isNotEmpty()) {
                appendLine("ลบไฟล์ต้นฉบับไม่สำเร็จ ${result.deleteFailures.size} ไฟล์:")
                result.deleteFailures.take(10).forEach { appendLine(" - ${it.first.name}: ${it.second}") }
                if (result.deleteFailures.size > 10) appendLine(" (แสดง 10 รายการแรก)")
            }
            if (result.skipped.isNotEmpty()) {
                appendLine("ข้ามไฟล์ที่ไม่มีข้อมูล ${result.skipped.size} ไฟล์:")
                result.skipped.take(10).forEach { appendLine(" - ${it.first.name}: ${it.second}") }
                if (result.skipped.size > 10) appendLine(" (แสดง 10 รายการแรก)")
            }
        }
    }

    private fun buildNoDataMessage(skipped: List<Pair<File, String>>): String =
        "⚠️ ไม่พบข้อมูลที่ต้องการในไฟล์ที่เลือก\n${skipped.joinToString("; ") { "${it.first.name}: ${it.second}" }}"
}
