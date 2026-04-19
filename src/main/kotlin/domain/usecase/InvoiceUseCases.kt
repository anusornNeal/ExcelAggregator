package domain.usecase

import data.config.ConfigLoader
import data.reader.InvoiceExcelReader
import domain.model.SheetData
import java.io.File
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter

/** Port: anything that can write aggregated sheets to a file. Implemented in data layer. */
fun interface InvoiceWriter {
    fun write(sheets: List<SheetData>, outputFile: File)
}

/** Orchestrates reading, aggregating, and writing invoice Excel files. */
class AggregateInvoicesUseCase(
    private val reader: InvoiceExcelReader = InvoiceExcelReader,
    private val writer: InvoiceWriter = InvoiceWriter { sheets, file -> data.writer.AggregatedExcelWriter.write(sheets, file) },
    private val configLoader: ConfigLoader = ConfigLoader
) {
    data class Result(
        val processedCount: Int,
        val skipped: List<Pair<File, String>>,
        val outFile: File?
    )

    fun execute(selectedFiles: List<File>): Result {
        val (sheets, processedFiles, skipped) = parseFiles(selectedFiles)

        if (sheets.isEmpty()) return Result(processedFiles.size, skipped, null)

        val outFile = resolveOutputFile(selectedFiles)
        writer.write(sheets, outFile)

        return Result(processedFiles.size, skipped, outFile)
    }

    private fun parseFiles(files: List<File>): Triple<List<SheetData>, List<File>, List<Pair<File, String>>> {
        val sheets = mutableListOf<SheetData>()
        val processed = mutableListOf<File>()
        val skipped = mutableListOf<Pair<File, String>>()

        for (file in files) {
            runCatching { reader.readAll(file) }
                .onSuccess { fileSheets ->
                    if (fileSheets.isNotEmpty()) { sheets.addAll(fileSheets); processed.add(file) }
                    else skipped.add(file to runCatching { reader.diagnose(file) }.getOrElse { "diagnose failed: ${it.message}" })
                }
                .onFailure { skipped.add(file to "อ่านไฟล์ไม่สำเร็จ: ${it.message}") }
        }

        return Triple(sheets, processed, skipped)
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
        val time    = ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"))
        val report  = File(File(".").absoluteFile, "diagnose_${file.name}_$time.txt")
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
            appendLine("ลบไฟล์ต้นฉบับ ${result.processedCount} ไฟล์แล้ว")
            if (result.skipped.isNotEmpty()) {
                appendLine("ข้ามไฟล์ที่ไม่มีข้อมูล ${result.skipped.size} ไฟล์:")
                result.skipped.take(10).forEach { appendLine(" - ${it.first.name}: ${it.second}") }
                if (result.skipped.size > 10) appendLine(" (แสดง 10 รายการแรก)")
                append(writeSkippedReport(outFile, result.skipped))
            }
        }
    }

    private fun buildNoDataMessage(skipped: List<Pair<File, String>>): String =
        "⚠️ ไม่พบข้อมูลที่ต้องการในไฟล์ที่เลือก\n${skipped.joinToString("; ") { "${it.first.name}: ${it.second}" }}"

    private fun writeSkippedReport(outFile: File, skipped: List<Pair<File, String>>): String =
        runCatching {
            val time   = ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"))
            val report = File(outFile.parentFile ?: File("."), "skipped_report_$time.txt")
            report.writeText(buildString {
                appendLine("Skipped files report"); appendLine("Generated: ${ZonedDateTime.now()}")
                appendLine("Output: ${outFile.absolutePath}\n")
                skipped.forEach { appendLine("${it.first.absolutePath} -> ${it.second}") }
            })
            "รายงานไฟล์ที่ข้ามบันทึกที่: ${report.absolutePath}\n"
        }.getOrElse { "ไม่สามารถบันทึกรายงานไฟล์ที่ข้าม: ${it.message}\n" }
}
