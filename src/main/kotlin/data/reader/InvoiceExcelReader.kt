package data.reader

import data.config.ConfigLoader
import domain.model.AggregatorConfig
import domain.model.PriceRow
import domain.model.SheetData
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream

/** Reads and parses invoice Excel files into domain model objects. */
object InvoiceExcelReader {

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    fun readAll(file: File, config: AggregatorConfig = ConfigLoader.load()): List<SheetData> =
        runCatching {
            FileInputStream(file).use { fis ->
                WorkbookFactory.create(fis).use { wb ->
                    (0 until wb.numberOfSheets).mapNotNull { index ->
                        parseSheet(wb.getSheetAt(index), index, file.name, config)
                    }
                }
            }
        }.getOrDefault(emptyList())

    fun diagnose(file: File, config: AggregatorConfig = ConfigLoader.load()): String =
        runCatching {
            FileInputStream(file).use { fis ->
                WorkbookFactory.create(fis).use { wb ->
                    buildDiagnosticReport(wb.getSheetAt(0), file, config)
                }
            }
        }.getOrElse { e -> "Failed to open/read file: ${e.message}" }

    // -------------------------------------------------------------------------
    // Sheet parsing
    // -------------------------------------------------------------------------

    private fun parseSheet(sheet: Sheet, index: Int, fileName: String, config: AggregatorConfig): SheetData? {
        val (date, branch) = parseDateAndBranch(sheet, config)
        val (headerRowIndex, columns) = detectHeaderAndColumns(sheet, config) ?: return null
        val rows = parseDataRows(sheet, headerRowIndex, columns, config)
        return rows.takeIf { it.isNotEmpty() }
            ?.let { SheetData(date = date, branch = branch, rows = it, sourceFileName = fileName, sourceSheetIndex = index, sourceSheetName = sheet.sheetName) }
    }

    // -------------------------------------------------------------------------
    // Header detection
    // -------------------------------------------------------------------------

    private data class ColumnMapping(val price: Int, val qty: Int, val value: Int)
    private data class HeaderScanResult(val found: Boolean, val price: Int, val qty: Int, val value: Int)

    private fun detectHeaderAndColumns(sheet: Sheet, config: AggregatorConfig): Pair<Int, ColumnMapping>? {
        val labels = HeaderLabels(config)
        for (row in sheet) {
            val scan = scanRowForHeader(row, labels)
            if (!scan.found) continue

            val initial = ColumnMapping(
                price = scan.price.takeIf { it >= 0 } ?: config.priceColumn,
                qty   = scan.qty.takeIf   { it >= 0 } ?: config.qtyColumn,
                value = scan.value.takeIf { it >= 0 } ?: config.valueColumn
            )
            val sampleRange = (row.rowNum + 1)..minOf(sheet.lastRowNum, row.rowNum + 21)
            val final = if (!sampleRange.isEmpty()) adjustColumnMapping(sheet, sampleRange, initial) else initial
            return row.rowNum to final
        }
        return null
    }

    private data class HeaderLabels(val price: String, val value: String, val total: String, val qty: String) {
        constructor(config: AggregatorConfig) : this(
            price = normalize(config.labelPrice).lowercase(),
            value = normalize(config.labelValue).lowercase(),
            total = normalize(config.labelTotal).lowercase(),
            qty   = normalize(config.labelQty).lowercase()
        )
    }

    private fun scanRowForHeader(row: Row, labels: HeaderLabels): HeaderScanResult {
        val headerMarkers = setOf(labels.price, labels.value, labels.total)
        val colByLabel = row
            .associateBy { cleanedCellText(it) }
            .mapValues { (_, cell) -> cell.columnIndex }

        val found = colByLabel.keys.any { it.isNotEmpty() && it in headerMarkers }
        return HeaderScanResult(
            found  = found,
            price  = colByLabel[labels.price] ?: -1,
            qty    = colByLabel[labels.qty]   ?: -1,
            value  = colByLabel[labels.value] ?: -1
        )
    }

    // -------------------------------------------------------------------------
    // Column mapping adjustment (heuristic)
    // -------------------------------------------------------------------------

    private fun adjustColumnMapping(sheet: Sheet, sampleRange: IntRange, initial: ColumnMapping): ColumnMapping {
        val maxCols  = maxOf(sheet.getRow(sampleRange.first)?.lastCellNum?.toInt() ?: 0, 10)
        val qtyStats = sampleStats(sheet, initial.qty, sampleRange)
        val skuCol   = detectSkuColumn(sheet, sampleRange, maxCols)
        return when {
            skuCol >= 0              -> adjustWithSkuColumn(sheet, sampleRange, initial, skuCol, qtyStats.numericCount, maxCols)
            qtyStats.numericCount == 0 -> adjustWithFallbackQty(sheet, sampleRange, initial, maxCols)
            else                     -> initial
        }
    }

    private fun adjustWithSkuColumn(
        sheet: Sheet, sampleRange: IntRange, initial: ColumnMapping,
        skuCol: Int, qtyNumCount: Int, maxCols: Int
    ): ColumnMapping {
        val skuStats  = sampleStats(sheet, skuCol, sampleRange)
        val minUseful = (sampleRange.last - sampleRange.first + 1) / 4
        if (qtyNumCount >= 2 || skuStats.numericCount < minUseful) return initial

        val byAvgDescending = (0 until maxCols)
            .map { c -> c to sampleStats(sheet, c, sampleRange).average }
            .sortedByDescending { it.second }
            .map { it.first }
            .filter { it != skuCol }

        return ColumnMapping(
            price = byAvgDescending.getOrElse(1) { initial.price },
            qty   = skuCol,
            value = byAvgDescending.getOrElse(0) { initial.value }
        )
    }

    private fun adjustWithFallbackQty(
        sheet: Sheet, sampleRange: IntRange, initial: ColumnMapping, maxCols: Int
    ): ColumnMapping {
        val numericCols = (0 until maxCols).filter { c -> sampleStats(sheet, c, sampleRange).numericCount > 0 }
        val fallbackQty = numericCols.firstOrNull { it != initial.price } ?: numericCols.firstOrNull() ?: initial.qty
        return initial.copy(qty = fallbackQty)
    }

    // -------------------------------------------------------------------------
    // Data row parsing
    // -------------------------------------------------------------------------

    private fun parseDataRows(sheet: Sheet, headerRowIndex: Int, cols: ColumnMapping, config: AggregatorConfig): List<PriceRow> {
        val rows = mutableListOf<PriceRow>()
        var hasData = false

        for (rowIndex in (headerRowIndex + 1)..sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex) ?: continue
            val rawPrice = CellUtils.getCellString(row.getCell(cols.price))
            val rawQty   = CellUtils.getCellString(row.getCell(cols.qty))
            val rawValue = CellUtils.getCellString(row.getCell(cols.value))

            val entry = classifyAndBuildRow(rawPrice, rawQty, rawValue, config)
            if (entry == null) {
                if (hasData && rawPrice.isEmpty() && rawQty.isEmpty() && rawValue.isEmpty()) return rows
                continue
            }

            rows.add(entry)
            if (!entry.isGP) hasData = true
        }

        return rows
    }

    private fun classifyAndBuildRow(rawPrice: String, rawQty: String, rawValue: String, config: AggregatorConfig): PriceRow? {
        val numericPrice = normalizeNumber(rawPrice)
        val numericQty   = normalizeNumber(rawQty)
        val numericValue = normalizeNumber(rawValue)

        return when {
            isTotalRow(rawPrice, config) -> PriceRow(
                price = null, qty = numericQty.toDoubleOrNull(), value = numericValue.toDoubleOrNull(), isTotal = true
            )
            isGpRow(rawPrice, rawQty, config) -> PriceRow(
                price = null, qty = null, value = numericValue.toDoubleOrNull(), isGP = true, gpPercent = CellUtils.parsePercent(rawQty)
            )
            numericPrice.toDoubleOrNull() != null -> PriceRow(
                price = numericPrice.toDoubleOrNull(), qty = numericQty.toDoubleOrNull(), value = numericValue.toDoubleOrNull()
            )
            else -> null
        }
    }

    // -------------------------------------------------------------------------
    // Date / branch detection
    // -------------------------------------------------------------------------

    private fun parseDateAndBranch(sheet: Sheet, config: AggregatorConfig): Pair<String, String> {
        var date   = ""
        var branch = ""
        for (row in sheet) {
            for (cell in row) {
                val text = normalize(CellUtils.getCellString(cell))
                if (date.isEmpty()   && config.dateKeywords.any   { text.contains(normalize(it)) }) date   = nextCellText(row, cell)
                if (branch.isEmpty() && config.branchKeywords.any { text.contains(normalize(it)) }) branch = nextCellText(row, cell)
            }
            if (date.isNotEmpty() && branch.isNotEmpty()) break
        }
        return date to branch
    }

    private fun nextCellText(row: Row, cell: Cell): String =
        normalize(CellUtils.getCellString(row.getCell(cell.columnIndex + 1)))

    // -------------------------------------------------------------------------
    // Row classification predicates
    // -------------------------------------------------------------------------

    private fun isTotalRow(rawPrice: String, config: AggregatorConfig): Boolean {
        val normalised = normalize(rawPrice)
        return config.totalKeywords.any { kw ->
            normalised.equals(normalize(kw), ignoreCase = true) || normalised.contains(normalize(kw))
        }
    }

    private fun isGpRow(rawPrice: String, rawQty: String, config: AggregatorConfig): Boolean =
        normalize(rawPrice).equals(normalize(config.gpKeyword), ignoreCase = true) && rawQty.isNotEmpty()

    // -------------------------------------------------------------------------
    // Sample statistics & SKU detection
    // -------------------------------------------------------------------------

    private data class ColStats(val numericCount: Int, val nonEmptyCount: Int, val average: Double)

    private fun sampleStats(sheet: Sheet, col: Int, range: IntRange): ColStats {
        var numericCount = 0; var nonEmptyCount = 0; var sum = 0.0
        for (r in range) {
            val raw = CellUtils.getCellString(sheet.getRow(r)?.getCell(col))
            if (raw.isNotEmpty()) nonEmptyCount++
            normalizeNumber(raw).toDoubleOrNull()?.let { v -> numericCount++; sum += v }
        }
        return ColStats(numericCount, nonEmptyCount, if (numericCount > 0) sum / numericCount else 0.0)
    }

    private fun detectSkuColumn(sheet: Sheet, range: IntRange, maxCols: Int): Int {
        fun isSkuLike(text: String) = text.length >= 4 && text.any { it.isLetter() } && text.any { it.isDigit() } && text.contains('-')
        val best = (0 until maxCols).maxByOrNull { col ->
            range.count { r -> isSkuLike(normalize(CellUtils.getCellString(sheet.getRow(r)?.getCell(col)))) }
        }
        return best.takeIf { it != null && range.count { r -> isSkuLike(normalize(CellUtils.getCellString(sheet.getRow(r)?.getCell(it)))) } > 0 } ?: -1
    }

    // -------------------------------------------------------------------------
    // Diagnostics
    // -------------------------------------------------------------------------

    private fun buildDiagnosticReport(sheet: Sheet, file: File, config: AggregatorConfig): String {
        val (date, branch) = parseDateAndBranch(sheet, config)
        val header = detectHeaderAndColumns(sheet, config)
        return buildString {
            appendLine("File: ${file.absolutePath}")
            appendLine("Date detected: ${date.ifEmpty { "(none)" }}")
            appendLine("Branch detected: ${branch.ifEmpty { "(none)" }}")
            if (header == null) {
                appendLine("Header row: NOT FOUND")
                appendLine("First 10 rows (normalized):")
                sheet.take(10).forEach { row -> appendLine("row ${row.rowNum}: ${rowTexts(row).joinToString(" | ")}") }
            } else {
                val (hdrIdx, cols) = header
                appendLine("Header row: $hdrIdx")
                appendLine("Detected cols -> price:${cols.price}, qty:${cols.qty}, value:${cols.value}")
                sheet.getRow(hdrIdx)?.let { appendLine(rowTexts(it).joinToString(" | ")) }
                appendLine("Next 5 data rows (raw):")
                (hdrIdx + 1..sheet.lastRowNum).asSequence().mapNotNull { sheet.getRow(it) }.take(5).forEach { row ->
                    val p = CellUtils.getCellString(row.getCell(cols.price))
                    val q = CellUtils.getCellString(row.getCell(cols.qty))
                    val v = CellUtils.getCellString(row.getCell(cols.value))
                    appendLine("row ${row.rowNum} -> price:'$p' qty:'$q' value:'$v'")
                }
            }
        }
    }

    private fun rowTexts(row: Row): List<String> = row.map { normalize(CellUtils.getCellString(it)) }

    // -------------------------------------------------------------------------
    // Text normalisation utilities
    // -------------------------------------------------------------------------

    private fun cleanedCellText(cell: Cell): String =
        listOf("sum of", "subtotal").fold(normalize(CellUtils.getCellString(cell)).lowercase()) { acc, t -> acc.replace(t, "") }.trim()

    private fun normalize(s: String?): String =
        s?.trim()?.replace("\u00A0", " ")?.replace(Regex("\\s+"), " ") ?: ""

    private fun normalizeNumber(s: String?): String =
        s?.trim()?.replace("\u00A0", " ")?.replace(Regex("[,\\s]"), "")?.replace("\u200B", "") ?: ""
}



