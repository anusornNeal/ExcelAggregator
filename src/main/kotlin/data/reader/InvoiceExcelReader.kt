package data.reader

import data.config.ConfigLoader
import domain.model.AggregatorConfig
import domain.model.PriceRow
import domain.model.SheetData
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.text.SimpleDateFormat
import kotlin.math.abs

/** Reads and parses invoice Excel files into domain model objects. */
object InvoiceExcelReader {
    private val whitespaceRegex = Regex("\\s+")
    private val numericCleanupRegex = Regex("[,\\s]")
    private val headerNoiseTokens = arrayOf("sum of", "subtotal")
    private val outputDateFormat = SimpleDateFormat("d/M/yyyy")

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    fun readAll(file: File, config: AggregatorConfig = ConfigLoader.load()): List<SheetData> =
        runCatching {
            val readerConfig = ReaderConfig(config)
            FileInputStream(file).use { fis ->
                WorkbookFactory.create(fis).use { wb ->
                    (0 until wb.numberOfSheets).mapNotNull { index ->
                        if (wb.isSheetHidden(index) || wb.isSheetVeryHidden(index)) {
                            null
                        } else {
                            parseSheet(wb.getSheetAt(index), index, file.name, readerConfig)
                        }
                    }
                }
            }
        }.getOrDefault(emptyList())

    fun diagnose(file: File, config: AggregatorConfig = ConfigLoader.load()): String =
        runCatching {
            val readerConfig = ReaderConfig(config)
            FileInputStream(file).use { fis ->
                WorkbookFactory.create(fis).use { wb ->
                    val visibleSheet = (0 until wb.numberOfSheets)
                        .firstOrNull { !wb.isSheetHidden(it) && !wb.isSheetVeryHidden(it) }
                        ?.let(wb::getSheetAt)
                        ?: return@use "No visible sheets found"
                    buildDiagnosticReport(visibleSheet, file, readerConfig)
                }
            }
        }.getOrElse { e -> "Failed to open/read file: ${e.message}" }

    // -------------------------------------------------------------------------
    // Sheet parsing
    // -------------------------------------------------------------------------

    private fun parseSheet(sheet: Sheet, index: Int, fileName: String, config: ReaderConfig): SheetData? {
        if (!isValidInvoiceSheet(sheet, config)) return null
        val (date, branch) = parseDateAndBranch(sheet, config)
        val (headerRowIndex, columns) = detectHeaderAndColumns(sheet, config) ?: return null
        if (date.isBlank() || branch.isBlank()) return null
        val rows = parseDataRows(sheet, headerRowIndex, columns, config)
        return rows.takeIf { it.isNotEmpty() }
            ?.let { SheetData(date = date, branch = branch, rows = it, sourceFileName = fileName, sourceSheetIndex = index, sourceSheetName = sheet.sheetName) }
    }

    // -------------------------------------------------------------------------
    // Header detection
    // -------------------------------------------------------------------------

    private data class ReaderConfig(
        val documentKeywords: List<String>,
        val priceColumn: Int,
        val labelPrice: String,
        val labelValue: String,
        val labelTotal: String,
        val labelQty: String,
        val dateKeywords: List<String>,
        val branchKeywords: List<String>,
        val totalKeywords: List<String>,
        val gpKeyword: String
    ) {
        constructor(config: AggregatorConfig) : this(
            documentKeywords = config.documentKeywords.map { normalize(it).lowercase() },
            priceColumn = config.priceColumn,
            labelPrice = normalize(config.labelPrice).lowercase(),
            labelValue = normalize(config.labelValue).lowercase(),
            labelTotal = normalize(config.labelTotal).lowercase(),
            labelQty = normalize(config.labelQty).lowercase(),
            dateKeywords = config.dateKeywords.map { normalize(it) },
            branchKeywords = config.branchKeywords.map { normalize(it) },
            totalKeywords = config.totalKeywords.map { normalize(it) },
            gpKeyword = normalize(config.gpKeyword)
        )
    }

    private data class ColumnMapping(val price: Int, val qty: Int, val value: Int)
    private data class HeaderScanResult(val found: Boolean, val price: Int, val qty: Int, val value: Int)

    private fun detectHeaderAndColumns(sheet: Sheet, config: ReaderConfig): Pair<Int, ColumnMapping>? {
        val labels = HeaderLabels(config)
        for (row in sheet) {
            val scan = scanRowForHeader(row, labels)
            if (!scan.found) continue

            val initial = buildInitialMapping(scan, config)
            val sampleRange = sampleRangeAfter(row.rowNum, sheet.lastRowNum)
            val resolved = sampleRange?.let { adjustColumnMapping(sheet, it, initial) } ?: initial
            return row.rowNum to resolved
        }
        return null
    }

    private fun buildInitialMapping(scan: HeaderScanResult, config: ReaderConfig): ColumnMapping {
        val anchor = resolveAnchor(scan, config)
        return ColumnMapping(
            price = scan.price.takeIf { it >= 0 } ?: anchor,
            qty = scan.qty.takeIf { it >= 0 } ?: (anchor + 1),
            value = scan.value.takeIf { it >= 0 } ?: (anchor + 2)
        )
    }

    private fun resolveAnchor(scan: HeaderScanResult, config: ReaderConfig): Int =
        when {
            scan.price >= 0 -> scan.price
            scan.qty >= 0 -> scan.qty - 1
            scan.value >= 0 -> scan.value - 2
            else -> config.priceColumn
        }.coerceAtLeast(0)

    private fun sampleRangeAfter(rowIndex: Int, lastRowIndex: Int): IntRange? {
        val end = minOf(lastRowIndex, rowIndex + 21)
        return if (end > rowIndex) (rowIndex + 1)..end else null
    }

    private data class HeaderLabels(val price: String, val value: String, val total: String, val qty: String) {
        constructor(config: ReaderConfig) : this(config.labelPrice, config.labelValue, config.labelTotal, config.labelQty)
    }

    private fun scanRowForHeader(row: Row, labels: HeaderLabels): HeaderScanResult {
        var found = false
        var price = -1
        var qty = -1
        var value = -1

        for (cell in row) {
            when (cleanedCellText(cell)) {
                labels.price -> {
                    found = true
                    price = cell.columnIndex
                }
                labels.value, labels.total -> {
                    found = true
                    value = cell.columnIndex
                }
                labels.qty -> qty = cell.columnIndex
            }
        }

        return HeaderScanResult(found = found, price = price, qty = qty, value = value)
    }

    // -------------------------------------------------------------------------
    // Column mapping adjustment (heuristic)
    // -------------------------------------------------------------------------

    private fun adjustColumnMapping(sheet: Sheet, sampleRange: IntRange, initial: ColumnMapping): ColumnMapping {
        val maxCols  = maxOf(sheet.getRow(sampleRange.first)?.lastCellNum?.toInt() ?: 0, 10)
        val statsCache = HashMap<Int, ColStats>(maxCols)
        fun statsFor(col: Int): ColStats = statsCache.getOrPut(col) { sampleStats(sheet, col, sampleRange) }

        val qtyStats = statsFor(initial.qty)
        val skuCol   = detectSkuColumn(sheet, sampleRange, maxCols)
        return when {
            skuCol >= 0                -> adjustWithSkuColumn(initial, skuCol, qtyStats.numericCount, maxCols, sampleRange) { col -> statsFor(col) }
            qtyStats.numericCount == 0 -> adjustWithFallbackQty(initial, maxCols) { col -> statsFor(col) }
            else                     -> initial
        }
    }

    private fun adjustWithSkuColumn(
        initial: ColumnMapping,
        skuCol: Int,
        qtyNumCount: Int,
        maxCols: Int,
        sampleRange: IntRange,
        statsFor: (Int) -> ColStats
    ): ColumnMapping {
        val skuStats  = statsFor(skuCol)
        val minUseful = (sampleRange.last - sampleRange.first + 1) / 4
        if (qtyNumCount >= 2 || skuStats.numericCount < minUseful) return initial

        var bestCol = initial.value
        var secondBestCol = initial.price
        var bestAvg = Double.NEGATIVE_INFINITY
        var secondBestAvg = Double.NEGATIVE_INFINITY

        for (col in 0 until maxCols) {
            if (col == skuCol) continue
            val average = statsFor(col).average
            if (average > bestAvg) {
                secondBestAvg = bestAvg
                secondBestCol = bestCol
                bestAvg = average
                bestCol = col
            } else if (average > secondBestAvg) {
                secondBestAvg = average
                secondBestCol = col
            }
        }

        return ColumnMapping(
            price = secondBestCol,
            qty   = skuCol,
            value = bestCol
        )
    }

    private fun adjustWithFallbackQty(
        initial: ColumnMapping,
        maxCols: Int,
        statsFor: (Int) -> ColStats
    ): ColumnMapping {
        var fallbackQty = initial.qty
        var firstNumericCol = -1

        for (col in 0 until maxCols) {
            if (statsFor(col).numericCount <= 0) continue
            if (firstNumericCol < 0) firstNumericCol = col
            if (col != initial.price) {
                fallbackQty = col
                break
            }
        }

        if (fallbackQty == initial.qty && firstNumericCol >= 0) fallbackQty = firstNumericCol
        return initial.copy(qty = fallbackQty)
    }

    // -------------------------------------------------------------------------
    // Data row parsing
    // -------------------------------------------------------------------------

    private fun parseDataRows(sheet: Sheet, headerRowIndex: Int, cols: ColumnMapping, config: ReaderConfig): List<PriceRow> {
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

    private fun isValidInvoiceSheet(sheet: Sheet, config: ReaderConfig): Boolean {
        val titleFound = (0..minOf(sheet.lastRowNum, 12)).any { rowIndex ->
            val row = sheet.getRow(rowIndex) ?: return@any false
            row.any { cell ->
                val text = normalize(CellUtils.getCellString(cell)).lowercase()
                text.isNotBlank() && config.documentKeywords.any { keyword -> text.contains(keyword) }
            }
        }
        if (!titleFound) return false

        val (date, branch) = parseDateAndBranch(sheet, config)
        if (date.isBlank() || branch.isBlank()) return false

        return detectHeaderAndColumns(sheet, config) != null
    }

    private fun classifyAndBuildRow(rawPrice: String, rawQty: String, rawValue: String, config: ReaderConfig): PriceRow? {
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

    private fun parseDateAndBranch(sheet: Sheet, config: ReaderConfig): Pair<String, String> {
        var date   = ""
        var branch = ""
        val evaluator = sheet.workbook.creationHelper.createFormulaEvaluator()
        for (row in sheet) {
            for (cell in row) {
                val text = normalize(CellUtils.getCellString(cell))
                if (date.isEmpty()   && config.dateKeywords.any { text.contains(it) }) date   = nextDateText(row, cell, evaluator)
                if (branch.isEmpty() && config.branchKeywords.any { text.contains(it) }) branch = nextCellText(row, cell)
            }
            if (date.isNotEmpty() && branch.isNotEmpty()) break
        }
        return date to branch
    }

    private fun nextCellText(row: Row, cell: Cell): String =
        normalize(CellUtils.getCellString(row.getCell(cell.columnIndex + 1)))

    private fun nextDateText(row: Row, cell: Cell, evaluator: FormulaEvaluator): String {
        val adjacentCell = row.getCell(cell.columnIndex + 1)
        return normalize(extractDateText(adjacentCell, evaluator) ?: CellUtils.getCellString(adjacentCell))
    }

    private fun extractDateText(cell: Cell?, evaluator: FormulaEvaluator): String? {
        if (cell == null) return null

        return when (cell.cellType) {
            CellType.NUMERIC ->
                cell.numericCellValue.takeIf { looksLikeExcelDateSerial(it) }?.let { outputDateFormat.format(DateUtil.getJavaDate(it)) }

            CellType.FORMULA -> {
                val evaluated = runCatching { evaluator.evaluate(cell) }.getOrNull()
                when (evaluated?.cellType) {
                    CellType.NUMERIC ->
                        evaluated.numberValue.takeIf { looksLikeExcelDateSerial(it) }?.let { outputDateFormat.format(DateUtil.getJavaDate(it)) }

                    else -> null
                }
            }

            else -> null
        }
    }

    private fun looksLikeExcelDateSerial(value: Double): Boolean =
        DateUtil.isValidExcelDate(value) && abs(value - value.toLong()) < 0.0000001 && value in 20000.0..60000.0

    // -------------------------------------------------------------------------
    // Row classification predicates
    // -------------------------------------------------------------------------

    private fun isTotalRow(rawPrice: String, config: ReaderConfig): Boolean {
        val normalised = normalize(rawPrice)
        return config.totalKeywords.any { kw -> normalised.equals(kw, ignoreCase = true) || normalised.contains(kw) }
    }

    private fun isGpRow(rawPrice: String, rawQty: String, config: ReaderConfig): Boolean =
        normalize(rawPrice).equals(config.gpKeyword, ignoreCase = true) && rawQty.isNotEmpty()

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
        var bestCol = -1
        var bestCount = 0

        for (col in 0 until maxCols) {
            var count = 0
            for (r in range) {
                if (isSkuLike(normalize(CellUtils.getCellString(sheet.getRow(r)?.getCell(col))))) count++
            }
            if (count > bestCount) {
                bestCount = count
                bestCol = col
            }
        }

        return bestCol.takeIf { bestCount > 0 } ?: -1
    }

    // -------------------------------------------------------------------------
    // Diagnostics
    // -------------------------------------------------------------------------

    private fun buildDiagnosticReport(sheet: Sheet, file: File, config: ReaderConfig): String {
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
        headerNoiseTokens.fold(normalize(CellUtils.getCellString(cell)).lowercase()) { acc, token -> acc.replace(token, "") }.trim()

    private fun normalize(s: String?): String =
        s?.trim()?.replace("\u00A0", " ")?.replace(whitespaceRegex, " ") ?: ""

    private fun normalizeNumber(s: String?): String =
        s?.trim()?.replace("\u00A0", " ")?.replace(numericCleanupRegex, "")?.replace("\u200B", "") ?: ""
}
