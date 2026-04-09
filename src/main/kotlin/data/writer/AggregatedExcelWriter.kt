package data.writer

import data.config.ConfigLoader
import domain.model.SheetData
import domain.model.getGpValueForTarget
import domain.model.getTotalRow
import domain.model.toPriceMap
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

/** Writes aggregated SheetData list into a single output Excel file. */
object AggregatedExcelWriter {

    fun write(sheets: List<SheetData>, outputFile: File) {
        require(sheets.isNotEmpty()) { "No sheets to aggregate" }
        val wb = XSSFWorkbook()
        val ws = wb.createSheet("Aggregated")
        val styles = Styles(
            header    = ExcelStyleFactory.headerStyle(wb),
            total     = ExcelStyleFactory.totalStyle(wb),
            data      = ExcelStyleFactory.dataStyle(wb),
            label     = ExcelStyleFactory.labelStyle(wb),
            sheetName = ExcelStyleFactory.sheetNameStyle(wb)
        )
        val layout = Layout.from(sheets)
        layout.initRows(ws)
        val totals = AggregationTotals()
        sheets.forEachIndexed { idx, sheet -> sheet.renderTo(ws, idx, layout, styles, totals) }
        SheetRowWriter.writeSummaryRows(
            ws, layout.summaryHeader, layout.summaryValue,
            layout.summaryGpHeader, layout.summaryGpValue, styles.total,
            totals.sumTotalQty, totals.sumTotalValue,
            totals.sumGpByTarget
        )
        FileOutputStream(outputFile).use { wb.write(it) }
        wb.close()
    }

    // -------------------------------------------------------------------------
    // Internal models
    // -------------------------------------------------------------------------

    private data class Styles(
        val header: CellStyle, val total: CellStyle,
        val data: CellStyle, val label: CellStyle, val sheetName: CellStyle
    )

    private data class AggregationTotals(
        var sumTotalQty: Double = 0.0,
        var sumTotalValue: Double = 0.0,
        val sumGpByTarget: MutableMap<Double, Double> = mutableMapOf()
    )

    // -------------------------------------------------------------------------
    // Layout
    // -------------------------------------------------------------------------

    private data class Layout(
        val allPrices: List<Double>,
        val fileCount: Int,
        val totalFiles: Int,
        val sheetToFileIndex: IntArray,
        val fileLastSheet: IntArray,
        // row indices (0-based)
        val fileInfo: Int = 0,
        val fileName: Int = 1,
        val sheetNameRow: Int = 2,
        val date: Int = 3,
        val branch: Int = 4,
        val header: Int = 5,
        val dataStart: Int = 6,
        val total: Int,
        val gp: Int,
        val summaryHeader: Int,
        val summaryValue: Int,
        val summaryGpHeader: Int,
        val summaryGpValue: Int
    ) {
        companion object {
            fun from(sheets: List<SheetData>): Layout {
                val allPrices = sheets
                    .flatMap { it.rows.filter { r -> !r.isTotal && !r.isGP } }
                    .mapNotNull { it.price }.toSortedSet().toList()

                val dataStart    = 6
                val total        = dataStart + allPrices.size
                val gp           = total + 1
                val summaryHeader     = gp + 2
                val summaryGpHeader   = summaryHeader + 2

                val fileNames     = sheets.map { it.sourceFileName ?: "" }
                val distinctFiles = fileNames.distinct()
                val sheetToFileIndex = IntArray(sheets.size) { distinctFiles.indexOf(fileNames[it]) + 1 }

                val totalFiles   = distinctFiles.size
                val fileLastSheet = IntArray(totalFiles + 1) { -1 }
                sheetToFileIndex.forEachIndexed { idx, fi ->
                    if (idx > fileLastSheet[fi]) fileLastSheet[fi] = idx
                }

                return Layout(
                    allPrices        = allPrices,
                    fileCount        = sheets.size,
                    totalFiles       = totalFiles,
                    sheetToFileIndex = sheetToFileIndex,
                    fileLastSheet    = fileLastSheet,
                    total            = total,
                    gp               = gp,
                    summaryHeader    = summaryHeader,
                    summaryValue     = summaryHeader + 1,
                    summaryGpHeader  = summaryGpHeader,
                    summaryGpValue   = summaryGpHeader + 1
                )
            }
        }

        fun initRows(ws: Sheet) {
            for (r in 0..summaryGpValue) if (ws.getRow(r) == null) ws.createRow(r)
        }

        /** Starting column for a given sheet index, with gap columns inserted between different files. */
        fun colBaseFor(sheetIdx: Int): Int {
            if (sheetIdx <= 0) return 0
            var gaps = 0
            for (i in 1..sheetIdx) if (sheetToFileIndex[i] != sheetToFileIndex[i - 1]) gaps++
            return sheetIdx * 3 + gaps
        }

        fun isFirstOfFile(sheetIdx: Int): Boolean =
            sheetIdx == 0 || sheetToFileIndex[sheetIdx] != sheetToFileIndex[sheetIdx - 1]

        fun isBoundaryAfter(sheetIdx: Int): Boolean {
            val next = sheetIdx + 1
            return next < fileCount && sheetToFileIndex[next] != sheetToFileIndex[sheetIdx]
        }

        fun lastSheetIndexOf(sheetIdx: Int): Int = fileLastSheet[sheetToFileIndex[sheetIdx]]
    }

    // -------------------------------------------------------------------------
    // Per-sheet rendering
    // -------------------------------------------------------------------------

    private fun SheetData.renderTo(ws: Sheet, idx: Int, layout: Layout, styles: Styles, totals: AggregationTotals) {
        val colBase   = layout.colBaseFor(idx)
        val fileIndex = layout.sheetToFileIndex[idx]
        val lastIdx   = layout.lastSheetIndexOf(idx)
        val endCol    = layout.colBaseFor(lastIdx) + 2

        if (layout.isFirstOfFile(idx)) {
            SheetRowWriter.writeFileInfoRow(ws, layout.fileInfo, colBase, endCol, fileIndex, layout.totalFiles, styles.header)
            SheetRowWriter.writeMergedFileNameRow(ws, layout.fileName, colBase, endCol, sourceFileName ?: "", styles.header)
            for (j in idx..lastIdx) SheetRowWriter.writeColumnHeaderRow(ws, layout.header, layout.colBaseFor(j), styles.header)
        }

        SheetRowWriter.writeSheetNameRow(ws, layout.sheetNameRow, colBase, this.sourceSheetName.orEmpty(), styles.sheetName)
        SheetRowWriter.writeDateBranchRows(ws, layout.date, layout.branch, colBase, this, styles.label)
        SheetRowWriter.writePriceRows(ws, layout.dataStart, colBase, layout.allPrices, toPriceMap(), styles.data)

        val totalQty   = getTotalRow()?.qty
        val totalValue = getTotalRow()?.value
        SheetRowWriter.writeTotalRow(ws, layout.total, colBase, totalQty, totalValue, styles.total)

        val gpTargets = ConfigLoader.load().gpTargets
        val gpResults: Map<Double, Pair<Double?, Double?>> =
            gpTargets.associateWith { getGpValueForTarget(it, totalValue) }
        SheetRowWriter.writeGpRow(ws, layout.gp, colBase, gpResults, styles.data)

        if (totalQty   != null) totals.sumTotalQty  += totalQty
        if (totalValue != null) totals.sumTotalValue += totalValue
        for ((target, pair) in gpResults) {
            val value = pair.second
            if (value != null) totals.sumGpByTarget[target] = (totals.sumGpByTarget[target] ?: 0.0) + value
        }

        ws.setColumnWidth(colBase,     12 * 256)
        ws.setColumnWidth(colBase + 1, 12 * 256)
        ws.setColumnWidth(colBase + 2, 12 * 256)
        if (layout.isBoundaryAfter(idx)) ws.setColumnWidth(colBase + 3, 3 * 256)
    }
}



