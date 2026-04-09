package data.writer

import data.config.ConfigLoader
import data.reader.CellUtils.setCellDoubleOrBlank
import domain.model.AggregatorConfig
import domain.model.PriceRow
import domain.model.SheetData
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress

/** Row-level rendering helpers for the aggregated Excel sheet. */
object SheetRowWriter {

    fun writeFileInfoRow(
        ws: Sheet, rowIdx: Int, startCol: Int, endCol: Int,
        fileIndex: Int, totalFiles: Int, style: CellStyle,
        config: AggregatorConfig = ConfigLoader.load()
    ) {
        val row = ws.getRow(rowIdx) ?: ws.createRow(rowIdx)
        row.createCell(startCol).apply {
            setCellValue("${config.labelFilePrefix}$fileIndex/$totalFiles".trim())
            cellStyle = style
        }
        if (endCol > startCol) ws.addMergedRegion(CellRangeAddress(rowIdx, rowIdx, startCol, endCol))
    }

    fun writeMergedFileNameRow(ws: Sheet, rowIdx: Int, startCol: Int, endCol: Int, fileName: String, style: CellStyle) {
        val row = ws.getRow(rowIdx) ?: ws.createRow(rowIdx)
        row.createCell(startCol).apply { setCellValue(fileName); cellStyle = style }
        if (endCol > startCol) ws.addMergedRegion(CellRangeAddress(rowIdx, rowIdx, startCol, endCol))
    }

    fun writeSheetNameRow(ws: Sheet, rowIdx: Int, colIdx: Int, sheetName: String, style: CellStyle) {
        val row = ws.getRow(rowIdx) ?: ws.createRow(rowIdx)
        row.createCell(colIdx).apply { setCellValue(sheetName); cellStyle = style }
        ws.addMergedRegion(CellRangeAddress(rowIdx, rowIdx, colIdx, colIdx + 2))
    }

    fun writeDateBranchRows(
        ws: Sheet, rowDate: Int, rowBranch: Int,
        colBase: Int, sheet: SheetData, style: CellStyle,
        config: AggregatorConfig = ConfigLoader.load()
    ) {
        ws.getRow(rowDate)?.let { row ->
            row.createCell(colBase).apply { setCellValue(config.labelDate); cellStyle = style }
            row.createCell(colBase + 1).apply { setCellValue(sheet.date) }
        }
        ws.getRow(rowBranch)?.let { row ->
            row.createCell(colBase).apply { setCellValue(config.labelBranch); cellStyle = style }
            row.createCell(colBase + 1).apply { setCellValue(sheet.branch) }
        }
    }

    fun writeColumnHeaderRow(ws: Sheet, rowIdx: Int, colBase: Int, style: CellStyle, config: AggregatorConfig = ConfigLoader.load()) {
        ws.getRow(rowIdx)?.let { row ->
            row.createCell(colBase).apply { setCellValue(config.labelPrice); cellStyle = style }
            row.createCell(colBase + 1).apply { setCellValue(config.labelQty); cellStyle = style }
            row.createCell(colBase + 2).apply { setCellValue(config.labelValue); cellStyle = style }
        }
    }

    fun writePriceRows(
        ws: Sheet, rowStart: Int, colBase: Int,
        allPrices: List<Double>, priceMap: Map<Double?, PriceRow>, style: CellStyle
    ) {
        allPrices.forEachIndexed { priceIdx, price ->
            val row = ws.getRow(rowStart + priceIdx) ?: ws.createRow(rowStart + priceIdx)
            val d = priceMap[price]
            row.createCell(colBase).apply { setCellValue(price); cellStyle = style }
            row.createCell(colBase + 1).apply { setCellDoubleOrBlank(this, d?.qty); cellStyle = style }
            row.createCell(colBase + 2).apply { setCellDoubleOrBlank(this, d?.value); cellStyle = style }
        }
    }

    fun writeTotalRow(
        ws: Sheet, rowIdx: Int, colBase: Int,
        totalQty: Double?, totalValue: Double?, style: CellStyle,
        config: AggregatorConfig = ConfigLoader.load()
    ) {
        ws.getRow(rowIdx)?.apply {
            createCell(colBase).apply { setCellValue(config.labelTotal); cellStyle = style }
            createCell(colBase + 1).apply { setCellDoubleOrBlank(this, totalQty); cellStyle = style }
            createCell(colBase + 2).apply { setCellDoubleOrBlank(this, totalValue); cellStyle = style }
        }
    }

    fun writeGpRow(
        ws: Sheet, rowIdx: Int, colBase: Int,
        gpResults: Map<Double, Pair<Double?, Double?>>,
        style: CellStyle, config: AggregatorConfig = ConfigLoader.load()
    ) {
        ws.getRow(rowIdx)?.apply {
            createCell(colBase).apply { setCellValue(config.labelGp); cellStyle = style }
            createCell(colBase + 1).apply {
                val pcts = gpResults.values.mapNotNull { it.first?.toString() }
                setCellValue(pcts.joinToString(", "))
                cellStyle = style
            }
            createCell(colBase + 2).apply {
                val values = gpResults.values.mapNotNull { it.second }
                if (values.isEmpty()) setBlank() else setCellValue(values.sum())
                cellStyle = style
            }
        }
    }

    fun writeSummaryRows(
        ws: Sheet, rowHeader: Int, rowValue: Int, rowGpHeader: Int, rowGpValue: Int, style: CellStyle,
        sumQty: Double, sumValue: Double, sumGpByTarget: Map<Double, Double>,
        config: AggregatorConfig = ConfigLoader.load()
    ) {
        ws.getRow(rowHeader)?.apply {
            createCell(0).apply { setCellValue("Total"); cellStyle = style }
            createCell(1).apply { setCellValue(config.labelQty); cellStyle = style }
            createCell(2).apply { setCellValue(config.labelValue); cellStyle = style }
        }
        ws.getRow(rowValue)?.apply {
            createCell(1).apply { setCellValue(sumQty); cellStyle = style }
            createCell(2).apply { setCellValue(sumValue); cellStyle = style }
        }
        val targets = config.gpTargets
        ws.getRow(rowGpHeader)?.apply {
            targets.forEachIndexed { i, t ->
                createCell(i).apply { setCellValue("GP $t"); cellStyle = style }
            }
        }
        ws.getRow(rowGpValue)?.apply {
            targets.forEachIndexed { i, t ->
                val sum = sumGpByTarget[t] ?: 0.0
                createCell(i).apply { if (sum == 0.0) setBlank() else setCellValue(sum); cellStyle = style }
            }
        }
    }
}

