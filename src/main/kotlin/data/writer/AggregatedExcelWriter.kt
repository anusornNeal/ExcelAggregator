package data.writer

import domain.model.SheetData
import domain.model.formatGpPercentLabel
import domain.model.getDetailRows
import domain.model.getGpSummaryText
import domain.model.getGpTotalValue
import domain.model.getGpTotalsByPercent
import domain.model.getTotalRow
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/** Writes aggregated SheetData into a single accountant-friendly workbook. */
object AggregatedExcelWriter {
    private val generatedAtFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm")

    fun write(sheets: List<SheetData>, outputFile: File) {
        require(sheets.isNotEmpty()) { "No sheets to aggregate" }

        XSSFWorkbook().use { wb ->
            val styles = Styles.create(wb)
            writeAccountantView(wb.createSheet("Accountant View"), sheets, styles)
            FileOutputStream(outputFile).use { wb.write(it) }
        }
    }

    private fun writeAccountantView(ws: org.apache.poi.ss.usermodel.Sheet, sheets: List<SheetData>, styles: Styles) {
        var rowIndex = 0
        val summaryHeaders = listOf("File Name", "Sheet Name", "Branch", "Date", "Lines", "Total Qty", "Total Value", "Total GP")
        val grandQty = sheets.sumOf { it.getTotalRow()?.qty ?: 0.0 }
        val grandValue = sheets.sumOf { it.getTotalRow()?.value ?: 0.0 }
        val grandGp = sheets.sumOf { it.getGpTotalValue() ?: 0.0 }
        val gpPercents = sheets
            .flatMap { it.getGpTotalsByPercent().keys }
            .distinct()
            .sorted()
        val grandGpByPercent = gpPercents.associateWith { percent ->
            sheets.sumOf { it.getGpTotalsByPercent()[percent] ?: 0.0 }
        }

        ws.createRow(rowIndex++).apply {
            createTextCell(0, "Combined Invoice Summary", styles.title)
        }
        ws.addMergedRegion(CellRangeAddress(0, 0, 0, maxOf(7, gpPercents.size * 2 - 1)))

        ws.createRow(rowIndex++).apply {
            createTextCell(0, "Generated", styles.metaLabel)
            createTextCell(1, LocalDateTime.now().format(generatedAtFormatter), styles.metaValue)
            createTextCell(4, "Invoice sheets", styles.metaLabel)
            createNumberCell(5, sheets.size.toDouble(), styles.metaInteger)
        }

        ws.createRow(rowIndex++).apply {
            createTextCell(0, "Total quantity", styles.kpiLabel)
            createNumberCell(1, grandQty, styles.kpiValue)
            createTextCell(2, "Total value", styles.kpiLabel)
            createNumberCell(3, grandValue, styles.kpiValue)
            createTextCell(4, "Total GP", styles.kpiLabel)
            createNumberCell(5, grandGp, styles.kpiValue)
        }

        if (gpPercents.isNotEmpty()) {
            ws.createRow(rowIndex++).apply {
                var col = 0
                gpPercents.forEach { percent ->
                    createTextCell(col++, formatGpPercentLabel(percent), styles.kpiLabel)
                    createNumberCell(col++, grandGpByPercent[percent] ?: 0.0, styles.kpiValue)
                }
            }
        }

        rowIndex++
        ws.createRow(rowIndex++).apply {
            createTextCell(0, "Sheet Summary", styles.sectionBanner)
        }
        ws.addMergedRegion(CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, summaryHeaders.lastIndex))

        writeHeaderRow(ws.createRow(rowIndex++), summaryHeaders, styles.header)
        sheets.forEach { sheet ->
            val total = sheet.getTotalRow()
            ws.createRow(rowIndex++).apply {
                createTextCell(0, sheet.sourceFileName.orEmpty(), styles.text)
                createTextCell(1, sheet.sourceSheetName.orEmpty(), styles.text)
                createTextCell(2, sheet.branch.ifBlank { "-" }, styles.text)
                createTextCell(3, sheet.date.ifBlank { "-" }, styles.text)
                createNumberCell(4, sheet.getDetailRows().size.toDouble(), styles.integer)
                createOptionalNumberCell(5, total?.qty, styles.number)
                createOptionalNumberCell(6, total?.value, styles.number)
                createOptionalNumberCell(7, sheet.getGpTotalValue(), styles.number)
            }
        }

        rowIndex += 2
        ws.createRow(rowIndex++).apply {
            createTextCell(0, "Details By File", styles.sectionBanner)
        }
        ws.addMergedRegion(CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 7))

        sheets.groupBy { it.sourceFileName.orEmpty() }.forEach { (fileName, fileSheets) ->
            val fileQty = fileSheets.sumOf { it.getTotalRow()?.qty ?: 0.0 }
            val fileValue = fileSheets.sumOf { it.getTotalRow()?.value ?: 0.0 }
            val fileGp = fileSheets.sumOf { it.getGpTotalValue() ?: 0.0 }

            ws.createRow(rowIndex++).apply {
                createTextCell(0, fileName.ifBlank { "(Unnamed file)" }, styles.fileBanner)
                createTextCell(4, "Qty", styles.fileBanner)
                createNumberCell(5, fileQty, styles.fileBannerNumber)
                createTextCell(6, "Value / GP", styles.fileBanner)
                createTextCell(7, formatMoneyPair(fileValue, fileGp), styles.fileBanner)
            }

            fileSheets.forEach { sheet ->
                ws.createRow(rowIndex++).apply {
                    createTextCell(0, "Sheet", styles.label)
                    createTextCell(1, sheet.sourceSheetName.orEmpty(), styles.text)
                    createTextCell(2, "Branch", styles.label)
                    createTextCell(3, sheet.branch.ifBlank { "-" }, styles.text)
                    createTextCell(4, "Date", styles.label)
                    createTextCell(5, sheet.date.ifBlank { "-" }, styles.text)
                    createTextCell(6, "Total GP", styles.label)
                    createOptionalNumberCell(7, sheet.getGpTotalValue(), styles.number)
                }

                sheet.getGpTotalsByPercent().toSortedMap().forEach { (percent, value) ->
                    ws.createRow(rowIndex++).apply {
                        createTextCell(6, formatGpPercentLabel(percent), styles.label)
                        createNumberCell(7, value, styles.number)
                    }
                }

                writeHeaderRow(ws.createRow(rowIndex++), listOf("Price", "Quantity", "Value"), styles.header)
                sheet.getDetailRows().forEach { detail ->
                    ws.createRow(rowIndex++).apply {
                        createOptionalNumberCell(0, detail.price, styles.number)
                        createOptionalNumberCell(1, detail.qty, styles.number)
                        createOptionalNumberCell(2, detail.value, styles.number)
                    }
                }

                sheet.getTotalRow()?.let { total ->
                    ws.createRow(rowIndex++).apply {
                        createTextCell(0, "Sheet Total", styles.totalLabel)
                        createOptionalNumberCell(1, total.qty, styles.totalNumber)
                        createOptionalNumberCell(2, total.value, styles.totalNumber)
                    }
                }

                sheet.getGpTotalValue()?.let { gp ->
                    ws.createRow(rowIndex++).apply {
                        createTextCell(0, "Sheet GP", styles.totalLabel)
                        createNumberCell(2, gp, styles.totalNumber)
                    }
                }

                rowIndex++
            }
        }

        ws.createFreezePane(0, 6)
        setColumnWidths(ws, listOf(28, 24, 22, 14, 12, 14, 14, 20))
    }

    private fun formatMoneyPair(value: Double, gp: Double): String =
        "${formatNumber(value)} / ${formatNumber(gp)}"

    private fun formatNumber(value: Double): String =
        "%,.2f".format(value)

    private fun writeHeaderRow(row: org.apache.poi.ss.usermodel.Row, values: List<String>, style: CellStyle) {
        values.forEachIndexed { index, value -> row.createTextCell(index, value, style) }
    }

    private fun setColumnWidths(ws: org.apache.poi.ss.usermodel.Sheet, widths: List<Int>) {
        widths.forEachIndexed { index, width -> ws.setColumnWidth(index, width * 256) }
    }

    private fun org.apache.poi.ss.usermodel.Row.createTextCell(column: Int, value: String, style: CellStyle): Cell =
        createCell(column).apply {
            setCellValue(value)
            cellStyle = style
        }

    private fun org.apache.poi.ss.usermodel.Row.createNumberCell(column: Int, value: Double, style: CellStyle): Cell =
        createCell(column).apply {
            setCellValue(value)
            cellStyle = style
        }

    private fun org.apache.poi.ss.usermodel.Row.createOptionalNumberCell(column: Int, value: Double?, style: CellStyle): Cell =
        createCell(column).apply {
            if (value == null) setBlank() else setCellValue(value)
            cellStyle = style
        }

    private data class Styles(
        val title: CellStyle,
        val metaLabel: CellStyle,
        val metaValue: CellStyle,
        val metaInteger: CellStyle,
        val kpiLabel: CellStyle,
        val kpiValue: CellStyle,
        val sectionBanner: CellStyle,
        val fileBanner: CellStyle,
        val fileBannerNumber: CellStyle,
        val header: CellStyle,
        val label: CellStyle,
        val text: CellStyle,
        val number: CellStyle,
        val integer: CellStyle,
        val totalLabel: CellStyle,
        val totalNumber: CellStyle
    ) {
        companion object {
            fun create(wb: XSSFWorkbook): Styles {
                val titleFont = wb.createFont().apply {
                    bold = true
                    fontHeightInPoints = 15
                    color = IndexedColors.WHITE.index
                }
                val boldFont = wb.createFont().apply { bold = true }
                val whiteBoldFont = wb.createFont().apply {
                    bold = true
                    color = IndexedColors.WHITE.index
                }

                val decimalFormat = wb.creationHelper.createDataFormat().getFormat("#,##0.00")
                val integerFormat = wb.creationHelper.createDataFormat().getFormat("#,##0")

                fun bordered(): CellStyle = wb.createCellStyle().apply {
                    verticalAlignment = VerticalAlignment.CENTER
                    borderTop = BorderStyle.THIN
                    borderBottom = BorderStyle.THIN
                    borderLeft = BorderStyle.THIN
                    borderRight = BorderStyle.THIN
                }

                val title = bordered().apply {
                    setFont(titleFont)
                    alignment = HorizontalAlignment.CENTER
                    fillForegroundColor = IndexedColors.DARK_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val metaLabel = bordered().apply {
                    setFont(boldFont)
                    fillForegroundColor = IndexedColors.GREY_25_PERCENT.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val metaValue = bordered().apply {
                    alignment = HorizontalAlignment.LEFT
                }
                val metaInteger = bordered().apply {
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = integerFormat
                }
                val kpiLabel = bordered().apply {
                    setFont(boldFont)
                    fillForegroundColor = IndexedColors.LIGHT_TURQUOISE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val kpiValue = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                    fillForegroundColor = IndexedColors.LIGHT_TURQUOISE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val sectionBanner = bordered().apply {
                    setFont(whiteBoldFont)
                    fillForegroundColor = IndexedColors.TEAL.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val fileBanner = bordered().apply {
                    setFont(whiteBoldFont)
                    fillForegroundColor = IndexedColors.BLUE_GREY.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val fileBannerNumber = bordered().apply {
                    setFont(whiteBoldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                    fillForegroundColor = IndexedColors.BLUE_GREY.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val header = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.CENTER
                    fillForegroundColor = IndexedColors.GREY_40_PERCENT.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val label = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.RIGHT
                    fillForegroundColor = IndexedColors.LEMON_CHIFFON.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val text = bordered().apply {
                    alignment = HorizontalAlignment.LEFT
                    wrapText = true
                }
                val number = bordered().apply {
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                }
                val integer = bordered().apply {
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = integerFormat
                }
                val totalLabel = bordered().apply {
                    setFont(boldFont)
                    fillForegroundColor = IndexedColors.LIGHT_CORNFLOWER_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val totalNumber = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                    fillForegroundColor = IndexedColors.LIGHT_CORNFLOWER_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }

                return Styles(
                    title,
                    metaLabel,
                    metaValue,
                    metaInteger,
                    kpiLabel,
                    kpiValue,
                    sectionBanner,
                    fileBanner,
                    fileBannerNumber,
                    header,
                    label,
                    text,
                    number,
                    integer,
                    totalLabel,
                    totalNumber
                )
            }
        }
    }
}
