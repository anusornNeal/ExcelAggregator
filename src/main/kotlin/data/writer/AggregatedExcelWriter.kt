package data.writer

import domain.model.SheetData
import domain.model.getDetailRows
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
import kotlin.math.abs

/** Writes aggregated invoice sheets into summary and price-detail workbook tabs. */
object AggregatedExcelWriter {
    private const val SUMMARY_SHEET_NAME = "สรุปรวม (Summary)"
    private const val DETAIL_SHEET_NAME = "รายละเอียดราคา (Price Detail)"

    fun write(sheets: List<SheetData>, outputFile: File) {
        require(sheets.isNotEmpty()) { "No sheets to aggregate" }

        XSSFWorkbook().use { wb ->
            val styles = Styles.create(wb)
            writeSummarySheet(wb.createSheet(SUMMARY_SHEET_NAME), sheets, styles)
            writePriceDetailSheet(wb.createSheet(DETAIL_SHEET_NAME), sheets, styles)
            FileOutputStream(outputFile).use { wb.write(it) }
        }
    }

    private fun writeSummarySheet(ws: org.apache.poi.ss.usermodel.Sheet, sheets: List<SheetData>, styles: Styles) {
        val headers = listOf(
            "#",
            "File Name",
            "Sheet Name",
            "สาขา",
            "จำนวนรายการ",
            "จำนวนตัว",
            "มูลค่า (บาท)",
            "GP Rate",
            "GP (บาท)",
            "หมายเหตุ"
        )
        val grandQty = sheets.sumOf { it.getTotalRow()?.qty ?: 0.0 }
        val grandValue = sheets.sumOf { it.getTotalRow()?.value ?: 0.0 }
        val grandGp = sheets.sumOf { it.getGpTotalValue() ?: 0.0 }

        ws.createRow(0).apply {
            createTextCell(0, "สรุปรวมใบส่งสินค้า Depart Counter - ทุกสาขา", styles.title)
        }
        ws.addMergedRegion(CellRangeAddress(0, 0, 0, headers.lastIndex))

        ws.createRow(1).apply {
            createTextCell(0, "รวม ${sheets.size} ชีท", styles.metaLabel)
            createTextCell(1, "ข้อมูลจาก ${sheets.map { fileNameOf(it) }.distinct().size} ไฟล์", styles.metaValue)
        }
        ws.addMergedRegion(CellRangeAddress(1, 1, 1, 3))

        writeHeaderRow(ws.createRow(3), headers, styles.header)

        var rowIndex = 4
        sheets.forEachIndexed { index, sheet ->
            val total = sheet.getTotalRow()
            ws.createRow(rowIndex++).apply {
                createNumberCell(0, (index + 1).toDouble(), styles.integer)
                createTextCell(1, fileNameOf(sheet), styles.fileText)
                createTextCell(2, sheetNameOf(sheet), styles.text)
                createTextCell(3, sheet.branch.ifBlank { "-" }, styles.text)
                createNumberCell(4, sheet.getDetailRows().size.toDouble(), styles.integer)
                createOptionalNumberCell(5, total?.qty, styles.integer)
                createOptionalNumberCell(6, total?.value, styles.number)
                createTextCell(7, gpRateText(sheet), styles.text)
                createOptionalNumberCell(8, sheet.getGpTotalValue(), styles.number)
                createTextCell(9, "", styles.text)
            }
        }

        ws.createRow(rowIndex).apply {
            createTextCell(0, "รวมทั้งหมด (Grand Total)", styles.grandLabel)
            createOptionalNumberCell(5, grandQty, styles.grandInteger)
            createOptionalNumberCell(6, grandValue, styles.grandNumber)
            createOptionalNumberCell(8, grandGp, styles.grandNumber)
        }
        ws.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, 0, 4))

        ws.createFreezePane(0, 4)
        setColumnWidths(ws, listOf(6, 28, 24, 30, 14, 14, 16, 14, 16, 18))
    }

    private fun writePriceDetailSheet(ws: org.apache.poi.ss.usermodel.Sheet, sheets: List<SheetData>, styles: Styles) {
        val headers = listOf(
            "File Name",
            "Sheet Name",
            "สาขา",
            "ราคา (บาท)",
            "จำนวนตัว",
            "มูลค่า (บาท)",
            "GP Rate",
            "GP (บาท)"
        )
        val grandQty = sheets.sumOf { it.getTotalRow()?.qty ?: 0.0 }
        val grandValue = sheets.sumOf { it.getTotalRow()?.value ?: 0.0 }
        val grandGp = sheets.sumOf { it.getGpTotalValue() ?: 0.0 }

        ws.createRow(0).apply {
            createTextCell(0, "ตารางรายละเอียดราคาแยกตามสาขา", styles.title)
        }
        ws.addMergedRegion(CellRangeAddress(0, 0, 0, headers.lastIndex))

        writeHeaderRow(ws.createRow(2), headers, styles.header)

        var rowIndex = 3
        sheets.groupBy(::fileNameOf).entries.forEachIndexed { fileIndex, (fileName, fileSheets) ->
            val rowStyles = styles.groupStyles[fileIndex % styles.groupStyles.size]
            val fileQty = fileSheets.sumOf { it.getTotalRow()?.qty ?: 0.0 }
            val fileValue = fileSheets.sumOf { it.getTotalRow()?.value ?: 0.0 }
            val fileGp = fileSheets.sumOf { it.getGpTotalValue() ?: 0.0 }

            ws.createRow(rowIndex).apply {
                createTextCell(0, fileName, rowStyles.fileHeader)
                createTextCell(1, "", rowStyles.fileHeader)
                createTextCell(2, "", rowStyles.fileHeader)
                createTextCell(3, "File Total", rowStyles.fileHeader)
                createOptionalNumberCell(4, fileQty, rowStyles.fileHeaderInteger)
                createOptionalNumberCell(5, fileValue, rowStyles.fileHeaderNumber)
                createTextCell(6, "${fileSheets.size} sheets", rowStyles.fileHeader)
                createOptionalNumberCell(7, fileGp, rowStyles.fileHeaderNumber)
            }
            ws.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, 0, 2))
            rowIndex++

            fileSheets.forEach { sheet ->
                val detailRows = sheet.getDetailRows()
                detailRows.forEachIndexed { detailIndex, detail ->
                    val showGroupInfo = detailIndex == 0
                    ws.createRow(rowIndex++).apply {
                        createTextCell(0, if (showGroupInfo) fileNameOf(sheet) else "", rowStyles.fileText)
                        createTextCell(1, if (showGroupInfo) sheetNameOf(sheet) else "", rowStyles.text)
                        createTextCell(2, if (showGroupInfo) sheet.branch.ifBlank { "-" } else "", rowStyles.text)
                        createOptionalNumberCell(3, detail.price, rowStyles.number)
                        createOptionalNumberCell(4, detail.qty, rowStyles.integer)
                        createOptionalNumberCell(5, detail.value, rowStyles.number)
                        createTextCell(6, "", rowStyles.text)
                        createTextCell(7, "", rowStyles.number)
                    }
                }

                val total = sheet.getTotalRow()
                ws.createRow(rowIndex).apply {
                    createTextCell(0, "รวม : ${sheetNameOf(sheet)}", rowStyles.sheetTotalLabel)
                    createOptionalNumberCell(4, total?.qty, rowStyles.sheetTotalInteger)
                    createOptionalNumberCell(5, total?.value, rowStyles.sheetTotalNumber)
                    createTextCell(6, gpRateText(sheet), rowStyles.sheetTotalLabel)
                    createOptionalNumberCell(7, sheet.getGpTotalValue(), rowStyles.sheetTotalNumber)
                }
                ws.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, 0, 3))
                rowIndex++
            }
        }

        ws.createRow(rowIndex).apply {
            createTextCell(0, "GRAND TOTAL - ทุกสาขา", styles.grandLabel)
            createOptionalNumberCell(4, grandQty, styles.grandInteger)
            createOptionalNumberCell(5, grandValue, styles.grandNumber)
            createOptionalNumberCell(7, grandGp, styles.grandNumber)
        }
        ws.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, 0, 3))

        ws.createFreezePane(0, 3)
        setColumnWidths(ws, listOf(28, 24, 32, 14, 14, 16, 14, 16))
    }

    private fun writeHeaderRow(row: org.apache.poi.ss.usermodel.Row, values: List<String>, style: CellStyle) {
        values.forEachIndexed { index, value -> row.createTextCell(index, value, style) }
    }

    private fun setColumnWidths(ws: org.apache.poi.ss.usermodel.Sheet, widths: List<Int>) {
        widths.forEachIndexed { index, width -> ws.setColumnWidth(index, width * 256) }
    }

    private fun fileNameOf(sheet: SheetData): String =
        sheet.sourceFileName?.takeIf { it.isNotBlank() } ?: "-"

    private fun sheetNameOf(sheet: SheetData): String =
        sheet.sourceSheetName?.takeIf { it.isNotBlank() } ?: "-"

    private fun gpRateText(sheet: SheetData): String =
        sheet.getGpTotalsByPercent().keys
            .sorted()
            .map(::formatPercent)
            .ifEmpty { listOf("-") }
            .joinToString(", ")

    private fun formatPercent(value: Double): String {
        val displayValue = if (abs(value) <= 1.0) value * 100 else value
        val formatted = if (displayValue % 1.0 == 0.0) displayValue.toLong().toString() else displayValue.toString()
        return "$formatted%"
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

    private data class GroupStyles(
        val fileHeader: CellStyle,
        val fileHeaderNumber: CellStyle,
        val fileHeaderInteger: CellStyle,
        val fileText: CellStyle,
        val text: CellStyle,
        val number: CellStyle,
        val integer: CellStyle,
        val sheetTotalLabel: CellStyle,
        val sheetTotalNumber: CellStyle,
        val sheetTotalInteger: CellStyle
    )

    private data class Styles(
        val title: CellStyle,
        val metaLabel: CellStyle,
        val metaValue: CellStyle,
        val header: CellStyle,
        val fileText: CellStyle,
        val text: CellStyle,
        val number: CellStyle,
        val integer: CellStyle,
        val totalLabel: CellStyle,
        val totalNumber: CellStyle,
        val totalInteger: CellStyle,
        val grandLabel: CellStyle,
        val grandNumber: CellStyle,
        val grandInteger: CellStyle,
        val groupStyles: List<GroupStyles>
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

                fun filled(fill: IndexedColors, dataFormat: Short? = null): CellStyle =
                    bordered().apply {
                        alignment = HorizontalAlignment.LEFT
                        wrapText = true
                        fillForegroundColor = fill.index
                        fillPattern = FillPatternType.SOLID_FOREGROUND
                        dataFormat?.let { this.dataFormat = it }
                    }

                fun filledNumber(fill: IndexedColors, format: Short): CellStyle =
                    filled(fill, format).apply {
                        alignment = HorizontalAlignment.RIGHT
                    }

                fun filledBold(fill: IndexedColors, dataFormat: Short? = null): CellStyle =
                    filled(fill, dataFormat).apply {
                        setFont(boldFont)
                    }

                fun filledBoldNumber(fill: IndexedColors, format: Short): CellStyle =
                    filledBold(fill, format).apply {
                        alignment = HorizontalAlignment.RIGHT
                    }

                val title = bordered().apply {
                    setFont(titleFont)
                    alignment = HorizontalAlignment.CENTER
                    fillForegroundColor = IndexedColors.DARK_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val metaLabel = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.CENTER
                    fillForegroundColor = IndexedColors.LIGHT_TURQUOISE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val metaValue = bordered().apply {
                    alignment = HorizontalAlignment.LEFT
                    fillForegroundColor = IndexedColors.LIGHT_TURQUOISE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val header = bordered().apply {
                    setFont(whiteBoldFont)
                    alignment = HorizontalAlignment.CENTER
                    fillForegroundColor = IndexedColors.BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val text = bordered().apply {
                    alignment = HorizontalAlignment.LEFT
                    wrapText = true
                }
                val fileText = bordered().apply {
                    alignment = HorizontalAlignment.LEFT
                    wrapText = false
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
                    fillForegroundColor = IndexedColors.LIGHT_GREEN.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val totalNumber = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                    fillForegroundColor = IndexedColors.LIGHT_GREEN.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val totalInteger = bordered().apply {
                    setFont(boldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = integerFormat
                    fillForegroundColor = IndexedColors.LIGHT_GREEN.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val grandLabel = bordered().apply {
                    setFont(whiteBoldFont)
                    fillForegroundColor = IndexedColors.DARK_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val grandNumber = bordered().apply {
                    setFont(whiteBoldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = decimalFormat
                    fillForegroundColor = IndexedColors.DARK_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val grandInteger = bordered().apply {
                    setFont(whiteBoldFont)
                    alignment = HorizontalAlignment.RIGHT
                    dataFormat = integerFormat
                    fillForegroundColor = IndexedColors.DARK_BLUE.index
                    fillPattern = FillPatternType.SOLID_FOREGROUND
                }
                val groupFills = listOf(IndexedColors.WHITE)
                val groupStyles = groupFills.map { fill ->
                    GroupStyles(
                        fileHeader = filledBold(IndexedColors.GREY_25_PERCENT).apply {
                            wrapText = false
                        },
                        fileHeaderNumber = filledBoldNumber(IndexedColors.GREY_25_PERCENT, decimalFormat),
                        fileHeaderInteger = filledBoldNumber(IndexedColors.GREY_25_PERCENT, integerFormat),
                        fileText = filled(fill).apply { wrapText = false },
                        text = filled(fill),
                        number = filledNumber(fill, decimalFormat),
                        integer = filledNumber(fill, integerFormat),
                        sheetTotalLabel = filledBold(fill),
                        sheetTotalNumber = filledBoldNumber(fill, decimalFormat),
                        sheetTotalInteger = filledBoldNumber(fill, integerFormat)
                    )
                }

                return Styles(
                    title = title,
                    metaLabel = metaLabel,
                    metaValue = metaValue,
                    header = header,
                    fileText = fileText,
                    text = text,
                    number = number,
                    integer = integer,
                    totalLabel = totalLabel,
                    totalNumber = totalNumber,
                    totalInteger = totalInteger,
                    grandLabel = grandLabel,
                    grandNumber = grandNumber,
                    grandInteger = grandInteger,
                    groupStyles = groupStyles
                )
            }
        }
    }
}
