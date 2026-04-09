package data.writer

import org.apache.poi.ss.usermodel.*

/** Creates reusable cell styles for the output workbook. */
object ExcelStyleFactory {

    fun headerStyle(wb: Workbook): CellStyle = wb.createCellStyle().apply {
        fillForegroundColor = IndexedColors.GREY_25_PERCENT.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
        setFont(wb.createFont().apply { bold = true })
        applyThinBorder()
        alignment = HorizontalAlignment.CENTER
    }

    fun totalStyle(wb: Workbook): CellStyle = wb.createCellStyle().apply {
        fillForegroundColor = IndexedColors.GREY_25_PERCENT.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
        setFont(wb.createFont().apply { bold = true })
        applyThinBorder()
        alignment = HorizontalAlignment.CENTER
    }

    fun dataStyle(wb: Workbook): CellStyle = wb.createCellStyle().apply {
        applyThinBorder()
        alignment = HorizontalAlignment.CENTER
    }

    fun labelStyle(wb: Workbook): CellStyle = wb.createCellStyle().apply {
        setFont(wb.createFont().apply { bold = true })
        alignment = HorizontalAlignment.RIGHT
    }

    fun sheetNameStyle(wb: Workbook): CellStyle = wb.createCellStyle().apply {
        setFont(wb.createFont().apply { bold = true })
        applyThinBorder()
        alignment = HorizontalAlignment.CENTER
    }

    private fun CellStyle.applyThinBorder() {
        borderTop    = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft   = BorderStyle.THIN
        borderRight  = BorderStyle.THIN
    }
}

