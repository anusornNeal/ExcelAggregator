package data.reader

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import java.text.SimpleDateFormat
import kotlin.math.floor

/** Low-level Excel cell helpers. */
object CellUtils {

    fun getCellString(cell: Cell?): String {
        if (cell == null) return ""
        return when (cell.cellType) {
            CellType.STRING  -> cell.stringCellValue.trim()
            CellType.NUMERIC -> if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat("d/M/yyyy").format(cell.dateCellValue)
            } else {
                val d = cell.numericCellValue
                if (d == floor(d)) d.toLong().toString() else d.toString()
            }
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            CellType.FORMULA -> runCatching { cell.stringCellValue.trim() }.getOrElse {
                runCatching {
                    val d = cell.numericCellValue
                    if (d == floor(d)) d.toLong().toString() else d.toString()
                }.getOrDefault("")
            }
            else -> ""
        }
    }

    fun parsePercent(raw: String): Double? =
        raw.trim().replace("%", "").toDoubleOrNull()

    fun setCellDoubleOrBlank(cell: Cell, value: Double?) {
        if (value == null) cell.setBlank() else cell.setCellValue(value)
    }
}

