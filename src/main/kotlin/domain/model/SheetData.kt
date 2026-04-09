package domain.model

import kotlin.math.abs

data class SheetData(
    val date: String,
    val branch: String,
    val rows: List<PriceRow>,
    val sourceFileName: String? = null,
    val sourceFileIndex: Int? = null,
    val sourceSheetIndex: Int? = null,
    val sourceSheetName: String? = null
)

fun SheetData.toPriceMap(): Map<Double?, PriceRow> =
    rows.filter { !it.isTotal && !it.isGP }.associateBy { it.price }

fun SheetData.getTotalRow(): PriceRow? =
    rows.firstOrNull { it.isTotal }

fun SheetData.getGpRows(): List<PriceRow> =
    rows.filter { it.isGP }

fun SheetData.getGpValueForTarget(target: Double, totalValue: Double?): Pair<Double?, Double?> {
    val percent = getGpRows().firstOrNull { it.gpPercent != null && abs(it.gpPercent - target) < 0.000001 }?.gpPercent
    val value   = if (percent != null && totalValue != null) percent * totalValue else null
    return percent to value
}

