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

fun SheetData.getDetailRows(): List<PriceRow> =
    rows.filter { !it.isTotal && !it.isGP }

fun SheetData.getTotalRow(): PriceRow? =
    rows.firstOrNull { it.isTotal }

fun SheetData.getGpRows(): List<PriceRow> =
    rows.filter { it.isGP }

fun SheetData.getGpTotalValue(): Double? =
    getGpRows().mapNotNull { it.value }.takeIf { it.isNotEmpty() }?.sum()

fun SheetData.getGpTotalsByPercent(): Map<Double, Double> =
    getGpRows()
        .filter { it.gpPercent != null && it.value != null }
        .groupBy { it.gpPercent!! }
        .mapValues { (_, rows) -> rows.sumOf { it.value ?: 0.0 } }

fun SheetData.getGpSummaryText(): String =
    getGpRows()
        .mapNotNull { row ->
            val percent = row.gpPercent?.let(::formatGpPercentLabel)
            val value = row.value?.cleanNumber()
            when {
                percent != null && value != null -> "$percent = $value"
                percent != null -> percent
                value != null -> value
                else -> null
            }
        }
        .ifEmpty { listOf("-") }
        .joinToString(", ")

fun Double.cleanNumber(): String =
    if (this % 1.0 == 0.0) toLong().toString() else toString()

fun formatGpPercentLabel(value: Double): String {
    val displayValue = if (kotlin.math.abs(value) <= 1.0) value * 100 else value
    val formatted = if (displayValue % 1.0 == 0.0) displayValue.toLong().toString() else displayValue.toString()
    return "GP $formatted%"
}

fun SheetData.getGpValueForTarget(target: Double, totalValue: Double?): Pair<Double?, Double?> {
    val percent = getGpRows().firstOrNull { it.gpPercent != null && abs(it.gpPercent - target) < 0.000001 }?.gpPercent
    val value   = if (percent != null && totalValue != null) percent * totalValue else null
    return percent to value
}

