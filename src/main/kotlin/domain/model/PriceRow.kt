package domain.model

data class PriceRow(
    val price: Double?,
    val qty: Double?,
    val value: Double?,
    val isTotal: Boolean = false,
    val isGP: Boolean = false,
    val gpPercent: Double? = null
)

