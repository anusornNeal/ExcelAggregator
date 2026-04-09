package domain.model

/** Simple configuration for the aggregator. Values come from aggregator.properties (working dir) or defaults. */
data class AggregatorConfig(
    val dateKeywords: List<String> = listOf("วันที่", "Date"),
    val branchKeywords: List<String> = listOf("สาขา", "Branch"),
    val headerKeyword: String = "ราคา",
    val priceColumn: Int = 7,
    val qtyColumn: Int = 8,
    val valueColumn: Int = 9,
    val totalKeywords: List<String> = listOf("Total", "รวม"),
    val gpKeyword: String = "GP",
    val gpTargets: List<Double> = listOf(0.12, 0.16, 0.2),
    // labels used in output
    val labelFilePrefix: String = "ไฟล์ที่ ",
    val labelDate: String = "วันที่ :",
    val labelBranch: String = "สาขา :",
    val labelPrice: String = "ราคา",
    val labelQty: String = "จำนวนตัว",
    val labelValue: String = "มูลค่า",
    val labelTotal: String = "Total",
    val labelGp: String = "GP"
)

