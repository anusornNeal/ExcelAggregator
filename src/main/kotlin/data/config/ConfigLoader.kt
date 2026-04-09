package data.config

import domain.model.AggregatorConfig
import java.io.File
import java.util.Properties

object ConfigLoader {
    private const val FILENAME = "aggregator.properties"

    fun load(): AggregatorConfig {
        val propsFile = File(FILENAME)
        if (!propsFile.exists()) return AggregatorConfig()
        val p = Properties().apply { load(propsFile.inputStream()) }
        fun csv(key: String, default: List<String>) = p.getProperty(key)?.split(',')?.map { it.trim() } ?: default
        return AggregatorConfig(
            dateKeywords   = csv("date.keywords", listOf("วันที่", "Date")),
            branchKeywords = csv("branch.keywords", listOf("สาขา", "Branch")),
            headerKeyword  = p.getProperty("header.keyword", "ราคา"),
            priceColumn    = p.getProperty("column.price", "7").toIntOrNull() ?: 7,
            qtyColumn      = p.getProperty("column.qty", "8").toIntOrNull() ?: 8,
            valueColumn    = p.getProperty("column.value", "9").toIntOrNull() ?: 9,
            totalKeywords  = csv("total.keywords", listOf("Total", "รวม")),
            gpKeyword      = p.getProperty("gp.keyword", "GP"),
            gpTargets      = p.getProperty("gp.targets", "0.12,0.16,0.2").split(',').mapNotNull { it.trim().toDoubleOrNull() }
        )
    }
}

