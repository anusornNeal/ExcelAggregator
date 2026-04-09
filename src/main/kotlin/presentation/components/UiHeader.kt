package presentation.components

import androidx.compose.foundation.layout.*
import androidx.compose.material.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp

@Composable
fun UiHeader() {
    Column {
        Text("Excel Aggregator", fontSize = 22.sp, fontWeight = FontWeight.Bold, color = Color(0xFF1976D2))
        Spacer(Modifier.height(4.dp))
        Text("รวมข้อมูลสรุปราคาจากหลายไฟล์ Invoice เข้าไฟล์เดียว", fontSize = 13.sp, color = Color.Gray)
    }
}

