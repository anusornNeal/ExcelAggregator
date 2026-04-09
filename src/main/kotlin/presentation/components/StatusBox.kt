package presentation.components

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp

@Composable
fun StatusBox(message: String) {
    val bgColor = when {
        message.startsWith("✅") -> Color(0xFFE8F5E9)
        message.startsWith("⚠️") || message.startsWith("❌") -> Color(0xFFFFEBEE)
        else -> Color(0xFFF5F5F5)
    }
    Box(modifier = Modifier.fillMaxWidth().background(bgColor, RoundedCornerShape(4.dp)).padding(12.dp)) {
        Text(message, fontSize = 13.sp)
    }
}

