package presentation.components

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import presentation.contract.ViewContract

@Composable
fun ButtonRow(state: ViewContract.State, sendEvent: (ViewContract.Event) -> Unit) {
    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
        Button(
            onClick = { sendEvent(ViewContract.Event.PickFiles) },
            colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF1976D2))
        ) { Text("เลือกไฟล์ (Excel)", color = Color.White) }

        Button(
            onClick = { sendEvent(ViewContract.Event.PickFolder) },
            colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF388E3C))
        ) { Text("เลือกโฟลเดอร์", color = Color.White) }

        if (state.selectedFiles.isNotEmpty()) {
            Button(
                onClick = { sendEvent(ViewContract.Event.ClearFiles) },
                colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFD32F2F))
            ) { Text("ล้างรายการ", color = Color.White) }
        }
    }
}

