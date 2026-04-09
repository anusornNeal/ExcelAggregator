package presentation.components

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import presentation.contract.ViewContract

@Composable
fun OutputRow(state: ViewContract.State, sendEvent: (ViewContract.Event) -> Unit) {
    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(8.dp)) {
        OutlinedTextField(
            value = state.outputPath,
            onValueChange = { sendEvent(ViewContract.Event.SetOutputPath(it)) },
            label = { Text("ไฟล์ผลลัพธ์ (.xlsx)") },
            modifier = Modifier.weight(1f),
            singleLine = true
        )
        Button(
            onClick = { sendEvent(ViewContract.Event.BrowseOutput) },
            colors = ButtonDefaults.buttonColors(containerColor = Color(0xFF5C6BC0))
        ) { Text("Browse", color = Color.White) }
    }
}

