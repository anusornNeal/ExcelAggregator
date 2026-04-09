package presentation.components

import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import presentation.contract.ViewContract

@Composable
fun ProcessButton(state: ViewContract.State, sendEvent: (ViewContract.Event) -> Unit) {
    Button(
        onClick = { sendEvent(ViewContract.Event.Process) },
        enabled = !state.isProcessing && state.selectedFiles.isNotEmpty() && state.outputPath.isNotBlank(),
        modifier = Modifier.fillMaxWidth().height(48.dp),
        colors = ButtonDefaults.buttonColors(containerColor = Color(0xFFFFA000))
    ) {
        Text(
            text = if (state.isProcessing) "กำลังประมวลผล..." else "▶  รวมไฟล์",
            color = Color.White, fontWeight = FontWeight.Bold, fontSize = 16.sp
        )
    }
}

