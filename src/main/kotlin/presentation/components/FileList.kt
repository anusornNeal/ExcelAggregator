package presentation.components

import androidx.compose.foundation.border
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.*
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Notifications
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import presentation.contract.ViewContract

@Composable
fun ColumnScope.FileList(state: ViewContract.State, sendEvent: (ViewContract.Event) -> Unit) {
    Text("ไฟล์ที่เลือก (${state.selectedFiles.size} ไฟล์):", fontSize = 14.sp)
    Spacer(Modifier.height(4.dp))
    Box(
        modifier = Modifier.weight(1f).fillMaxWidth()
            .border(1.dp, Color.LightGray, RoundedCornerShape(4.dp))
    ) {
        LazyColumn(modifier = Modifier.fillMaxSize().padding(4.dp)) {
            items(state.selectedFiles) { file ->
                Row(
                    modifier = Modifier.fillMaxWidth().padding(vertical = 2.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Text(file.name, modifier = Modifier.weight(1f), fontSize = 13.sp)
                    Text(
                        text = file.parent.take(40) + if (file.parent.length > 40) "..." else "",
                        fontSize = 11.sp, color = Color.Gray, modifier = Modifier.width(220.dp)
                    )
                    IconButton(onClick = { sendEvent(ViewContract.Event.RemoveFile(file)) }, modifier = Modifier.size(24.dp)) {
                        Icon(Icons.Default.Close, contentDescription = "ลบ", tint = Color.Red)
                    }
                    IconButton(onClick = { sendEvent(ViewContract.Event.DiagnoseFile(file)) }, modifier = Modifier.size(24.dp)) {
                        Icon(Icons.Default.Notifications, contentDescription = "Diagnose", tint = Color(0xFF1976D2))
                    }
                }
                Divider(color = Color(0xFFEEEEEE))
            }
        }
    }
}

