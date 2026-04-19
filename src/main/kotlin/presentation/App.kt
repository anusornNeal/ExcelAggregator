package presentation

import androidx.compose.desktop.ui.tooling.preview.Preview
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.material.MaterialTheme
import androidx.compose.material.lightColors
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.remember
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.Window
import androidx.compose.ui.window.application
import presentation.components.ButtonRow
import presentation.components.FileList
import presentation.components.ProcessButton
import presentation.components.StatusBox
import presentation.components.UiHeader
import presentation.contract.ViewContract
import java.io.File
import javax.swing.JFileChooser
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter

@Composable
@Preview
fun App() {
    val vm = remember { AggregatorViewModel() }
    EffectsHandler(vm)
    AppContent(state = vm.state, sendEvent = vm::sendEvent)
}

@Composable
private fun EffectsHandler(vm: AggregatorViewModel) {
    LaunchedEffect(vm.effects.size) {
        vm.effects.firstOrNull()?.let { effect ->
            when (effect) {
                ViewContract.Effect.OpenPickFilesDialog -> {
                    val files = pickExcelFiles()
                    vm.sendEvent(ViewContract.Event.AddFiles(files))
                }
                ViewContract.Effect.OpenPickFolderDialog -> {
                    pickFolder()?.let { folder ->
                        val excels = folder.walkTopDown()
                            .filter { it.isFile && it.extension.lowercase() in listOf("xlsx", "xls") }
                            .toList()
                        if (excels.isNotEmpty()) vm.sendEvent(ViewContract.Event.AddFiles(excels))
                    }
                }
                ViewContract.Effect.StartProcessing -> vm.processFiles()
            }
            vm.consumeEffect(effect)
        }
    }
}

@Composable
private fun AppContent(state: ViewContract.State, sendEvent: (ViewContract.Event) -> Unit) {
    MaterialTheme(colors = lightColors(primary = Color(0xFF1976D2))) {
        Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
            UiHeader()
            Spacer(Modifier.size(16.dp))
            ButtonRow(state = state, sendEvent = sendEvent)
            Spacer(Modifier.size(12.dp))
            if (state.selectedFiles.isNotEmpty()) FileList(state = state, sendEvent = sendEvent)
            Spacer(Modifier.size(12.dp))
            ProcessButton(state = state, sendEvent = sendEvent)
            if (state.statusMessage.isNotEmpty()) {
                Spacer(Modifier.size(8.dp))
                StatusBox(message = state.statusMessage)
            }
        }
    }
}

private fun pickExcelFiles(): List<File> {
    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName())
    val chooser = JFileChooser().apply {
        isMultiSelectionEnabled = true
        fileFilter = FileNameExtensionFilter("Excel Files (*.xlsx, *.xls)", "xlsx", "xls")
        dialogTitle = "เลือกไฟล์ Excel"
    }
    return if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
        chooser.selectedFiles.toList()
    } else {
        emptyList()
    }
}

private fun pickFolder(): File? {
    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName())
    val chooser = JFileChooser().apply {
        fileSelectionMode = JFileChooser.DIRECTORIES_ONLY
        dialogTitle = "เลือกโฟลเดอร์"
    }
    return if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) chooser.selectedFile else null
}

fun main() = application {
    Window(onCloseRequest = ::exitApplication, title = "Excel Aggregator") { App() }
}
