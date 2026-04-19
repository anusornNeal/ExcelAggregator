package presentation

import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateListOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import domain.usecase.AggregateInvoicesUseCase
import domain.usecase.BuildStatusMessageUseCase
import domain.usecase.DiagnoseFileUseCase
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import presentation.contract.ViewContract
import java.io.File

class AggregatorViewModel(
    private val aggregateInvoices: AggregateInvoicesUseCase = AggregateInvoicesUseCase(),
    private val diagnoseFile: DiagnoseFileUseCase = DiagnoseFileUseCase(),
    private val buildStatusMessage: BuildStatusMessageUseCase = BuildStatusMessageUseCase()
) {
    var state by mutableStateOf(ViewContract.State())
        private set

    private val _effects = mutableStateListOf<ViewContract.Effect>()
    val effects: List<ViewContract.Effect> get() = _effects

    private val eventBus = MutableSharedFlow<ViewContract.Event>(extraBufferCapacity = 64)
    private val scope = CoroutineScope(Dispatchers.Default + SupervisorJob())

    init {
        scope.launch { eventBus.collect { onEvent(it) } }
    }

    fun sendEvent(event: ViewContract.Event) {
        scope.launch { eventBus.emit(event) }
    }

    fun consumeEffect(effect: ViewContract.Effect) {
        _effects.remove(effect)
    }

    private fun onEvent(event: ViewContract.Event) {
        when (event) {
            ViewContract.Event.PickFiles -> _effects.add(ViewContract.Effect.OpenPickFilesDialog)
            ViewContract.Event.PickFolder -> _effects.add(ViewContract.Effect.OpenPickFolderDialog)
            is ViewContract.Event.AddFiles -> addFiles(event.files)
            is ViewContract.Event.RemoveFile -> removeFile(event.file)
            is ViewContract.Event.DiagnoseFile -> runDiagnose(event.file)
            ViewContract.Event.ClearFiles -> state = state.copy(selectedFiles = emptyList(), statusMessage = "")
            ViewContract.Event.Process -> startProcessing()
        }
    }

    private fun addFiles(files: List<File>) {
        if (files.isEmpty()) return
        val merged = (state.selectedFiles + files).distinctBy { it.absolutePath }
        val outputFolder = merged.firstOrNull()?.parentFile?.absolutePath.orEmpty()
        state = state.copy(
            selectedFiles = merged,
            statusMessage = if (outputFolder.isBlank()) {
                "Selected ${merged.size} files"
            } else {
                "Selected ${merged.size} files. Output will be saved in $outputFolder"
            }
        )
    }

    private fun removeFile(file: File) {
        val remain = state.selectedFiles.filter { it != file }
        val outputFolder = remain.firstOrNull()?.parentFile?.absolutePath
        state = state.copy(
            selectedFiles = remain,
            statusMessage = when {
                remain.isEmpty() -> ""
                outputFolder.isNullOrBlank() -> "Selected ${remain.size} files"
                else -> "Selected ${remain.size} files. Output will be saved in $outputFolder"
            }
        )
    }

    private fun startProcessing() {
        if (state.selectedFiles.isEmpty()) {
            state = state.copy(statusMessage = "Please select Excel files first")
            return
        }
        state = state.copy(isProcessing = true, statusMessage = "Processing files...")
        _effects.add(ViewContract.Effect.StartProcessing)
    }

    suspend fun processFiles() {
        try {
            val result = withContext(Dispatchers.IO) {
                aggregateInvoices.execute(state.selectedFiles)
            }
            val status = buildStatusMessage.execute(result)
            state = state.copy(selectedFiles = emptyList(), isProcessing = false, statusMessage = status)
        } catch (e: Exception) {
            state = state.copy(isProcessing = false, statusMessage = "Error: ${e.message}")
        }
    }

    private fun runDiagnose(file: File) {
        scope.launch {
            val result = withContext(Dispatchers.IO) { diagnoseFile.execute(file) }
            runCatching { java.awt.Desktop.getDesktop().open(File(result.reportPath)) }
            state = state.copy(statusMessage = "Diagnostic written: ${result.reportPath}")
        }
    }
}
