package presentation.contract

import java.io.File

sealed class ViewContract {
    data class State(
        val selectedFiles: List<File> = emptyList(),
        val statusMessage: String = "",
        val isProcessing: Boolean = false
    ) : ViewContract()

    sealed class Event : ViewContract() {
        object PickFiles    : Event()
        object PickFolder   : Event()
        data class AddFiles(val files: List<File>) : Event()
        data class RemoveFile(val file: File) : Event()
        data class DiagnoseFile(val file: File) : Event()
        object ClearFiles : Event()
        object Process : Event()
    }

    sealed class Effect : ViewContract() {
        object OpenPickFilesDialog  : Effect()
        object OpenPickFolderDialog : Effect()
        object StartProcessing      : Effect()
    }
}

