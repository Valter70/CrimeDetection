import javafx.fxml.FXML
import javafx.scene.control.*
import javafx.stage.Stage
import java.io.*
import java.time.LocalDate
import javax.swing.JFileChooser
import javax.swing.JOptionPane
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter

class MainController {

    @FXML
    lateinit var btnOK: Button

    @FXML
    lateinit var btnCancel: Button

    @FXML
    lateinit var inputFileDialog: Button

    @FXML
    lateinit var outputFileDialog: Button

    @FXML
    lateinit var inLabel: Label

    @FXML
    lateinit var monthLabel: Label

    @FXML
    lateinit var outLabel: Label

    @FXML
    lateinit var inputFileName: TextField

    @FXML
    lateinit var outputFileName: TextField

    @FXML
    lateinit var months: ComboBox<String>

    @FXML
    private fun getInputXLSFileName() {
        val fileOpen = JFileChooser(".")
        fileOpen.fileFilter = FileNameExtensionFilter("Excel", "xls", "xlsx")
        val ret = fileOpen.showDialog(null, "Открыть файл")
        if (ret == JFileChooser.APPROVE_OPTION) {
            inputFileName.text = fileOpen.selectedFile.absolutePath
            if (outputFileName.text == "") {
                outputFileName.text = inputFileName.text.substringBeforeLast("\\") + "\\" + setDefaultOutputFileName(inputFileName.text)
            }
        }
    }

    @FXML
    private fun getOutputXLSFileName() {
        val fileSave = JFileChooser(".")
        fileSave.fileFilter = FileNameExtensionFilter("Excel", "xls", "xlsx")
        val ret = fileSave.showDialog(null, "Сохранить в файл")
        if (ret == JFileChooser.APPROVE_OPTION) {
            outputFileName.text = fileSave.selectedFile.absolutePath
        }
    }

    @FXML
    private fun actCancel() {
        closeDialog(btnCancel)
        System.exit(1)
    }

    @FXML
    private fun actOK() {
        INPUT_XLS_FILE_NAME = inputFileName.text
        OUTPUT_XLS_FILE_NAME = outputFileName.text

        if(!isValidFileName()) {
            JOptionPane.showMessageDialog(null, "Не коректні імена файлів!\n" + "Необхідні файли *.xls!   ", "Помилка!", JOptionPane.ERROR_MESSAGE)
        } else {
            closeDialog(btnCancel)
            val fileWriter = FileWriter(RESOURCE_FILE_NAME, false)
            fileWriter.write(INPUT_XLS_FILE_NAME + "\n")
            fileWriter.write(OUTPUT_XLS_FILE_NAME + "\n")
            fileWriter.close()
        }
    }

    private fun isValidFileName() : Boolean =
        INPUT_XLS_FILE_NAME != "" && OUTPUT_XLS_FILE_NAME != "" &&
        INPUT_XLS_FILE_NAME.substringAfter(".") == "xls" && OUTPUT_XLS_FILE_NAME.substringAfter(".") == "xls"


    private fun closeDialog(btn: Button) {
        val stage = btn.scene.window as Stage
        stage.close()
    }

    fun initialize() {
        months.items.addAll("Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень")
        months.selectionModel.select(LocalDate.now().monthValue - 1)
        val rsFile = File(RESOURCE_FILE_NAME)
        if (rsFile.exists()) {
            val filesName = rsFile.readLines()
            inputFileName.text = filesName[0]
            outputFileName.text = filesName[1]
        }

        localizeDialog()
    }

    private fun localizeDialog() {
        UIManager.put("FileChooser.saveButtonText", "Сохранить");
        UIManager.put("FileChooser.saveButtonToolTipText", "Сохранить выбранный файл");
        UIManager.put("FileChooser.openButtonText", "Открыть");
        UIManager.put("FileChooser.openButtonToolTipText", "Открыть выбранный файл");
        UIManager.put("FileChooser.cancelButtonText", "Отмена");
        UIManager.put("FileChooser.cancelButtonToolTipText", "Отменить открытие файла");
        UIManager.put("FileChooser.fileNameLabelText", "Имя файла");
        UIManager.put("FileChooser.filesOfTypeLabelText", "Типы файлов");
        UIManager.put("FileChooser.lookInLabelText", "Директория");
        UIManager.put("FileChooser.saveInLabelText", "Сохранить в директории");
        UIManager.put("FileChooser.folderNameLabelText", "Путь директории");
    }

    private fun setDefaultOutputFileName(path: String) : String {
        return path.substringAfterLast("\\").substringBefore(".") + "-outTab." + path.substringAfter(".")
    }
}
