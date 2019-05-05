import javafx.fxml.FXML
import javafx.scene.control.*
import javafx.stage.Stage
import org.eclipse.swt.SWT
import org.eclipse.swt.widgets.FileDialog
import org.eclipse.swt.widgets.MessageBox
import org.eclipse.swt.widgets.Shell
import java.io.*
import java.time.LocalDate

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
        val dlgOpen = FileDialog(Shell(), SWT.OPEN)
        dlgOpen.filterNames = arrayOf("Excel")
        dlgOpen.filterExtensions = arrayOf("*.xls")
        val fName = dlgOpen.open()
        if (fName != null)
            inputFileName.text = fName
    }

    @FXML
    private fun getOutputXLSFileName() {
        val dlgSave = FileDialog(Shell(), SWT.SAVE)
        dlgSave.filterNames = arrayOf("Excel")
        dlgSave.filterExtensions = arrayOf("*.xls")
        val fName = dlgSave.open()
        if (fName != null)
            outputFileName.text = fName
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
            val style = SWT.ICON_ERROR
            val mBox = MessageBox(Shell(SWT.APPLICATION_MODAL), style)
            mBox.text = "Помилка!"
            mBox.message = "Не коректні імена файлів!\n" + "Необхідні файли *.xls!   "
            mBox.open()
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
    }
}