import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.time.LocalDate

const val FIRST_GLOBAL_INDICATOR_ROW = 5
const val FIRST_DETAIL_INDICATOR_ROW = 7
const val GLOBAL_COLUMN_WIDTH = 4800
const val DETAIL_COLUMN_WIDTH = 4020
const val RESOURCE_FILE_NAME = "CrimeFiles.rs"
var INPUT_XLS_FILE_NAME = ""
var OUTPUT_XLS_FILE_NAME = ""
val DEPART_COLUMN = 0..3
val ARTICLE_COLUMN = 5..8
val CURRENT_YEAR = LocalDate.now().year
val WB_OUT = HSSFWorkbook()
