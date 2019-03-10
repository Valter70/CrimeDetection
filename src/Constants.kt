import java.time.LocalDate

const val FIRST_GLOBAL_INDICATOR_ROW = 5
const val FIRST_DETAIL_INDICATOR_ROW = 7
const val GLOBAL_COLUMN_WIDTH = 4800
const val DETAIL_COLUMN_WIDTH = 4020
var INPUT_XLS_FILE_NAME = ""
var OUTPUT_XLS_FILE_NAME = ""
val DEPART_COLUMN = 0..3
val ARTICLE_COLUMN = 5..8
val CURRENT_YEAR = LocalDate.now().year

