import org.apache.poi.hssf.usermodel.HSSFWorkbook

interface RowHeight {
    val startRange: Int
    val endRange: Int
    fun createAllRowOnSheet()
}

class GlobalRow(private val wb: HSSFWorkbook) : RowHeight {
    override val startRange: Int
        get() = FIRST_GLOBAL_INDICATOR_ROW
    override val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size + statByGravity.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = wb.getSheetAt(0)
        createRowWithHeight(sheet, rowRange)
    }
}

class DetailRow(private val wb: HSSFWorkbook) : RowHeight {
    override val startRange: Int
        get() = FIRST_DETAIL_INDICATOR_ROW
    override val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = wb.getSheetAt(1)
        createRowWithHeight(sheet, rowRange)
    }
}