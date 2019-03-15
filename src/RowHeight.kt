import org.apache.poi.hssf.usermodel.HSSFSheet

open class RowHeight {
    open fun createAllRowOnSheet() {}
    protected fun createRowWithHeight(sheet: HSSFSheet, range: IntRange) {
        for(row in range) {
            sheet.createRow(row).heightInPoints = 26F
        }
    }
}

class GlobalRow : RowHeight() {
    private val startRange: Int
        get() = FIRST_GLOBAL_INDICATOR_ROW
    private val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size + statByGravity.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = WB_OUT.getSheetAt(0)
        createRowWithHeight(sheet, rowRange)
    }
}

class DetailRow : RowHeight() {
    private val startRange: Int
        get() = FIRST_DETAIL_INDICATOR_ROW
    private val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = WB_OUT.getSheetAt(1)
        createRowWithHeight(sheet, rowRange)
    }
}