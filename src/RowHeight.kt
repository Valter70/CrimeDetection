interface RowHeight {
    val startRange: Int
    val endRange: Int
    fun createAllRowOnSheet()
}

class GlobalRow : RowHeight {
    override val startRange: Int
        get() = FIRST_GLOBAL_INDICATOR_ROW
    override val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size + statByGravity.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = WB_OUT.getSheetAt(0)
        createRowWithHeight(sheet, rowRange)
    }
}

class DetailRow : RowHeight {
    override val startRange: Int
        get() = FIRST_DETAIL_INDICATOR_ROW
    override val endRange: Int
        get() = startRange + statByArticle.size + statByArticle185.size
    private val rowRange = startRange..endRange

    override fun createAllRowOnSheet() {
        val sheet = WB_OUT.getSheetAt(1)
        createRowWithHeight(sheet, rowRange)
    }
}