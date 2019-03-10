import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun writeArticleTabOnMainSheet(wb: HSSFWorkbook) {
    writeGravityIndicator(wb)
    val startArticleRow = FIRST_GLOBAL_INDICATOR_ROW + 2
    writeArticleIndicator(wb, startArticleRow)

    val totalArticleTabSize = statByArticle.size + statByArticle185.size + 2
    val startRangRow = FIRST_GLOBAL_INDICATOR_ROW
    val endRangeRow = startRangRow + totalArticleTabSize - 1
    val rangeRows = startRangRow..endRangeRow
    setStyleForIndicatorRow(wb, rangeRows, ARTICLE_COLUMN, 0)

    writeArticleTotalSumm(wb)
}

fun writeArticleIndicator(wb: HSSFWorkbook, startRow: Int) {
    var part185 = 0
    for(index in 0..(statByArticle.size - 1)) {
        val row = wb.getSheetAt(0).getRow(startRow + index + part185)
        writeIndicatorToTab(row, 5, statByArticle[index])
        val articleName = statByArticle[index].nameIndicator.substring(3)
        if(articleName == "185") {
            part185 = 5
            writeArticle185Indicator(wb, startRow + index + 1)
        }
    }

}

fun writeGravityIndicator(wb: HSSFWorkbook) {
    for(index in 0..1) {
        val row = wb.getSheetAt(0).getRow(FIRST_GLOBAL_INDICATOR_ROW + index)
        writeIndicatorToTab(row, 5, statByGravity[index + 2])
    }
}

fun writeArticle185Indicator(wb: HSSFWorkbook, startIndex: Int) {
    for(index in 0..(statByArticle185.size - 1)) {
        val row = wb.getSheetAt(0).getRow(startIndex + index)
        writeIndicatorToTab(row, 5, statByArticle185[index])
    }

}

fun createStatByArticle() : MutableList<StatForOutXLS> {
    val listOfArticle = crimeList.map { it.article.substringBefore(' ') }.toSortedSet()
    val statByArticle: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))

    for(rec in listOfArticle) {
        val currentYear = crimeList.count { it.isCurrentYear && it.article.substringBefore(' ') == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.article.substringBefore(' ') == rec }
        statByArticle.add(StatForOutXLS("ст.$rec", pastYear, currentYear))
    }
    statByArticle.removeAt(0)
    return statByArticle
}

fun writeArticleTotalSumm(wb: HSSFWorkbook) {
    val endRowIndex = FIRST_GLOBAL_INDICATOR_ROW + 2 + statByArticle.size + 5
    val totalSummRow = wb.getSheetAt(0).getRow(endRowIndex)
    val totalSumm = createTotalSumm(statByArticle)
    writeIndicatorToTab(totalSummRow, ARTICLE_COLUMN.first, totalSumm)
    setStyleForTotalSummCells(wb, ARTICLE_COLUMN, totalSummRow)

}

fun createStatByArticle185() : MutableList<StatForOutXLS> {
    val statByArticle185: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))

    for(rec in 1..5) {
        val part185 = "Ч.$rec"
        val currentYear = crimeOfArticle185.count { it.isCurrentYear && it.article.substringAfter(' ') == part185 }
        val pastYear = crimeOfArticle185.count { !it.isCurrentYear && it.article.substringAfter(' ') == part185 }
        statByArticle185.add(StatForOutXLS("-" + part185.toLowerCase(), pastYear, currentYear))
    }
    statByArticle185.removeAt(0)
    return statByArticle185
}
