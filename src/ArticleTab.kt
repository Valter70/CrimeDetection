import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun writeArticleTabOnMainSheet(wb: HSSFWorkbook) {
    writeArticleIndicator(wb)

    val totalArticleTabSize = statByArticle.size + statByArticle185.size
    val rangeRows = FIRST_GLOBAL_INDICATOR_ROW..(FIRST_GLOBAL_INDICATOR_ROW + totalArticleTabSize - 1)
    setStyleForIndicatorRow(wb, rangeRows, ARTICLE_COLUMN)

    writeArticleTotalSumm(wb)
}

fun writeArticleIndicator(wb: HSSFWorkbook) {
    var part185 = 0
    for(index in 0..(statByArticle.size - 1)) {
        val row = wb.getSheetAt(0).getRow(FIRST_GLOBAL_INDICATOR_ROW + index + part185)
        writeIndicatorToTab(row, 5, statByArticle[index])
        val articleName = statByArticle[index].nameIndicator.substring(3)
        if(articleName == "185") {
            part185 = 5
            writeArticle185Indicator(wb, FIRST_GLOBAL_INDICATOR_ROW + index + 1)
        }
    }

}

fun writeArticle185Indicator(wb: HSSFWorkbook, startIndex: Int) {
    for(index in 0..(statByArticle185.size - 1)) {
        val row = wb.getSheetAt(0).getRow(startIndex + index)
        writeIndicatorToTab(row, 5, statByArticle185[index])
    }

}

fun createStatByArticle() : MutableList<StatForOutXLS> {
    val listOfArticle = crimeList.map { it.article.substringBefore(' ') }.toSet()
    val statByArticle: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))
    for(rec in GravityOfCrime.values()) {
        val gravity = rec.statName.toLowerCase()
        val currentYear = crimeList.count { it.isCurrentYear && it.gravity == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.gravity == rec }
        statByArticle.add(StatForOutXLS(gravity, pastYear, currentYear))
    }

    for(rec in listOfArticle) {
        val currentYear = crimeList.count { it.isCurrentYear && it.article.substringBefore(' ') == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.article.substringBefore(' ') == rec }
        statByArticle.add(StatForOutXLS("ст.$rec", pastYear, currentYear))
    }
    statByArticle.removeAt(2)
    statByArticle.removeAt(1)
    statByArticle.removeAt(0)
    return statByArticle
}

fun writeArticleTotalSumm(wb: HSSFWorkbook) {
    val endRowIndex = FIRST_GLOBAL_INDICATOR_ROW + statByArticle.size + 5
    val totalSummRow = wb.getSheetAt(0).getRow(endRowIndex)
    statByArticle.removeAt(1)
    statByArticle.removeAt(0)
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
