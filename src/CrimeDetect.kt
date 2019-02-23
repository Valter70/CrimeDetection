import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun main(args: Array<String>) {

    val crimeList = createCrimeListFromXLS()
    val crimeOfArticle185 = crimeList.filter { it.article.substringBefore(' ') == "185" }

    val statByDepart = createStatByDepart(crimeList)
    val statByArticle = createStatByArticle(crimeList)
    val statByArticle185 = createStatByArticle185(crimeOfArticle185)

    val wbOut = HSSFWorkbook()
    val sheetOut = wbOut.createSheet("Загальна")

    writeHeaderOnMainSheet(wbOut, sheetOut)
    val endOfRange = FIRST_INDICATOR_ROW + statByArticle.size + statByArticle185.size
    setRowHeightOnMainSheet(sheetOut, FIRST_INDICATOR_ROW..endOfRange)
    writeDepartTabOnMainSheet(wbOut, statByDepart)
    writeArticleTabOnMainSheet(wbOut, statByArticle, statByArticle185)

    closeOutXLSFile(wbOut)

}
