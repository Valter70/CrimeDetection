import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun main(args: Array<String>) {

    val wbOut = HSSFWorkbook()
    val sheetGlobal = wbOut.createSheet("Загальна")

    writeHeaderOnMainSheet(wbOut)
    val endOfRange = FIRST_INDICATOR_ROW + statByArticle.size + statByArticle185.size
    setRowHeightOnMainSheet(sheetGlobal, FIRST_INDICATOR_ROW..endOfRange)
    writeDepartTabOnMainSheet(wbOut)
    writeArticleTabOnMainSheet(wbOut)

    val sheetDetail = wbOut.createSheet("Детальна")
    writeHeaderOnDetailSheet(wbOut)

    closeOutXLSFile(wbOut)

}
