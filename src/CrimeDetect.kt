import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun main(args: Array<String>) {

    val wbOut = HSSFWorkbook()
    val sheetGlobal = wbOut.createSheet("Загальна")

    writeHeaderOnMainSheet(wbOut)
    var endOfRange = FIRST_GLOBAL_INDICATOR_ROW + statByArticle.size + statByArticle185.size + statByGravity.size
    setRowHeightOnMainSheet(sheetGlobal, FIRST_GLOBAL_INDICATOR_ROW..endOfRange)
    writeDepartTabOnMainSheet(wbOut)
    writeArticleTabOnMainSheet(wbOut)

    val sheetDetail = wbOut.createSheet("Детальна")
    writeHeaderOnDetailSheet(wbOut)
    wbOut.setActiveSheet(1)

    closeOutXLSFile(wbOut)

}
