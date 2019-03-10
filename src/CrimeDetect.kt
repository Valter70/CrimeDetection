import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun main(args: Array<String>) {

    INPUT_XLS_FILE_NAME = args[0]
    OUTPUT_XLS_FILE_NAME = args[1]

    val wbOut = HSSFWorkbook()
    val sheetGlobal = wbOut.createSheet("Загальна")

    println("Формування статистики по службах...")
    writeHeaderOnMainSheet(wbOut)
    var endOfRange = FIRST_GLOBAL_INDICATOR_ROW + statByArticle.size + statByArticle185.size + statByGravity.size
    setRowHeightOnMainSheet(sheetGlobal, FIRST_GLOBAL_INDICATOR_ROW..endOfRange)
    writeDepartTabOnMainSheet(wbOut)

    println("Формування статистики по статтях...")
    writeArticleTabOnMainSheet(wbOut)

    println("Формування статистики служб по статтях...")
    val sheetDetail = wbOut.createSheet("Детальна")
    statByDepart.add(0, createTotalSumm(statByDepart))
    writeHeaderOnDetailSheet(wbOut)
    endOfRange = FIRST_DETAIL_INDICATOR_ROW + statByArticle.size + statByArticle185.size
    setRowHeightOnMainSheet(sheetDetail, FIRST_DETAIL_INDICATOR_ROW..endOfRange)
    writeTabOnDetailSheet(wbOut)

    println("Визначення параметрів документа для друку...")
    setPrintSetup(wbOut)

    wbOut.setActiveSheet(1)

    closeOutXLSFile(wbOut)

    println("End of hard work!")
}
