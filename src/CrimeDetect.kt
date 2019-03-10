import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun main(args: Array<String>) {

    INPUT_XLS_FILE_NAME = args[0]
    OUTPUT_XLS_FILE_NAME = args[1]

    val wbOut = HSSFWorkbook()
    wbOut.createSheet("Загальна")

    println("Формування статистики по службах...")
    writeHeaderOnMainSheet(wbOut)
    GlobalRow(wbOut).createAllRowOnSheet()
    writeDepartTabOnMainSheet(wbOut)

    println("Формування статистики по статтях...")
    writeArticleTabOnMainSheet(wbOut)

    println("Формування статистики служб по статтях...")
    wbOut.createSheet("Детальна")
    statByDepart.add(0, createTotalSumm(statByDepart))
    writeHeaderOnDetailSheet(wbOut)
    DetailRow(wbOut).createAllRowOnSheet()
    writeTabOnDetailSheet(wbOut)

    println("Визначення параметрів документа для друку...")
    setPrintSetup(wbOut)

    wbOut.setActiveSheet(1)

    closeOutXLSFile(wbOut)

    println("End of hard work!")
}
