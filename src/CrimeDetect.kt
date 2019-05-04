
fun crimeDetection() {

    WB_OUT.createSheet("Загальна")

    println("Формування статистики по службах...")
    writeHeaderOnMainSheet()
    createRow(GlobalRow())
    writeDepartTabOnMainSheet()

    println("Формування статистики по статтях...")
    writeArticleTabOnMainSheet()

    println("Формування статистики служб по статтях...")
    WB_OUT.createSheet("Детальна")
    statByDepart.add(0, createTotalSumm(statByDepart))
    writeHeaderOnDetailSheet()
    createRow(DetailRow())
    writeTabOnDetailSheet()

    println("Визначення параметрів документа для друку...")
    setPrintSetup(WB_OUT)

    WB_OUT.setActiveSheet(1)

    closeOutXLSFile(WB_OUT)

    println("End of hard work!")
}
