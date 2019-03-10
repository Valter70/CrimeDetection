import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalDate

fun getArticleOfCriminalStat(currentRow: HSSFRow) : String {
    val articleCell = headerNames().indexOf("Квал.злоч.КК 2001")
    return currentRow.getCell(articleCell).stringCellValue.substring(3)
}

fun getDepartOfCriminalStat(currentRow: HSSFRow) : String {
    val departCell = headerNames().indexOf("17.Особа виявлена службою")
    return currentRow.getCell(departCell).stringCellValue
}
fun getGravityOfCriminalStat(currentRow: HSSFRow) : String {
    val gravityCell = headerNames().indexOf("Ф1 14.Квал.злоч.-тяжкiсть")
    return currentRow.getCell(gravityCell).stringCellValue
}

fun getNumberOfCriminalStat(currentRow: HSSFRow) : String {
    var numKP = ""
    for(i in 0..2) {
        val partOfNumber = currentRow.getCell(i).numericCellValue.toInt().toString()
        numKP += partOfNumber
    }
    numKP += createFullNumber(currentRow.getCell(3).numericCellValue.toInt().toString())
    return numKP
}

fun createFullNumber(number: String) : String {
    var result = ""
    for(i in 1..(7 - number.length))
        result += "0"
    return result + number
}


fun createCrimeCase(rowCP: HSSFRow) : CrimeCaseF2ForStat {
    val number = getNumberOfCriminalStat(rowCP)
    val article = getArticleOfCriminalStat(rowCP)
    val depart = getCodeDepart(getDepartOfCriminalStat(rowCP))
    val gravity = getCodeGravity(getGravityOfCriminalStat(rowCP))

    return CrimeCaseF2ForStat(number, article, depart, gravity)
}

fun createCrimeListFromXLS() : List<CrimeCaseF2ForStat> {
    val wbStat = HSSFWorkbook(FileInputStream(INPUT_XLS_FILE_NAME))
    val sheetStat = wbStat.getSheetAt(0)

    removeAllMergedRegions(wbStat)

    val crimeList: MutableList<CrimeCaseF2ForStat> = mutableListOf(CrimeCaseF2ForStat("", "", Department.VKP, GravityOfCrime.T1, true))
    var currentRecord = 5
    while(currentRecord <= sheetStat.lastRowNum) {
        val firstNumberValue = sheetStat.getRow(currentRecord).getCell(0).numericCellValue
        if(isCorrectNumber(firstNumberValue)) {
            val crimeCase = createCrimeCase(sheetStat.getRow(currentRecord))
            crimeList.add(crimeCase)
        }
        currentRecord++
    }

    crimeList.removeAt(0)
    wbStat.close()
    return crimeList
}

fun removeAllMergedRegions(wb: HSSFWorkbook) = with(wb.getSheetAt(0)) {
    for(index in 0..numMergedRegions)
        removeMergedRegion(index)
}

fun isCorrectNumber(number: Double) = number != 0.0

fun closeOutXLSFile(wb: HSSFWorkbook) {
    val fosE = FileOutputStream(OUTPUT_XLS_FILE_NAME)
    wb.write(fosE)
    fosE.close()
}

fun writeIndicatorToTab(row: HSSFRow, startCell: Int, indicator: StatForOutXLS) {
    row.createCell(startCell).setCellValue(indicator.nameIndicator)
    row.createCell(startCell + 1).setCellValue(indicator.total.toDouble())
    writeYearIndicatorToTab(row, startCell + 2, indicator)
}

fun writeYearIndicatorToTab(row: HSSFRow, startCell: Int, indicator: StatForOutXLS) {
    row.createCell(startCell)
    if(indicator.pastYears != 0)
        row.getCell(startCell).setCellValue(indicator.pastYears.toDouble())
    row.createCell(startCell + 1)
    if(indicator.currentYear != 0)
        row.getCell(startCell + 1).setCellValue(indicator.currentYear.toDouble())

}

fun createTotalSumm(stat: List<StatForOutXLS>) : StatForOutXLS {
    val name = "Всього"
    val past = stat.sumBy { it.pastYears }
    val cur = stat.sumBy { it.currentYear }
    return StatForOutXLS(name, past, cur)
}

fun setPrintSetup(wb: HSSFWorkbook) {
    var ps = wb.getSheetAt(0).printSetup
    wb.getSheetAt(0).autobreaks = true
    ps.fitHeight = 1
    ps.fitWidth = 1

    ps = wb.getSheetAt(1).printSetup
    wb.getSheetAt(0).autobreaks = true
    ps.fitHeight = 1
    ps.fitWidth = 1

}

fun getStringOfMonth() : String =
    when(LocalDate.now().monthValue) {
        1 -> "січень"
        2 -> "лютий"
        3 -> "березень"
        4 -> "квітень"
        5 -> "травень"
        6 -> "червень"
        7 -> "липень"
        8 -> "серпень"
        9 -> "вересень"
        10 -> "жовтень"
        11 -> "листопад"
        12 -> "грудень"
        else -> throw IndexOutOfBoundsException("Невірний місяць")
}

private fun headerNames() : List<String> {
    val wb = HSSFWorkbook(FileInputStream(INPUT_XLS_FILE_NAME))
    var numRow = 3
    val sheet = wb.getSheetAt(0)
    val listHeader = mutableListOf("")
    for(i in 0..(sheet.getRow(numRow).lastCellNum - 1)) {
        val cellValue = sheet.getRow(numRow).getCell(i).stringCellValue
        if(cellValue == "Форма2") {
            numRow++
            listHeader.add(sheet.getRow(numRow).getCell(i).stringCellValue)
        }
        else
            listHeader.add(cellValue)
    }
    listHeader.removeAt(0)
    wb.close()
    return listHeader
}
