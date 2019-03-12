import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.util.CellRangeAddress
import java.io.FileInputStream

val wb = HSSFWorkbook(FileInputStream(INPUT_XLS_FILE_NAME))

fun writeDepartTabOnMainSheet() {
    writeVPIndicators()
    writeGUNPTitle()

    statByDepart.removeIf { it.total == 0 && it.nameIndicator in DEPART_GUNP }

    writeGUNPIndicators()
    setDepartIndicatorsStyle()
    writeDepartTotalSumm()

    wb.close()
}

fun writeVPIndicators() {
    val sheet = wb.getSheetAt(0)
    for(index in 0..3) {
        val row = sheet.getRow(FIRST_GLOBAL_INDICATOR_ROW + index)
        writeIndicatorToTab(row, 0, statByDepart[index])
    }
}

fun writeGUNPTitle() {
    val sheet = wb.getSheetAt(0)
    sheet.addMergedRegion(CellRangeAddress(9, 9, 0, 3))
    val style = createStyleForTitle(wb)
    style.setFont(setFontForTabTitle(wb.createFont(), 18))
    writeCellTitle(sheet.getRow(9), "ГУНП", style)
}

fun writeGUNPIndicators() {
    val sheet = wb.getSheetAt(0)
    for(index in 4..(statByDepart.size - 1)) {
        val row = sheet.getRow(index + 6)
        writeIndicatorToTab(row, 0, statByDepart[index])
    }
}

fun createStatByDepart() : MutableList<StatForOutXLS> {
    val statByDepart: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))
    for(rec in Department.values()) {
        val depart = rec.shortName
        val currentYear = crimeList.count { it.isCurrentYear && it.depart == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.depart == rec }
        statByDepart.add(StatForOutXLS(depart, pastYear, currentYear))
    }
    statByDepart.removeAt(0)
    return  statByDepart
}

fun writeDepartTotalSumm() {
    val endRowIndex = FIRST_GLOBAL_INDICATOR_ROW + statByDepart.size + 1
    val totalSummRow = wb.getSheetAt(0).getRow(endRowIndex)
    val totalSumm = createTotalSumm(statByDepart)

    writeIndicatorToTab(totalSummRow, DEPART_COLUMN.first, totalSumm)

    setStyleForTotalSummCells(wb, DEPART_COLUMN, totalSummRow)

}

fun setDepartIndicatorsStyle() {
    var rangeRows = FIRST_GLOBAL_INDICATOR_ROW..(FIRST_GLOBAL_INDICATOR_ROW + 3)
    setStyleForIndicatorRow(wb, rangeRows, DEPART_COLUMN, 0)

    rangeRows = (FIRST_GLOBAL_INDICATOR_ROW + 5)..(FIRST_GLOBAL_INDICATOR_ROW + statByDepart.size)
    setStyleForIndicatorRow(wb, rangeRows, DEPART_COLUMN, 0)

}