import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook

fun writeTabOnDetailSheet(wb: HSSFWorkbook) {
    val totalSummRow = wb.getSheetAt(1).createRow(6)
    writeDetailTotalSumm(wb, totalSummRow)
    setStyleForTotalSummCells(wb, 0..(statByDepart.size * 2), totalSummRow)
    writeDetailArticleName(wb)


}

fun writeDetailTotalSumm(wb: HSSFWorkbook, totalSummRow: HSSFRow) {
    totalSummRow.createCell(0).setCellValue("Всього")
    for(col in 1..(statByDepart.size)) {
        totalSummRow.createCell(col * 2 - 1).setCellValue(statByDepart[col - 1].pastYears.toDouble())
        totalSummRow.createCell(col * 2).setCellValue(statByDepart[col - 1].currentYear.toDouble())
    }

}

fun writeDetailArticleName(wb: HSSFWorkbook) {
    var part185 = 0
    val sheet = wb.getSheetAt(1)
    sheet.getRow(FIRST_DETAIL_INDICATOR_ROW).createCell(0).setCellValue("тяжкі")
    for(index in 1..(statByArticle.size)) {
        val row = sheet.getRow(FIRST_DETAIL_INDICATOR_ROW + index + part185)
        val articleName = statByArticle[index - 1].nameIndicator.substring(3)
        row.createCell(0).setCellValue(articleName)
        if(articleName == "185") {
            part185 = 5
            writeDetailArticle185Title(sheet, row.rowNum + 1)
        }
    }

}

fun writeDetailArticle185Title(sheet: HSSFSheet, startRow: Int) {
    for(index in 0..(statByArticle185.size - 1)) {
        val cell = sheet.getRow(startRow + index).createCell(0)
        cell.setCellValue(statByArticle185[index].nameIndicator)
    }
}