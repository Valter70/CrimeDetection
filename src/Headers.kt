import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress


fun writeHeaderOnMainSheet(wb: HSSFWorkbook) {
    val sheet = wb.getSheetAt(0)
    writeFirstTitle(wb, sheet, 8)
    writeSecondTitle(wb, sheet, 8)
    setColumnWidth(sheet, 0..8, GLOBAL_COLUMN_WIDTH)
    writeTabTitleOnMainSheet(wb, sheet)
}

fun writeHeaderOnDetailSheet(wb: HSSFWorkbook) {
    val sheet = wb.getSheetAt(1)
    //Каждя служба по 2 колонки плюс общие цифры
    val lastColumn = statByDepart.size * 2
    writeFirstTitle(wb, sheet, lastColumn)
    writeSecondTitle(wb, sheet, lastColumn)
    setColumnWidth(sheet, 0..lastColumn, DETAIL_COLUMN_WIDTH)
    writeTabTitleOnDetailSheet(wb, sheet)
}

fun writeFirstTitle(wb: HSSFWorkbook, sheet: HSSFSheet, lastColumnTitle: Int) {
    val style = createStyleForTitle(wb)
    style.setFont(setFontForFirstTitle(wb.createFont()))
    sheet.addMergedRegion(CellRangeAddress(0, 0, 0, lastColumnTitle))
    val row = sheet.createRow(0)
    row.heightInPoints = 31.5F
    writeCellTitle(row, "ПОВІДОМЛЕННЯ ПРО ПІДОЗРУ", style)
}

fun writeSecondTitle(wb: HSSFWorkbook, sheet: HSSFSheet, lastColumnTitle: Int) {
    val style = createStyleForTitle(wb)
    style.setFont(setFontForSecondTitle(wb.createFont()))
    sheet.addMergedRegion(CellRangeAddress(1, 1, 0, lastColumnTitle))
    val row = sheet.createRow(1)
    row.heightInPoints = 21F
    writeCellTitle(row, "по службах Рівненського ВП за грудень ${CURRENT_YEAR}р.", style)
}

fun writeTabTitleOnMainSheet(wb: HSSFWorkbook, sheet: HSSFSheet) {

    val row = sheet.createRow(4)
    row.heightInPoints = 26F
    row.createCell(0).setCellValue("Служба")
    row.createCell(1).setCellValue("Всього")
    row.createCell(2).setCellValue("мин.роки")
    row.createCell(3).setCellValue(CURRENT_YEAR.toString())
    val style = createStyleWithBorder(wb)
    style.setFont(setFontForTabTitle(wb.createFont(), 18))
    setStyleForTabTitle(0..3, row, style)

    row.createCell(5).setCellValue("Стаття")
    row.createCell(6).setCellValue("Всього")
    row.createCell(7).setCellValue("мин.роки")
    row.createCell(8).setCellValue(CURRENT_YEAR.toString())
    setStyleForTabTitle(5..8, row, style)
}

fun writeTabTitleOnDetailSheet(wb: HSSFWorkbook, sheet: HSSFSheet) {
    wb.getSheetAt(1).createRow(4)
    wb.getSheetAt(1).createRow(5)
    val style = createStyleWithBorder(wb)
    style.setFont(setFontForTabTitle(wb.createFont(), 16))

    sheet.getRow(4).createCell(0)
    sheet.getRow(5).createCell(0)
    sheet.addMergedRegion(CellRangeAddress(4, 5, 0, 0))

    for(col in 1..(statByDepart.size)) {
        writeBlockTitle(sheet, (col * 2 - 1), statByDepart[col - 1].nameIndicator)
    }
    setStyleForTabTitle(0..(statByDepart.size * 2), sheet.getRow(4), style)
    setStyleForTabTitle(0..(statByDepart.size * 2), sheet.getRow(5), style)

}

fun writeBlockTitle(sheet: HSSFSheet, startColumn: Int, title: String) {
    val endColumn = startColumn + 1
    sheet.addMergedRegion(CellRangeAddress(4, 4, startColumn, endColumn))
    sheet.getRow(4) .createCell(startColumn).setCellValue(title)
    sheet.getRow(4) .createCell(endColumn)

    sheet.getRow(5).createCell(startColumn).setCellValue("мин.роки")
    sheet.getRow(5).createCell(endColumn).setCellValue("2018")

}

fun writeCellTitle(row: HSSFRow, cellValue: String, cellStyle: HSSFCellStyle) {
    val cell = row.createCell(0)
    cell.setCellValue(cellValue)
    cell.setCellStyle(cellStyle)
}