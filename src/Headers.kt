import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress


fun writeHeaderOnMainSheet() {
    val sheet = WB_OUT.getSheetAt(0)
    writeFirstTitle(sheet, 8)
    writeSecondTitle(sheet, 8)
    setColumnWidth(sheet, 0..8, GLOBAL_COLUMN_WIDTH)
    writeTabTitleOnMainSheet(sheet)
}

fun writeHeaderOnDetailSheet() {
    val sheet = WB_OUT.getSheetAt(1)
    //Каждя служба по 2 колонки
    val lastColumn = statByDepart.size * 2
    writeFirstTitle(sheet, lastColumn)
    writeSecondTitle(sheet, lastColumn)
    setColumnWidth(sheet, 0..lastColumn, DETAIL_COLUMN_WIDTH)
    writeTabTitleOnDetailSheet(sheet)
}

fun writeFirstTitle(sheet: HSSFSheet, lastColumnTitle: Int) {
    val style = createStyleForTitle()
    style.setFont(setFontForFirstTitle(WB_OUT.createFont()))
    sheet.addMergedRegion(CellRangeAddress(0, 0, 0, lastColumnTitle))
    val row = sheet.createRow(0)
    row.heightInPoints = 31.5F
    writeCellTitle(row, "ПОВІДОМЛЕННЯ ПРО ПІДОЗРУ", style)
}

fun writeSecondTitle(sheet: HSSFSheet, lastColumnTitle: Int) {
    val style = createStyleForTitle()
    style.setFont(setFontForSecondTitle(WB_OUT.createFont()))
    sheet.addMergedRegion(CellRangeAddress(1, 1, 0, lastColumnTitle))
    val row = sheet.createRow(1)
    row.heightInPoints = 21F
    writeCellTitle(row, "по службах Рівненського ВП за ${getStringOfMonth()} ${CURRENT_YEAR}р.", style)
}

fun writeTabTitleOnMainSheet(sheet: HSSFSheet) {

    val row = sheet.createRow(4)
    row.heightInPoints = 26F
    createTabTitleCell(row, DEPART_COLUMN.first)
    val style = createStyleWithBorder()
    style.setFont(setFontForTabTitle(WB_OUT.createFont(), 18))
    setStyleForTabTitle(DEPART_COLUMN, row, style)

    createTabTitleCell(row, ARTICLE_COLUMN.first)
    setStyleForTabTitle(ARTICLE_COLUMN, row, style)
}

fun createTabTitleCell(row: HSSFRow, startColumn: Int) {
    row.createCell(startColumn).setCellValue("Стаття")
    row.createCell(startColumn + 1).setCellValue("Всього")
    row.createCell(startColumn + 2).setCellValue("мин.роки")
    row.createCell(startColumn + 3).setCellValue(CURRENT_YEAR.toString())

}

fun writeTabTitleOnDetailSheet(sheet: HSSFSheet) {
    sheet.createRow(4)
    sheet.createRow(5)
    val style = createStyleWithBorder()
    style.setFont(setFontForTabTitle(WB_OUT.createFont(), 16))

    sheet.getRow(4).createCell(0)
    sheet.getRow(5).createCell(0)
    sheet.addMergedRegion(CellRangeAddress(4, 5, 0, 0))

    for(col in 1..(statByDepart.size)) {
        writeBlockTitle(sheet, (col * 2 - 1), statByDepart[col - 1].nameIndicator)
    }
    setStyleForTabTitle(0..(statByDepart.size * 2), sheet.getRow(4), style)
    setStyleForTabTitle(0..(statByDepart.size * 2), sheet.getRow(5), style)

    writeGUNPHeader(sheet, style)
}

fun writeGUNPHeader(sheet: HSSFSheet, style: HSSFCellStyle) {
    sheet.createRow(3)
    val firstGUNPCol = 11
    val lastGUNPCol = firstGUNPCol + (statByDepart.size - 5) * 2 - 1
    for(index in firstGUNPCol..lastGUNPCol)
        sheet.getRow(3).createCell(index)
    sheet.getRow(3).getCell(firstGUNPCol).setCellValue("ІНШІ")
    sheet.addMergedRegion(CellRangeAddress(3, 3, firstGUNPCol, lastGUNPCol))
    setStyleForTabTitle(firstGUNPCol..lastGUNPCol, sheet.getRow(3), style)

}

fun writeBlockTitle(sheet: HSSFSheet, startColumn: Int, title: String) {
    val endColumn = startColumn + 1
    sheet.addMergedRegion(CellRangeAddress(4, 4, startColumn, endColumn))
    sheet.getRow(4) .createCell(startColumn).setCellValue(title)
    sheet.getRow(4) .createCell(endColumn)

    sheet.getRow(5).createCell(startColumn).setCellValue("мин.роки")
    sheet.getRow(5).createCell(endColumn).setCellValue(CURRENT_YEAR.toString())

}

fun writeCellTitle(row: HSSFRow, cellValue: String, cellStyle: HSSFCellStyle) {
    val cell = row.createCell(0)
    cell.setCellValue(cellValue)
    cell.setCellStyle(cellStyle)
}