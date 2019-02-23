import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress


fun writeHeaderOnMainSheet(wb: HSSFWorkbook, sheet: HSSFSheet) {

    writeFirstTitleOnMainSheet(wb, sheet)
    writeSecondTitleOnMainSheet(wb, sheet)
    setColumnWidthOnMainSheet(sheet)
    writeTabTitleOnMainSheet(wb, sheet)
}

fun writeFirstTitleOnMainSheet(wb: HSSFWorkbook, sheet: HSSFSheet) {
    val style = createStyleForTitle(wb)
    style.setFont(setFontForFirstTitle(wb.createFont()))
    sheet.addMergedRegion(CellRangeAddress(0, 0, 0, 8))
    val row = sheet.createRow(0)
    row.heightInPoints = 31.5F
    writeCellTitle(row, "ПОВІДОМЛЕННЯ ПРО ПІДОЗРУ", style)
}

fun writeSecondTitleOnMainSheet(wb: HSSFWorkbook, sheet: HSSFSheet) {
    val style = createStyleForTitle(wb)
    style.setFont(setFontForSecondTitle(wb.createFont()))
    sheet.addMergedRegion(CellRangeAddress(1, 1, 0, 8))
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
    style.setFont(setFontForTabTitle(wb.createFont()))
    setStyleForTabTitle(0..3, row, style)

    row.createCell(5).setCellValue("Стаття")
    row.createCell(6).setCellValue("Всього")
    row.createCell(7).setCellValue("мин.роки")
    row.createCell(8).setCellValue(CURRENT_YEAR.toString())
    setStyleForTabTitle(5..8, row, style)
}

fun writeCellTitle(row: HSSFRow, cellValue: String, cellStyle: HSSFCellStyle) {
    val cell = row.createCell(0)
    cell.setCellValue(cellValue)
    cell.setCellStyle(cellStyle)
}