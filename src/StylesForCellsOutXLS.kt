import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.*

fun setStyleForTabTitle(range: IntRange, row: HSSFRow, style: HSSFCellStyle) {
    for(cell in range) {
        row.getCell(cell).setCellStyle(style)
    }
}

// 4800 соответствует размеру 18 в Excel
fun setColumnWidthOnMainSheet(sheet: HSSFSheet) {
    val setOfColumn = setOf(0, 1, 2, 3 ,5, 6 ,7 ,8)
    for(col in setOfColumn)
        sheet.setColumnWidth(col, 4800)
}

fun setFontForFirstTitle(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 24
    font.bold = true
    font.underline = Font.U_SINGLE
    return font
}

fun setFontForSecondTitle(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 16
    font.italic = true
    return font
}

fun setFontForTabTitle(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 18
    font.bold = true
    return font
}

fun setFontForTabIndicator(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 18
    return font
}

fun setFontForTotalSumm(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 18
    font.bold = true
    font.underline = Font.U_SINGLE
    return font
}

fun createStyleForTitle(wb: HSSFWorkbook) : HSSFCellStyle {
    val style = wb.createCellStyle()
    style.alignment = HorizontalAlignment.CENTER
    style.verticalAlignment = VerticalAlignment.CENTER
    return style
}

fun createStyleWithBorder(wb: HSSFWorkbook) : HSSFCellStyle {
    val style = createStyleForTitle(wb)
    style.borderTop = BorderStyle.THIN
    style.borderBottom = BorderStyle.THIN
    style.borderLeft = BorderStyle.THIN
    style.borderRight = BorderStyle.THIN
    return  style
}

fun setRowHeightOnMainSheet(sheet: HSSFSheet, range: IntRange) {
    for(row in range) {
        sheet.createRow(row).heightInPoints = 26F
    }
}

fun setStyleForTotalSummCells(wb: HSSFWorkbook, rangeColumn: IntRange, row: HSSFRow) {
    val style = createStyleWithBorder(wb)
    style.setFont(setFontForTotalSumm(wb.createFont()))
    setStyleForTabTitle(rangeColumn, row, style)

}

fun setStyleForIndicatorRow(wb: HSSFWorkbook, rangeRows: IntRange, rangeColumns: IntRange) {
    val style = createStyleWithBorder(wb)
    style.setFont(setFontForTabIndicator(wb.createFont()))

    for(rowNum in rangeRows) {
        val row = wb.getSheetAt(0).getRow(rowNum)
        setStyleForTabTitle(rangeColumns, row, style)
    }

}