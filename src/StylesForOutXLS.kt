import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.*

fun setStyleForTabTitle(rangeCell: IntRange, row: HSSFRow, style: HSSFCellStyle) {
    for(cell in rangeCell) {
        row.getCell(cell).setCellStyle(style)
    }
}

fun setColumnWidth(sheet: HSSFSheet, columnRange: IntRange, columnWidth: Int) {
    for(col in columnRange)
        sheet.setColumnWidth(col, columnWidth)
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

fun setFontForTabTitle(font: Font, height: Short) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = height
    font.bold = true
    return font
}

fun setFontForTabIndicator(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 16
    return font
}

fun setFontForTotalSumm(font: Font) : Font {
    font.fontName = "Colibri"
    font.fontHeightInPoints = 16
    font.bold = true
    font.underline = Font.U_SINGLE
    return font
}

fun createStyleForTitle() : HSSFCellStyle {
    val style = WB_OUT.createCellStyle()
    style.alignment = HorizontalAlignment.CENTER
    style.verticalAlignment = VerticalAlignment.CENTER
    return style
}

fun createStyleWithBorder() : HSSFCellStyle {
    val style = createStyleForTitle()
    style.borderTop = BorderStyle.THIN
    style.borderBottom = BorderStyle.THIN
    style.borderLeft = BorderStyle.THIN
    style.borderRight = BorderStyle.THIN
    return  style
}

fun setStyleForTotalSummCells(rangeColumn: IntRange, row: HSSFRow) {
    val style = createStyleWithBorder()
    style.setFont(setFontForTotalSumm(WB_OUT.createFont()))
    setStyleForTabTitle(rangeColumn, row, style)
}

fun setStyleForIndicatorRow(rangeRows: IntRange, rangeColumns: IntRange, sheet: Int) {
    val style = createStyleWithBorder()
    style.setFont(setFontForTabIndicator(WB_OUT.createFont()))

    for(rowNum in rangeRows) {
        val row = WB_OUT.getSheetAt(sheet).getRow(rowNum)
        setStyleForTabTitle(rangeColumns, row, style)
    }
}
