import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet

fun writeTabOnDetailSheet() {
    val totalSummRow = WB_OUT.getSheetAt(1).createRow(6)
    createAllDetailCells()
    writeDetailTotalSumm(totalSummRow)
    setStyleForTotalSummCells(0..(statByDepart.size * 2), totalSummRow)
    writeDetailArticleName()
    writeDetailGravity()
    writeDetailIndicator()

    val rangRow = FIRST_DETAIL_INDICATOR_ROW ..WB_OUT.getSheetAt(1).lastRowNum
    val endColumn = WB_OUT.getSheetAt(1).getRow(FIRST_DETAIL_INDICATOR_ROW - 1).lastCellNum - 1
    val rangeColumn = 0..endColumn
    setStyleForIndicatorRow(rangRow, rangeColumn, 1)

}

fun writeDetailTotalSumm(totalSummRow: HSSFRow) {
    totalSummRow.createCell(0).setCellValue("Всього")
    for(col in 1..(statByDepart.size)) {
        totalSummRow.createCell(col * 2 - 1).setCellValue(statByDepart[col - 1].pastYears.toDouble())
        totalSummRow.createCell(col * 2).setCellValue(statByDepart[col - 1].currentYear.toDouble())
    }
}

fun writeDetailArticleName() {
    var part185 = 0
    val sheet = WB_OUT.getSheetAt(1)
    sheet.getRow(FIRST_DETAIL_INDICATOR_ROW).createCell(0).setCellValue("тяжкі")
    for(index in 1..(statByArticle.size)) {
        val row = sheet.getRow(FIRST_DETAIL_INDICATOR_ROW + index + part185)
        val articleName = statByArticle[index - 1].nameIndicator.substring(3)
        row.createCell(0).setCellValue("ст.$articleName")
        if(articleName == "185") {
            part185 = 5
            writeDetailArticle185Title(sheet, row.rowNum + 1)
        }
    }
}

fun writeDetailGravity() {
    val row = WB_OUT.getSheetAt(1).getRow(FIRST_DETAIL_INDICATOR_ROW)
    row.createCell(1).setCellValue(statByGravity[2].pastYears.toDouble())
    row.createCell(2).setCellValue(statByGravity[2].currentYear.toDouble())

    statByDepart.removeAt(0)
    var currentBlock = 2
    for(depat in statByDepart) {
        val depatCode = getCodeDepart(depat.nameIndicator)
        val pastYear = crimeList.count { !it.isCurrentYear && it.depart == depatCode && it.gravity == Gravity.T3 }
        val currentYear = crimeList.count { it.isCurrentYear && it.depart == depatCode && it.gravity == Gravity.T3 }
        writeYearIndicatorToTab(row, (currentBlock * 2 - 1), StatForOutXLS("", pastYear, currentYear))
        currentBlock++
    }

}

fun writeDetailArticle185Title(sheet: HSSFSheet, startRow: Int) {
    for(index in 0..(statByArticle185.size - 1)) {
        val cell = sheet.getRow(startRow + index).createCell(0)
        cell.setCellValue(statByArticle185[index].nameIndicator)
    }
}

fun writeDetailIndicator() {
    val startCell = Cell(FIRST_DETAIL_INDICATOR_ROW + 1, 1)
    writeDetailDepartIndicator(startCell, statByArticle, statByArticle185)

    var tabBlock = 2
    for(depat in statByDepart) {
        val crimeBlock = crimeList.filter { it.depart == getCodeDepart(depat.nameIndicator) }
        val statByBlock = createStatByBlock(crimeBlock)
        val statBy185Block = createStatBy185Block(crimeBlock)
        startCell.cell = tabBlock * 2 - 1
        writeDetailDepartIndicator(startCell, statByBlock, statBy185Block)
        tabBlock++
    }

}

fun writeDetailDepartIndicator(startCell: Cell, statByBlock: List<StatForOutXLS>, statBy185Block: List<StatForOutXLS>) {
    val setOfArticle = statByArticle.map { it.nameIndicator.substringAfter('.') }.toSet()
    for(index in 0..(statByBlock.size - 1)) {
        val articleName = statByBlock[index].nameIndicator.substringAfter('.')
        val rowIndex = getArticleIndex(setOfArticle, articleName) + 8
        val row = WB_OUT.getSheetAt(1).getRow(rowIndex)
        writeYearIndicatorToTab(row, startCell.cell, statByBlock[index])
        if(articleName == "185") {
            startCell.row = rowIndex + 1
            writeDetailArticle185Indicator(startCell, statBy185Block)
        }
    }
}

fun writeDetailArticle185Indicator(startCell: Cell, statBy185Block: List<StatForOutXLS>) {
    for(index in 0..(statByArticle185.size - 1)) {
        val row = WB_OUT.getSheetAt(1).getRow(startCell.row + index)
        writeYearIndicatorToTab(row, startCell.cell, statBy185Block[index])
    }
}

fun createStatByBlock(crimeBlock: List<CriminalCase>) : List<StatForOutXLS> {
    val listOfArticle = crimeBlock.map { it.article.substringBefore(' ') }.toSet()
    val statByArticle: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))

    for(rec in listOfArticle) {
        val currentYear = crimeBlock.count { it.isCurrentYear && it.article.substringBefore(' ') == rec }
        val pastYear = crimeBlock.count { !it.isCurrentYear && it.article.substringBefore(' ') == rec }
        statByArticle.add(StatForOutXLS("ст.$rec", pastYear, currentYear))
    }
    statByArticle.removeAt(0)
    return statByArticle

}

fun createStatBy185Block(crimeBlock: List<CriminalCase>) : List<StatForOutXLS> {
    val statBy185Block: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))
    for(rec in 1..5) {
        val part185 = "Ч.$rec"
        val currentYear = crimeBlock.count { it.isCurrentYear && it.article == "185 $part185" }
        val pastYear = crimeBlock.count { !it.isCurrentYear && it.article == "185 $part185" }
        statBy185Block.add(StatForOutXLS("-" + part185.toLowerCase(), pastYear, currentYear))
    }
    statBy185Block.removeAt(0)
    return statBy185Block
}


fun getArticleIndex(setOfArticle: Set<String>, article: String) : Int {
    var result = setOfArticle.indexOf(article)
    if(article > "185") {
        result += 5
    }
    return result
}

fun createAllDetailCells() {
    val lasrColumn = statByDepart.size * 2 + 2
    val lastRow = WB_OUT.getSheetAt(1).lastRowNum
    for(index in 7..lastRow) {
        createRowCells(WB_OUT.getSheetAt(1).getRow(index), lasrColumn)
    }
}

fun createRowCells(row: HSSFRow, lasrColumn: Int) {
    for(index in 1..lasrColumn) {
        row.createCell(index)
    }
}