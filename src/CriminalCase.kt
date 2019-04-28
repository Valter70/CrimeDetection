import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.FileInputStream
import java.time.LocalDate

class CriminalCase(currentRowIndex: Int) {
    private val wbIn = HSSFWorkbook(FileInputStream(INPUT_XLS_FILE_NAME))
    private val currentRow = wbIn.getSheetAt(0).getRow(currentRowIndex)
    private val titleList = headerNames()
    val number: String
        get() = getNumberValue()
    val article: String
        get() = getArticleValue()
    val depart: Department
        get() = getDepartValue()
    val gravity: Gravity
        get() = getGravityValue()
    val suspicionDate: LocalDate
        get() = getSuspicionDateValue()
    val isCurrentYear: Boolean = number.substring(1..4).toInt() == CURRENT_YEAR

    private fun getNumberValue() : String {
        var numKP = ""
        if (currentRow.getCell(0).numericCellValue == 0.0)
            return "00000"
        for(i in 0..2) {
            val partOfNumber = currentRow.getCell(i).numericCellValue.toInt().toString()
            numKP += partOfNumber
        }
        numKP += createFullNumber(currentRow.getCell(3).numericCellValue.toInt().toString())
        return numKP
    }

    private fun createFullNumber(number: String) : String {
        var result = ""
        for(i in 1..(7 - number.length))
            result += "0"
        return result + number
    }

    private fun getArticleValue() : String {
        val articleCell = titleList.indexOf("Квал.злоч.КК 2001")
        return currentRow.getCell(articleCell).stringCellValue.substring(3)
    }

    private fun getDepartValue() : Department {
        val departCell = headerNames().indexOf("17.Особа виявлена службою")
        return getCodeDepart(currentRow.getCell(departCell).stringCellValue)
    }

    private fun getGravityValue() : Gravity {
        val gravityCell = headerNames().indexOf("Ф1 14.Квал.злоч.-тяжкiсть")
        return getCodeGravity(currentRow.getCell(gravityCell).stringCellValue)
    }

    private fun getSuspicionDateValue() : LocalDate {
        val suspicionDateCell = headerNames().indexOf("19.Дата повiдомлення про пiдозру")
        return java.sql.Date(currentRow.getCell(suspicionDateCell).dateCellValue.time).toLocalDate()
    }

    private fun headerNames() : List<String> {
        var numRow = 3
        val sheet = wbIn.getSheetAt(0)
        val listHeader = mutableListOf("")
        for(i in 0 until sheet.getRow(numRow).lastCellNum) {
            val cellValue = sheet.getRow(numRow).getCell(i).stringCellValue
            if(cellValue == "Форма2") {
                numRow++
                listHeader.add(sheet.getRow(numRow).getCell(i).stringCellValue)
            }
            else
                listHeader.add(cellValue)
        }
        listHeader.removeAt(0)
        return listHeader
    }

}