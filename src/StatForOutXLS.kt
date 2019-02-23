// Структура для формирования статистики
data class StatForOutXLS(
    val nameIndicator: String,
    var pastYears: Int = 0,
    var currentYear: Int = 0,
    var total: Int = pastYears + currentYear
)

