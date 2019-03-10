// Структура для формирования списка КП
data class CrimeCaseF2ForStat(
    val number: String,
    val article: String,
    val depart: Department,
    val gravity: Gravity,
    val isCurrentYear: Boolean = number.substring(1, 5).toInt() == CURRENT_YEAR) {

    override fun toString() =
        "Провадження(№ $number; стаття-$article; служба-${depart.shortName}; ${gravity.statName})"
}