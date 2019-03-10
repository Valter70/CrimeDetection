// Тяжкость КП
enum class GravityOfCrime(val statName: String) {
    T1("НЕВЕЛИКОЇ ТЯЖКОСТI"),
    T2("СЕРЕДНЬОЇ ТЯЖКОСТI"),
    T3("ТЯЖКИЙ"),
    T4("ОС.ТЯЖКИЙ")
}

fun getCodeGravity(strGravity: String) =
    when(strGravity){
        "НЕВЕЛИКОЇ ТЯЖКОСТI" -> GravityOfCrime.T1
        "СЕРЕДНЬОЇ ТЯЖКОСТI" -> GravityOfCrime.T2
        "ТЯЖКИЙ" -> GravityOfCrime.T3
        "ОС.ТЯЖКИЙ" -> GravityOfCrime.T4
        else -> throw IndexOutOfBoundsException("Невірна назва тяжкості: $strGravity")
    }

fun createStatByGravity() : MutableList<StatForOutXLS> {
    val statByGravity: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))
    for(rec in GravityOfCrime.values()) {
        val gravity = rec.statName.toLowerCase()
        val currentYear = crimeList.count { it.isCurrentYear && it.gravity == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.gravity == rec }
        statByGravity.add(StatForOutXLS(gravity, pastYear, currentYear))
    }

    statByGravity.removeAt(0)
    return statByGravity
}
