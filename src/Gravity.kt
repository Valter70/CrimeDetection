// Тяжкость КП
enum class Gravity(val statName: String) {
    T1("НЕВЕЛИКОЇ ТЯЖКОСТI"),
    T2("СЕРЕДНЬОЇ ТЯЖКОСТI"),
    T3("ТЯЖКИЙ"),
    T4("ОСОБЛИВО ТЯЖКИЙ")
}

fun getCodeGravity(strGravity: String) =
    when(strGravity){
        "НЕВЕЛИКОЇ ТЯЖКОСТI" -> Gravity.T1
        "СЕРЕДНЬОЇ ТЯЖКОСТI" -> Gravity.T2
        "ТЯЖКИЙ" -> Gravity.T3
        "ОСОБЛИВО ТЯЖКИЙ" -> Gravity.T4
        else -> throw IndexOutOfBoundsException("Невірна назва тяжкості: $strGravity")
    }

fun createStatByGravity() : MutableList<StatForOutXLS> {
    val statByGravity: MutableList<StatForOutXLS> = mutableListOf(StatForOutXLS(""))
    for(rec in Gravity.values()) {
        val gravity = if(rec.statName.toLowerCase() == "особливо тяжкий") "ос.тяжкий"
                      else rec.statName.toLowerCase()
        val currentYear = crimeList.count { it.isCurrentYear && it.gravity == rec }
        val pastYear = crimeList.count { !it.isCurrentYear && it.gravity == rec }
        statByGravity.add(StatForOutXLS(gravity, pastYear, currentYear))
    }

    statByGravity.removeAt(0)
    return statByGravity
}
