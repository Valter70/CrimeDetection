// Тяжкость КП
enum class GravityOfCrime(val statName: String) {
    T1("НЕВЕЛИКОЇ ТЯЖКОСТІ"),
    T2("СЕРЕДНЬОЇ ТЯЖКОСТІ"),
    T3("ТЯЖКІ"),
    T4("ОС. ТЯЖКІ")
}

fun getCodeGravity(strGravity: String) =
    when(strGravity){
        "НЕВЕЛИКОЇ ТЯЖКОСТІ" -> GravityOfCrime.T1
        "СЕРЕДНЬОЇ ТЯЖКОСТІ" -> GravityOfCrime.T2
        "ТЯЖКІ" -> GravityOfCrime.T3
        "ОСОБЛИВО ТЯЖКІ" -> GravityOfCrime.T4
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
