// Тяжкість КП
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
