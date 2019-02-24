// Службы раскрытия
enum class Department(val statName: String, val shortName: String) {
    VKP("ПРАЦ.КР", "ВКП"),
    VPD("ДIЛЬНИЧИЙ ОФІЦЕР ПОЛІЦІЇ", "ВПД"),
    UPPD("ПРАЦ.ПРЕВЕНТ. ДІЯЛЬНОСТІ З ЮВЕН. ПРЕВЕНЦІЇ", "ЮППД"),
    SV("СЛІДЧИЙ НАЦІОНАЛЬНОЇ ПОЛІЦІЇ", "СВ"),
    UPN("ПРАЦ. ПРОТИДІЙ НАРКОЗЛОЧИННОСТІ", "УПН"),
    VVB("ПРАЦ.ВБ", "ВВБ"),
    UZE("ПРАЦ.ДЕПАРТ. ЗАХИЧТУ ЕКОНОМІКИ (ДЗЕ)", "УЗЕ"),
    UKP("ПРАЦ.ПIДРОЗ. КІБЕРПОЛІЦІЇ", "УКП"),
    VBTL("ПРАЦ.ПIДРОЗД.БОР.ЗI ЗЛОЧ.ПОВЯЗАН.З ТОРГ.ЛЮДЬМИ", "ВБЗПТЛ")

}

val DEPART_GUNP = setOf<String>("УПН", "ВВБ", "УЗЕ", "УКП", "ВБЗПТЛ")

fun getCodeDepart(strDepart: String) =
    when(strDepart) {
        "ДIЛЬНИЧИЙ ОФІЦЕР ПОЛІЦІЇ" -> Department.VPD
        "ПРАЦ.КР" -> Department.VKP
        "ПРАЦ.ПРЕВЕНТ. ДІЯЛЬНОСТІ З ЮВЕН. ПРЕВЕНЦІЇ" -> Department.UPPD
        "СЛІДЧИЙ НАЦІОНАЛЬНОЇ ПОЛІЦІЇ" -> Department.SV
        "ПРАЦ. ПРОТИДІЙ НАРКОЗЛОЧИННОСТІ" -> Department.UPN
        "ПРАЦ.ВБ" -> Department.VVB
        "ПРАЦ.ДЕПАРТ. ЗАХИЧТУ ЕКОНОМІКИ (ДЗЕ)" -> Department.UZE
        "ПРАЦ.ПIДРОЗ. КІБЕРПОЛІЦІЇ" -> Department.UKP
        "ПРАЦ.ПIДРОЗД.БОР.ЗI ЗЛОЧ.ПОВЯЗАН.З ТОРГ.ЛЮДЬМИ" -> Department.VBTL
        else -> throw IndexOutOfBoundsException("Невірна назва служби: $strDepart")
    }
