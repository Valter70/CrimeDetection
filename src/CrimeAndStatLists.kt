val crimeList = createCrimeListFromXLS()
val crimeOfArticle185 = crimeList.filter { it.article.substringBefore(' ') == "185" }

val statByDepart = createStatByDepart()
val statByArticle = createStatByArticle()
val statByArticle185 = createStatByArticle185()
