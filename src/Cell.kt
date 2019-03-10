class Cell {
    var cell: Int
    var row: Int

    constructor(_row: Int, _cell: Int) {
        row = _row
        cell = _cell
    }

    constructor(nameCell: String) {
        row = nameCell[1].toString().toInt() - 1
        cell = nameCell[0].toInt() - 65
    }

    val namedCell: String
        get() = (cell + 65).toChar().toString() + (row + 1).toString()

    override fun toString() = "row = $row; cel = $cell"

}