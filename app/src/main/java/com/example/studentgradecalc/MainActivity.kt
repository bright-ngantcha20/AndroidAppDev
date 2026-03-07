import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

// 1️⃣ Data Class
data class Student(val name: String, val score: Int?)

// 2️⃣ Grade Function
fun getGrade(score: Int): Char = when(score) {
    in 90..100 -> 'A'
    in 80..89  -> 'B'
    in 70..79  -> 'C'
    in 60..69  -> 'D'
    else       -> 'F'
}

// 3️⃣ Function to Read Excel File
fun readStudentsFromExcel(filePath: String): List<Student> {
    val students = mutableListOf<Student>()

    FileInputStream(File(filePath)).use { file ->
        val workbook = XSSFWorkbook(file)
        val sheet = workbook.getSheetAt(0)

        // Skip header row
        for (row in sheet.drop(1)) {
            val name = row.getCell(0)?.stringCellValue ?: continue
            val scoreCell = row.getCell(1)

            val score = when (scoreCell?.cellType) {
                org.apache.poi.ss.usermodel.CellType.NUMERIC -> scoreCell.numericCellValue.toInt()
                else -> null
            }

            students.add(Student(name, score))
        }

        workbook.close()
    }

    return students
}

// 4️⃣ Validation Function
fun validateScore(score: Int?): Boolean {
    return score != null && score in 0..100
}

// 5️⃣ Formatting Function
fun formatStudent(student: Student): String {
    return if (student.score == null) {
        "No score for ${student.name}"
    } else {
        "${student.name} scored ${student.score}"
    }
}

// 6️⃣ Main Function Demonstration
fun main() {

    val students = readStudentsFromExcel("students.xlsx") // Update path if needed

    println("All Students:")
    students.forEach { println(formatStudent(it)) }

    println("\nStudents who passed:")
    // ✅ Using validateScore function here
    val passedStudents = students.filter { validateScore(it.score) && it.score!! >= 60 }
    passedStudents.forEach {
        val grade = getGrade(it.score!!)
        println("${it.name} : Grade $grade")
    }
}