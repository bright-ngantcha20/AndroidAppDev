package com.example.studentgradecalc

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.material3.Text
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.padding
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream

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

// 3️⃣ Validation Function
fun validateScore(score: Int?): Boolean {
    return score != null && score in 0..100
}

// 4️⃣ Formatting Function
fun formatStudent(student: Student): String {
    return if (student.score == null) {
        "No score for ${student.name}"
    } else {
        "${student.name} scored ${student.score}"
    }
}

class MainActivity : ComponentActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        // Read Excel file
        val students = readStudentsFromExcel()

        // All students formatted
        val allStudents = students.map { formatStudent(it) }

        // Students who passed
        val passedStudents = students.filter { validateScore(it.score) && it.score!! >= 60 }

        val passedWithGrades = passedStudents.map {
            "${it.name} : Grade ${getGrade(it.score!!)}"
        }

        setContent {

            Column(modifier = Modifier.padding(16.dp)) {

                Text("All Students:")

                allStudents.forEach {
                    Text(it)
                }

                Text("\nStudents Who Passed:")

                passedWithGrades.forEach {
                    Text(it)
                }
            }
        }
    }

    // 5️⃣ Function to Read Excel from Assets
    private fun readStudentsFromExcel(): List<Student> {

        val students = mutableListOf<Student>()

        val inputStream: InputStream = assets.open("students.xlsx")

        val workbook = XSSFWorkbook(inputStream)
        val sheet = workbook.getSheetAt(0)

        for (i in 1..sheet.lastRowNum) {

            val row = sheet.getRow(i)

            val name = row.getCell(0)?.stringCellValue ?: continue
            val scoreCell = row.getCell(1)

            val score = when(scoreCell?.cellType) {
                CellType.NUMERIC -> scoreCell.numericCellValue.toInt()
                else -> null
            }

            students.add(Student(name, score))
        }

        workbook.close()

        return students
    }
}