package com.example.studentgradecalculator

import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import com.example.studentgradecalculator.ui.theme.StudentGradeCalculatorTheme
import org.apache.poi.xssf.usermodel.XSSFWorkbook

// ---------------- DATA MODEL ----------------
data class Student(
    val name: String,
    val score: Int?,
    val grade: Char?
)

// ---------------- MAIN ACTIVITY ----------------
class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        setContent {
            StudentGradeCalculatorTheme {
                StudentScreen()
            }
        }
    }
}

// ---------------- UI ----------------
@Composable
fun StudentScreen() {

    val context = LocalContext.current
    var students by remember { mutableStateOf<List<Student>>(emptyList()) }

    val launcher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let {
            students = readExcel(context, it)
        }
    }

    Scaffold(
        modifier = Modifier.fillMaxSize()
    ) { padding ->

        Column(
            modifier = Modifier
                .padding(padding)
                .padding(16.dp)
        ) {

            Button(onClick = { launcher.launch("*/*") }) {
                Text("Import Excel File")
            }

            Spacer(modifier = Modifier.height(16.dp))

            LazyColumn {
                items(students) { student ->
                    Text(
                        text = if (student.score != null)
                            "${student.name} scored ${student.score} : Grade ${student.grade}"
                        else
                            "No score for ${student.name}"
                    )
                    Spacer(modifier = Modifier.height(8.dp))
                }
            }
        }
    }
}

// ---------------- LOGIC ----------------
fun getGrade(score: Int): Char {
    return when (score) {
        in 85..100 -> 'A'
        in 70..84 -> 'B'
        in 55..69 -> 'C'
        in 45..54 -> 'D'
        else -> 'F'
    }
}

fun readExcel(context: android.content.Context, uri: Uri): List<Student> {

    val students = mutableListOf<Student>()

    context.contentResolver.openInputStream(uri)?.use { inputStream ->

        val workbook = XSSFWorkbook(inputStream)
        val sheet = workbook.getSheetAt(0)

        for (rowIndex in 1..sheet.lastRowNum) {

            val row = sheet.getRow(rowIndex) ?: continue

            val name = row.getCell(0)?.stringCellValue ?: continue
            val score = row.getCell(1)?.numericCellValue?.toInt()
            val grade = score?.let { getGrade(it) }

            students.add(Student(name, score, grade))
        }

        workbook.close()
    }

    return students
}