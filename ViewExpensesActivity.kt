package com.example.expensetracker

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.*
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.FileProvider
import androidx.lifecycle.lifecycleScope
import com.example.expensetracker.data.ExpenseDatabase
import com.example.expensetracker.model.Expense
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

class ViewExpensesActivity : AppCompatActivity() {

    private lateinit var exportButton: Button
    private lateinit var listView: ListView
    private lateinit var db: ExpenseDatabase
    private lateinit var adapter: ArrayAdapter<String>
    private var expenses: List<Expense> = emptyList()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_view_expenses)

        exportButton = findViewById(R.id.exportExcelButton)
        listView = findViewById(R.id.expenseListView)
        db = ExpenseDatabase.getDatabase(this)

        // Load all expenses and set to ListView
        lifecycleScope.launch {
            expenses = db.expenseDao().getAllExpenses()
            val expenseTitles = expenses.map {
                "${it.title} - ₹${it.amount} - ${it.date}"
            }
            adapter = ArrayAdapter(
                this@ViewExpensesActivity,
                android.R.layout.simple_list_item_1,
                expenseTitles
            )
            listView.adapter = adapter
        }

        exportButton.setOnClickListener {
            lifecycleScope.launch {
                val data = db.expenseDao().getAllExpenses()
                exportToExcelAndShare(data)
            }
        }
    }

    private fun exportToExcelAndShare(expenseList: List<Expense>) {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Expenses")

        // Header row
        val header: Row = sheet.createRow(0)
        header.createCell(0).setCellValue("ID")
        header.createCell(1).setCellValue("Title")
        header.createCell(2).setCellValue("Amount")
        header.createCell(3).setCellValue("Date")

        // Fill data
        for ((index, expense) in expenseList.withIndex()) {
            val row = sheet.createRow(index + 1)
            row.createCell(0).setCellValue(expense.id.toDouble())
            row.createCell(1).setCellValue(expense.title)
            row.createCell(2).setCellValue(expense.amount)
            row.createCell(3).setCellValue(expense.date)
        }

        try {
            val exportDir = File(getExternalFilesDir(null), "exports")
            if (!exportDir.exists()) exportDir.mkdirs()

            val file = File(exportDir, "Expenses.xlsx")
            val fos = FileOutputStream(file)
            workbook.write(fos)
            fos.close()

            Toast.makeText(this, "Exported to: ${file.absolutePath}", Toast.LENGTH_SHORT).show()

            val uri: Uri = FileProvider.getUriForFile(
                this,
                "${applicationContext.packageName}.fileprovider",
                file
            )

            val shareIntent = Intent(Intent.ACTION_SEND).apply {
                type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                putExtra(Intent.EXTRA_STREAM, uri)
                addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            }

            startActivity(Intent.createChooser(shareIntent, "Share Excel File"))

        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Export failed: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
}

