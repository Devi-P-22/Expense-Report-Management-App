package com.example.expensetracker

import android.app.Activity
import android.app.DatePickerDialog
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.widget.*
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.FileProvider
import androidx.lifecycle.lifecycleScope
import com.example.expensetracker.data.ExpenseDatabase
import com.example.expensetracker.model.Expense
import kotlinx.coroutines.launch
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.Row
import java.io.File
import java.io.FileOutputStream
import java.util.*

class AddExpenseActivity : AppCompatActivity() {

    private lateinit var categorySpinner: Spinner
    private lateinit var statusSpinner: Spinner
    private lateinit var viewBillButton: Button
    private lateinit var uploadBillButton: Button
    private lateinit var billFilePathEditText: EditText

    private var selectedFilePath: String = ""
    private val PICK_FILE_REQUEST_CODE = 101

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_add_expense)

        // Bind Views
        categorySpinner = findViewById(R.id.categorySpinner)
        statusSpinner = findViewById(R.id.statusSpinner)
        viewBillButton = findViewById(R.id.viewBillButton)
        uploadBillButton = findViewById(R.id.uploadButton)
        billFilePathEditText = findViewById(R.id.billFilePathEditText)
        val dateEditText = findViewById<EditText>(R.id.date)
        val gstSwitch = findViewById<Switch>(R.id.gstSwitch)

        // Spinner Data
        val categoryOptions = arrayOf(
            "F&B - IN OFFICE", "F&B - TOUR", "UTILITY", "MAINTENANCE",
            "OFFICE ASSET", "R&D - ELECTRONICS", "R&D - MECHANICAL",
            "PRODUCTION - ELECTRONICS", "PRODUCTION - MECHANICAL",
            "EMPLOYEE WELLBEING", "TRAVEL", "STAY", "MISC"
        )
        val statusOptions = arrayOf("DONE", "PENDING")
        categorySpinner.adapter = ArrayAdapter(this, android.R.layout.simple_spinner_dropdown_item, categoryOptions)
        statusSpinner.adapter = ArrayAdapter(this, android.R.layout.simple_spinner_dropdown_item, statusOptions)

        // Date Picker
        dateEditText.setOnClickListener {
            val c = Calendar.getInstance()
            DatePickerDialog(this, { _, y, m, d ->
                dateEditText.setText("$d/${m + 1}/$y")
            }, c.get(Calendar.YEAR), c.get(Calendar.MONTH), c.get(Calendar.DAY_OF_MONTH)).show()
        }

        // Upload Bill File
        uploadBillButton.setOnClickListener {
            val intent = Intent(Intent.ACTION_GET_CONTENT)
            intent.type = "*/*"
            startActivityForResult(Intent.createChooser(intent, "Select Bill File"), PICK_FILE_REQUEST_CODE)
        }

        // View Bill File
        viewBillButton.setOnClickListener {
            if (selectedFilePath.isNotEmpty()) {
                try {
                    val uri = Uri.parse(selectedFilePath)
                    val viewIntent = Intent(Intent.ACTION_VIEW).apply {
                        setDataAndType(uri, "*/*")
                        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                    }
                    startActivity(Intent.createChooser(viewIntent, "Open Bill File With"))
                } catch (e: Exception) {
                    e.printStackTrace()
                    Toast.makeText(this, "Cannot open file", Toast.LENGTH_SHORT).show()
                }
            } else {
                Toast.makeText(this, "No file selected", Toast.LENGTH_SHORT).show()
            }
        }

        // Submit Expense
        findViewById<Button>(R.id.submitExpense).setOnClickListener {
            lifecycleScope.launch {
                val expense = Expense(
                    title = findViewById<EditText>(R.id.expenseName).text.toString(),
                    amount = findViewById<EditText>(R.id.price).text.toString().toDoubleOrNull() ?: 0.0,
                    name = findViewById<EditText>(R.id.paidBy).text.toString(),
                    category = categorySpinner.selectedItem.toString(),
                    price = findViewById<EditText>(R.id.price).text.toString().toDoubleOrNull() ?: 0.0,
                    paidBy = findViewById<EditText>(R.id.paidBy).text.toString(),
                    status = statusSpinner.selectedItem.toString(),
                    date = dateEditText.text.toString(),
                    gstAvailable = gstSwitch.isChecked,
                    billFilePath = selectedFilePath
                )

                // Save to DB
                ExpenseDatabase.getDatabase(this@AddExpenseActivity).expenseDao().insertExpense(expense)
                Toast.makeText(this@AddExpenseActivity, "Expense Submitted Successfully!", Toast.LENGTH_SHORT).show()

                // Export All Expenses
                val expenses = ExpenseDatabase.getDatabase(this@AddExpenseActivity).expenseDao().getAllExpenses()
                exportToExcel(expenses)
            }
        }
    }

    // File Picker Result
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == PICK_FILE_REQUEST_CODE && resultCode == Activity.RESULT_OK) {
            val uri = data?.data
            if (uri != null) {
                selectedFilePath = uri.toString()
                billFilePathEditText.setText(selectedFilePath)
                Toast.makeText(this, "File Selected", Toast.LENGTH_SHORT).show()
            } else {
                Toast.makeText(this, " File not selected", Toast.LENGTH_SHORT).show()
            }
        }
    }


    private fun exportToExcel(expenseList: List<Expense>) {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Expenses")

        val header: Row = sheet.createRow(0)
        header.createCell(0).setCellValue("ID")
        header.createCell(1).setCellValue("Title")
        header.createCell(2).setCellValue("Amount")
        header.createCell(3).setCellValue("Date")

        for ((index, expense) in expenseList.withIndex()) {
            val row: Row = sheet.createRow(index + 1)
            row.createCell(0).setCellValue(expense.id.toDouble())
            row.createCell(1).setCellValue(expense.title)
            row.createCell(2).setCellValue(expense.amount)
            row.createCell(3).setCellValue(expense.date)
        }

        try {
            val downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
            if (!downloadsDir.exists()) downloadsDir.mkdirs()

            val file = File(downloadsDir, "Expenses.xlsx")
            val fos = FileOutputStream(file)
            workbook.write(fos)
            fos.close()

            Toast.makeText(this, "Excel exported to: ${file.absolutePath}", Toast.LENGTH_LONG).show()

            val uri = FileProvider.getUriForFile(
                this,
                "${applicationContext.packageName}.fileprovider",
                file
            )

            val viewIntent = Intent(Intent.ACTION_VIEW).apply {
                setDataAndType(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            }

            startActivity(Intent.createChooser(viewIntent, "Open Excel File"))

        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Export failed: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
}

