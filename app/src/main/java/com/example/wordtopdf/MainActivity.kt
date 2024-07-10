package com.example.wordtopdf

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts.GetContent
import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.FileProvider
import com.example.wordtopdf.ui.theme.WordToPdfTheme
import com.itextpdf.kernel.pdf.PdfWriter
import com.itextpdf.layout.Document
import com.itextpdf.layout.element.Paragraph
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.InputStream

class MainActivity : ComponentActivity() {
    @OptIn(ExperimentalMaterial3Api::class)
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            WordToPdfTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = Color(0xff80AF81)
                ) {
                    Scaffold(
                        modifier = Modifier.fillMaxSize(),
                        containerColor = Color(0xff80AF81),
                        topBar = {
                            TopAppBar(title = {
                                Text(
                                    text = "Word to PDF Converter",
                                    fontSize = 24.sp,
                                    modifier = Modifier.padding(bottom = 16.dp),
                                    fontWeight = FontWeight.Bold
                                )
                            },
                            colors = TopAppBarDefaults.topAppBarColors(Color(0xffFFB1B1)))
                        }
                    ) {
                        WordToPdfApp()
                    }
                }
            }
        }
    }
}

@Composable
fun WordToPdfApp() {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    var status by remember { mutableStateOf("") }
    var pdfUri by remember { mutableStateOf<Uri?>(null) }
    var pdfFileName by remember { mutableStateOf("") }

    val filePickerLauncher = rememberLauncherForActivityResult(contract = GetContent()) { uri: Uri? ->
        uri?.let {
            val fileName = getFileNameFromUri(context, it)
            scope.launch(Dispatchers.IO) {
                val result = convertWordToPdf(context, it, fileName)
                pdfFileName = fileName.replace(".docx", ".pdf")
                status = if (result) {
                    pdfUri = getPdfUri(context, pdfFileName)
                    "Conversion Successful"
                } else {
                    pdfUri = null
                    "Conversion Failed"
                }
            }
        }
    }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.Center,
        horizontalAlignment = Alignment.CenterHorizontally
    ) {

        Button(onClick = { filePickerLauncher.launch("application/vnd.openxmlformats-officedocument.wordprocessingml.document") }) {
            Text(text = "Select Word File")
        }
        Spacer(modifier = Modifier.height(16.dp))
        Text(text = status, fontSize = 18.sp)
        Spacer(modifier = Modifier.height(16.dp))
        pdfUri?.let { uri ->
            Button(onClick = { downloadPdf(context, uri, pdfFileName) }) {
                Text(text = "Download PDF")
            }
        }
    }
}

fun getFileNameFromUri(context: Context, uri: Uri): String {
    var fileName = "document.pdf"
    val cursor = context.contentResolver.query(uri, null, null, null, null)
    cursor?.use {
        if (it.moveToFirst()) {
            val nameIndex = it.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
            if (nameIndex != -1) {
                fileName = it.getString(nameIndex)
            }
        }
    }
    return fileName
}

fun convertWordToPdf(context: Context, uri: Uri, fileName: String): Boolean {
    return try {
        val inputStream: InputStream? = context.contentResolver.openInputStream(uri)
        val document = XWPFDocument(inputStream)
        val outputFile = File(context.cacheDir, fileName.replace(".docx", ".pdf"))
        val pdfWriter = PdfWriter(outputFile)
        val pdfDoc = com.itextpdf.kernel.pdf.PdfDocument(pdfWriter)
        val documentPdf = Document(pdfDoc)

        document.paragraphs.forEach { paragraph ->
            documentPdf.add(Paragraph(paragraph.text))
        }

        documentPdf.close()
        pdfDoc.close()
        pdfWriter.close()
        inputStream?.close()
        true
    } catch (e: Exception) {
        e.printStackTrace()
        false
    }
}

fun getPdfUri(context: Context, fileName: String): Uri? {
    val pdfFile = File(context.cacheDir, fileName)
    return if (pdfFile.exists()) {
        FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", pdfFile)
    } else {
        null
    }
}

fun downloadPdf(context: Context, uri: Uri, fileName: String) {
    val intent = Intent(Intent.ACTION_SEND).apply {
        type = "application/pdf"
        putExtra(Intent.EXTRA_STREAM, uri)
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        putExtra(Intent.EXTRA_TITLE, fileName)
    }
    context.startActivity(Intent.createChooser(intent, "Download PDF"))
}