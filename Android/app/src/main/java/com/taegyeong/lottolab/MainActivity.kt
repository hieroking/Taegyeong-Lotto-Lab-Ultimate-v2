package com.taegyeong.lottolab

import android.content.Context
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.google.mlkit.vision.common.InputImage
import com.google.mlkit.vision.text.TextRecognition
import com.google.mlkit.vision.text.korean.KoreanTextRecognizerOptions
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.tasks.await
import kotlinx.coroutines.withContext
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.util.concurrent.ConcurrentHashMap
import kotlin.math.abs

data class Draw(
    val round: Int,
    val numbers: List<Int>,
    val bonus: Int?
)

data class Recommendation(
    val numbers: List<Int>,
    val categoryScore: Double,
    val totalScore: Double,
    val inputScore: Double,
    val pairScore: Double,
    val tripleScore: Double,
    val recentScore: Double,
    val pairText: String
)

enum class Category(val label: String) {
    RECOMMEND("추천조합"),
    INPUT("나온횟수"),
    PAIR("동반수"),
    TRIPLE("트리플"),
    RECENT("최근패턴"),
    SELF("자체추천")
}

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            LottoTheme {
                LottoApp(applicationContext)
            }
        }
    }
}

@Composable
fun LottoTheme(content: @Composable () -> Unit) {
    val colors = darkColorScheme(
        primary = Color(0xFFD4AF37),
        secondary = Color(0xFFE8D9A7),
        background = Color(0xFF111111),
        surface = Color(0xFF1A1A1A),
        onPrimary = Color(0xFF111111),
        onBackground = Color(0xFFF4F0E6),
        onSurface = Color(0xFFF4F0E6),
    )
    MaterialTheme(colorScheme = colors, content = content)
}

@Composable
fun LottoApp(context: Context) {
    val scope = rememberCoroutineScope()
    var draws by remember { mutableStateOf<List<Draw>>(emptyList()) }
    var inputText by remember { mutableStateOf("") }
    var selected by remember { mutableStateOf(Category.RECOMMEND) }
    var results by remember { mutableStateOf<List<Recommendation>>(emptyList()) }
    var status by remember { mutableStateOf("역대 Excel을 먼저 불러오세요.") }
    var busy by remember { mutableStateOf(false) }

    val excelLauncher = rememberLauncherForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri ->
        if (uri != null) {
            scope.launch {
                busy = true
                status = "Excel 분석 중..."
                runCatching {
                    withContext(Dispatchers.IO) { parseExcel(context, uri) }
                }.onSuccess {
                    draws = it
                    status = "Excel ${it.size}회 분석 완료. 자체추천 계산 중..."
                    selected = Category.SELF
                    results = withContext(Dispatchers.Default) {
                        recommend(it, emptyMap(), Category.SELF)
                    }
                    status = "자체추천 ${results.size}조합 계산 완료"
                }.onFailure {
                    status = "Excel 오류: ${it.message}"
                }
                busy = false
            }
        }
    }

    val photoLauncher = rememberLauncherForActivityResult(
        ActivityResultContracts.OpenMultipleDocuments()
    ) { uris ->
        if (uris.isNotEmpty()) {
            scope.launch {
                busy = true
                status = "사진 OCR 처리 중..."
                runCatching {
                    recognizePhotos(context, uris)
                }.onSuccess { numbers ->
                    val added = numbers.joinToString(" ")
                    inputText = listOf(inputText.trim(), added)
                        .filter { it.isNotBlank() }
                        .joinToString("\n")
                    status = "사진에서 숫자 ${numbers.size}개 인식 완료"
                    if (draws.isNotEmpty() && inputCounts(inputText).size >= 6) {
                        selected = Category.RECOMMEND
                        results = withContext(Dispatchers.Default) {
                            recommend(draws, inputCounts(inputText), selected)
                        }
                        status = "추천조합 ${results.size}개 계산 완료"
                    }
                }.onFailure {
                    status = "사진 OCR 오류: ${it.message}"
                }
                busy = false
            }
        }
    }

    fun calculate(category: Category) {
        selected = category
        scope.launch {
            if (draws.isEmpty()) {
                status = "역대 Excel을 먼저 불러오세요."
                return@launch
            }
            val counts = inputCounts(inputText)
            if (category != Category.SELF && counts.size < 6) {
                status = "${category.label}: 고유 번호 6개 이상이 필요합니다."
                results = emptyList()
                return@launch
            }
            busy = true
            status = "${category.label} 계산 중..."
            results = withContext(Dispatchers.Default) {
                recommend(draws, counts, category)
            }
            status = "${category.label} ${results.size}조합 계산 완료"
            busy = false
        }
    }

    Column(
        Modifier
            .fillMaxSize()
            .background(MaterialTheme.colorScheme.background)
            .padding(12.dp)
    ) {
        Text(
            "太炅 Lotto Lab",
            fontSize = 28.sp,
            fontWeight = FontWeight.Bold,
            color = MaterialTheme.colorScheme.primary
        )
        Text("Android v1.0", color = MaterialTheme.colorScheme.secondary)

        Spacer(Modifier.height(10.dp))

        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
            Button(
                onClick = { excelLauncher.launch(arrayOf(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.ms-excel"
                )) },
                enabled = !busy
            ) { Text("역대 Excel") }

            Button(
                onClick = { photoLauncher.launch(arrayOf("image/*")) },
                enabled = !busy
            ) { Text("사진 OCR") }
        }

        Spacer(Modifier.height(8.dp))

        OutlinedTextField(
            value = inputText,
            onValueChange = { inputText = it },
            modifier = Modifier.fillMaxWidth().height(110.dp),
            label = { Text("번호 직접 입력·OCR 결과 수정") },
            placeholder = { Text("예: 3 6 11 24 37 42") }
        )

        Spacer(Modifier.height(8.dp))

        val scroll = rememberScrollState()
        Row(
            Modifier.horizontalScroll(scroll),
            horizontalArrangement = Arrangement.spacedBy(6.dp)
        ) {
            Category.entries.forEach { category ->
                FilterChip(
                    selected = selected == category,
                    onClick = { calculate(category) },
                    label = { Text(category.label) }
                )
            }
        }

        Spacer(Modifier.height(6.dp))
        Text(
            status,
            color = if (status.contains("오류")) Color(0xFFFF8A80)
            else MaterialTheme.colorScheme.secondary
        )
        if (busy) {
            LinearProgressIndicator(Modifier.fillMaxWidth())
        }

        Spacer(Modifier.height(8.dp))
        RecommendationList(results)
    }
}

@Composable
fun RecommendationList(results: List<Recommendation>) {
    if (results.isEmpty()) {
        Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
            Text("계산 결과가 여기에 표시됩니다.")
        }
        return
    }

    LazyColumn(verticalArrangement = Arrangement.spacedBy(7.dp)) {
        itemsIndexed(results) { index, r ->
            Card(
                colors = CardDefaults.cardColors(containerColor = Color(0xFF1A1A1A)),
                shape = RoundedCornerShape(10.dp)
            ) {
                Column(Modifier.fillMaxWidth().padding(12.dp)) {
                    Text(
                        "${index + 1}위  ${r.numbers.joinToString(" · ")}",
                        fontWeight = FontWeight.Bold,
                        fontSize = 18.sp,
                        color = when {
                            r.categoryScore >= 70 -> Color(0xFF7CFF8A)
                            r.categoryScore >= 50 -> Color(0xFFFFD95A)
                            else -> Color(0xFFF4F0E6)
                        }
                    )
                    Text(
                        "항목 ${"%.1f".format(r.categoryScore)} · 종합 ${"%.1f".format(r.totalScore)} · 합계 ${r.numbers.sum()}",
                        fontSize = 13.sp
                    )
                    Text(
                        "나온횟수 ${r.inputScore.toInt()} · 동반수 ${r.pairScore.toInt()} · 트리플 ${r.tripleScore.toInt()} · 최근 ${r.recentScore.toInt()}",
                        fontSize = 12.sp,
                        color = Color(0xFFD0D0D0)
                    )
                    Text(
                        r.pairText,
                        fontSize = 12.sp,
                        color = MaterialTheme.colorScheme.primary
                    )
                }
            }
        }
    }
}

fun inputCounts(text: String): Map<Int, Int> =
    Regex("""\d{1,2}""")
        .findAll(text)
        .map { it.value.toInt() }
        .filter { it in 1..45 }
        .groupingBy { it }
        .eachCount()

suspend fun recognizePhotos(context: Context, uris: List<Uri>): List<Int> {
    val recognizer = TextRecognition.getClient(
        KoreanTextRecognizerOptions.Builder().build()
    )
    val numbers = mutableListOf<Int>()
    try {
        for (uri in uris) {
            val image = InputImage.fromFilePath(context, uri)
            val text = recognizer.process(image).await().text
                .replace(Regex("""\b\d{1,2}:\d{2}\b"""), " ")
                .replace(Regex("""\b\d+\s*/\s*\d+\b"""), " ")
            numbers += Regex("""(?<!\d)\d{1,2}(?!\d)""")
                .findAll(text)
                .map { it.value.toInt() }
                .filter { it in 1..45 }
                .toList()
        }
    } finally {
        recognizer.close()
    }
    return numbers
}

fun parseExcel(context: Context, uri: Uri): List<Draw> {
    context.contentResolver.openInputStream(uri).use { input ->
        requireNotNull(input) { "Excel 파일을 열 수 없습니다." }
        XSSFWorkbook(input).use { workbook ->
            var best = emptyList<Draw>()
            for (sheetIndex in 0 until workbook.numberOfSheets) {
                val sheet = workbook.getSheetAt(sheetIndex)
                var headerRowIndex = -1
                for (r in 0..minOf(sheet.lastRowNum, 30)) {
                    val row = sheet.getRow(r) ?: continue
                    if ((0 until row.lastCellNum).any {
                            row.getCell(it)?.toString()?.contains("회차") == true
                        }) {
                        headerRowIndex = r
                        break
                    }
                }
                if (headerRowIndex < 0) continue

                val header = sheet.getRow(headerRowIndex)
                val headers = (0 until header.lastCellNum).map {
                    header.getCell(it)?.toString()?.trim().orEmpty()
                }

                fun findHeader(vararg keys: String): Int =
                    headers.indexOfFirst { h -> keys.any { h.contains(it, ignoreCase = true) } }

                val roundCol = findHeader("회차", "round")
                val bonusCol = findHeader("보너스", "bonus")
                val orderWords = listOf("첫번째", "두번째", "세번째", "네번째", "다섯번째", "여섯번째")
                var numberCols = orderWords.map { word ->
                    headers.indexOfFirst { it.contains(word) }
                }.filter { it >= 0 }

                if (numberCols.size < 6) {
                    numberCols = (1..6).mapNotNull { n ->
                        headers.indexOfFirst {
                            Regex("""(번호|num|ball)\s*$n$""", RegexOption.IGNORE_CASE).containsMatchIn(it)
                        }.takeIf { it >= 0 }
                    }
                }
                if (roundCol < 0 || numberCols.size < 6) continue

                val rows = mutableListOf<Draw>()
                for (r in headerRowIndex + 1..sheet.lastRowNum) {
                    val row = sheet.getRow(r) ?: continue
                    val round = numericCell(row.getCell(roundCol))?.toInt() ?: continue
                    val nums = numberCols.take(6).mapNotNull {
                        numericCell(row.getCell(it))?.toInt()
                    }.sorted()
                    if (nums.size != 6 || nums.toSet().size != 6 || nums.any { it !in 1..45 }) continue
                    val bonus = if (bonusCol >= 0) numericCell(row.getCell(bonusCol))?.toInt() else null
                    rows += Draw(round, nums, bonus?.takeIf { it in 1..45 })
                }
                if (rows.size > best.size) best = rows
            }
            require(best.isNotEmpty()) { "당첨번호 6개를 인식하지 못했습니다." }
            return best.sortedBy { it.round }
        }
    }
}

fun numericCell(cell: org.apache.poi.ss.usermodel.Cell?): Double? {
    if (cell == null) return null
    return when (cell.cellType) {
        CellType.NUMERIC -> cell.numericCellValue
        CellType.STRING -> cell.stringCellValue.trim().toDoubleOrNull()
        CellType.FORMULA -> runCatching { cell.numericCellValue }.getOrNull()
        else -> null
    }
}

fun recommend(
    draws: List<Draw>,
    input: Map<Int, Int>,
    category: Category
): List<Recommendation> {
    val numberCount = IntArray(46)
    val recentCount = IntArray(46)
    val pairCount = ConcurrentHashMap<Pair<Int, Int>, Int>()
    val recentPair = ConcurrentHashMap<Pair<Int, Int>, Int>()
    val tripleCount = ConcurrentHashMap<Triple<Int, Int, Int>, Int>()
    val first = mutableSetOf<List<Int>>()
    val second = mutableSetOf<List<Int>>()

    fun pairs(nums: List<Int>): List<Pair<Int, Int>> =
        nums.indices.flatMap { i ->
            (i + 1 until nums.size).map { j -> nums[i] to nums[j] }
        }

    fun triples(nums: List<Int>): List<Triple<Int, Int, Int>> =
        nums.indices.flatMap { i ->
            (i + 1 until nums.size).flatMap { j ->
                (j + 1 until nums.size).map { k ->
                    Triple(nums[i], nums[j], nums[k])
                }
            }
        }

    draws.forEach { draw ->
        draw.numbers.forEach { numberCount[it]++ }
        pairs(draw.numbers).forEach { pairCount[it] = (pairCount[it] ?: 0) + 1 }
        triples(draw.numbers).forEach { tripleCount[it] = (tripleCount[it] ?: 0) + 1 }
        first += draw.numbers
        draw.bonus?.let { b ->
            draw.numbers.indices.forEach { remove ->
                second += (draw.numbers.filterIndexed { idx, _ -> idx != remove } + b).sorted()
            }
        }
    }
    draws.takeLast(100).forEach { draw ->
        draw.numbers.forEach { recentCount[it]++ }
        pairs(draw.numbers).forEach { recentPair[it] = (recentPair[it] ?: 0) + 1 }
    }

    val source = if (category == Category.SELF) {
        (1..45).associateWith { maxOf(1, numberCount[it]) }
    } else input

    require(source.size >= 6) { "고유 번호 6개 이상이 필요합니다." }

    val pool = selectPool(source, category, numberCount, recentCount, pairCount, tripleCount)
    val maxInput = source.values.maxOrNull()?.toDouble() ?: 1.0
    val maxPair = pairCount.values.maxOrNull()?.toDouble() ?: 1.0
    val maxTriple = tripleCount.values.maxOrNull()?.toDouble() ?: 1.0
    val maxRecent = recentCount.maxOrNull()?.toDouble() ?: 1.0
    val maxRecentPair = recentPair.values.maxOrNull()?.toDouble() ?: 1.0

    val candidates = mutableListOf<Recommendation>()
    combinations6(pool).forEach { combo ->
        val sum = combo.sum()
        if (sum !in 20..300) return@forEach
        val odd = combo.count { it % 2 == 1 }
        val high = combo.count { it >= 23 }
        if (odd !in 2..4 || high !in 2..4) return@forEach
        if (combo.zipWithNext().count { it.second - it.first == 1 } > 2) return@forEach
        if (combo in first || combo in second) return@forEach

        val comboPairs = pairs(combo)
        val comboTriples = triples(combo)

        val inputScore = (combo.sumOf { source[it] ?: 0 } / (maxInput * 6) * 100).coerceIn(0.0, 100.0)
        val pairValues = comboPairs.map { pairCount[it]?.toDouble() ?: 0.0 }.sortedDescending()
        val pairScore = (pairValues.take(5).sum() / (maxPair * 5) * 100).coerceIn(0.0, 100.0)
        val tripleValues = comboTriples.map { tripleCount[it]?.toDouble() ?: 0.0 }.sortedDescending()
        val tripleScore = (tripleValues.take(5).sum() / (maxTriple * 5) * 100).coerceIn(0.0, 100.0)
        val recentScore = (
            combo.sumOf { recentCount[it] } / (maxRecent * 6) * 55 +
            comboPairs.sumOf { recentPair[it] ?: 0 } / (maxRecentPair * 15) * 45
        ).coerceIn(0.0, 100.0)

        var structure = 100.0
        structure -= abs(odd - 3) * 12
        structure -= abs(high - 3) * 10
        structure = structure.coerceIn(0.0, 100.0)

        val totalScore = inputScore * .30 + pairScore * .25 + tripleScore * .20 + recentScore * .15 + structure * .10
        val categoryScore = when (category) {
            Category.RECOMMEND -> totalScore
            Category.INPUT -> inputScore * .70 + totalScore * .30
            Category.PAIR -> pairScore * .70 + totalScore * .30
            Category.TRIPLE -> tripleScore * .70 + totalScore * .30
            Category.RECENT -> recentScore * .70 + totalScore * .30
            Category.SELF -> totalScore
        }

        val pairText = comboPairs
            .sortedByDescending { pairCount[it] ?: 0 }
            .take(3)
            .joinToString(", ") { "${it.first}↔${it.second} ${pairCount[it] ?: 0}회" }

        candidates += Recommendation(
            combo, categoryScore, totalScore, inputScore,
            pairScore, tripleScore, recentScore, pairText
        )
    }

    return candidates.sortedByDescending { it.categoryScore }.take(100)
}

fun selectPool(
    source: Map<Int, Int>,
    category: Category,
    all: IntArray,
    recent: IntArray,
    pairs: Map<Pair<Int, Int>, Int>,
    triples: Map<Triple<Int, Int, Int>, Int>
): List<Int> {
    if (source.size <= 24) return source.keys.sorted()

    val nums = source.keys
    val maxInput = source.values.maxOrNull()?.toDouble() ?: 1.0
    val maxRecent = nums.maxOf { recent[it] }.toDouble().coerceAtLeast(1.0)
    val pairCentral = nums.associateWith { n ->
        nums.filter { it != n }.sumOf { o -> pairs[minOf(n,o) to maxOf(n,o)] ?: 0 }
    }
    val tripleCentral = nums.associateWith { n ->
        triples.filterKeys { t -> t.first == n || t.second == n || t.third == n }.values.sum()
    }
    val maxPair = pairCentral.values.maxOrNull()?.toDouble() ?: 1.0
    val maxTriple = tripleCentral.values.maxOrNull()?.toDouble() ?: 1.0

    return nums.sortedByDescending { n ->
        val i = (source[n] ?: 0) / maxInput * 100
        val p = (pairCentral[n] ?: 0) / maxPair * 100
        val t = (tripleCentral[n] ?: 0) / maxTriple * 100
        val r = recent[n] / maxRecent * 100
        when (category) {
            Category.INPUT -> i*.75 + p*.10 + r*.15
            Category.PAIR -> p*.75 + i*.15 + r*.10
            Category.TRIPLE -> t*.75 + p*.15 + i*.10
            Category.RECENT -> r*.75 + p*.15 + i*.10
            else -> i*.30 + p*.25 + t*.20 + r*.25
        }
    }.take(24).sorted()
}

fun combinations6(values: List<Int>): Sequence<List<Int>> = sequence {
    val n = values.size
    for (a in 0 until n-5)
        for (b in a+1 until n-4)
            for (c in b+1 until n-3)
                for (d in c+1 until n-2)
                    for (e in d+1 until n-1)
                        for (f in e+1 until n)
                            yield(listOf(values[a], values[b], values[c], values[d], values[e], values[f]))
}
