<?php
function word(){
// Подключение к базе данных PostgreSQL
$host = "localhost";
$port = "5432";
$dbname = "mydb";
$user = "myuser";
$password = "mypassword";

$dbconn = pg_connect("host={$host} port={$port} dbname={$dbname} user={$user} password={$password}");

// Выполнение запроса к базе данных
$query = "SELECT \"Products\".\"Name\" as Товар,\"Products\".\"Group\"as Группа,\"Products\".\"Fat_content\"as Жирность,LOWER(Concat(\"Products\".\"Weight\",' ',\"Products\".\"Ed.measurements\"))as Вес,\"Supplier\".\"Name\"as Поставщик,\"Products\".\"Price\"as Цена,(\"Products\".\"Remains\" || ' шт.') as Всего  FROM \"Products\" inner join \"Supplier\" on \"Products\".\"Id_supplier\"=\"Supplier\".\"Id_supplier\"";
$result = pg_query($dbconn, $query);

// Создание нового документа Word
$wordApp = new COM("Word.Application");
$doc = $wordApp->Documents->Add();

// Установка ширины страницы на A4 и уменьшение полей до 1 см
$doc->PageSetup->PaperSize = 9; // wdPaperA4 = 9
$doc->PageSetup->LeftMargin = $wordApp->CentimetersToPoints(1);
$doc->PageSetup->RightMargin = $wordApp->CentimetersToPoints(1);
$doc->PageSetup->TopMargin = $wordApp->CentimetersToPoints(1);
$doc->PageSetup->BottomMargin = $wordApp->CentimetersToPoints(1);

// Добавление заголовка
$title = $doc->Paragraphs->Add();
$title->Range->Text = "Ассортимент молочной продукции";
$title->Range->Font->Size = 16;
$title->Range->Font->Bold = 1;
$title->Range->ParagraphFormat->Alignment = 1; // wdAlignParagraphCenter = 1
$title->Range->InsertParagraphAfter();

// Получение данных из результата запроса и создание таблицы
$num_rows = pg_num_rows($result);
$num_cols = pg_num_fields($result);
$table = $doc->Tables->Add($doc->Paragraphs[2]->Range, $num_rows + 1, $num_cols, 1, 2); // wdWord9TableBehavior = 1, wdAutoFitWindow = 2
$table->Borders->InsideLineStyle = 1; // wdLineStyleSingle = 1
$table->Borders->OutsideLineStyle = 1;
$table->AllowAutoFit = true;

// Заполнение заголовков таблицы
for ($i = 0; $i < $num_cols; $i++) {
    $table->Cell(1, $i + 1)->Range->Text = pg_field_name($result, $i);
    $table->Cell(1, $i + 1)->Range->Font->Bold = 1;
}

// Заполнение таблицы данными из результата запроса
$row = 2;
while ($data = $result->fetch_assoc()) {
    $col = 1;
    foreach ($data as $value) {
        $table->Cell($row, $col)->Range->Text = $value;
        $col++;
    }
    $row++;
}

// Форматирование текста в таблице
$table->Range->Font->Size = 12;
$table->Range->ParagraphFormat->Alignment = 1;
$table->Rows[1]->Range->Font->ColorIndex = 1;
$table->Rows[1]->Range->Shading->BackgroundPatternColorIndex = 13;

// Сохранение документа Word в файле
$doc->SaveAs("C:\\Users\\lukya\\OneDrive\\Рабочий стол\\Курсовая работа по БД\\Delivery\\Documents\\catalog.docx");
$doc->Close();
$wordApp->Quit();
// Закрытие соединения с базой данных
$conn->close();
}
?>