XlsxToCsvConverter 
===========
---
Микро-сервис для преобразования исходного Excel 2007 (xlsx) файла в csv формат,
где каждый лист конвертируется в отдельный файл типа "*имя файла*.csv" и 
помещается в zip архив.

### Опции:
*URL* - Ссылка на исходный файл (если запрос в виде json)
	
*DESIRED_SHEETS* - Строка с названиями листов (разделены ","), 
которые нужно конвертировать. По умолчанию конвертирует все листы.

*ROWS_TO_SKIP* - Количество строк от начала документа,
которые стоит пропустить. По умолчанию 0.

*COLUMNS_TO_SKIP* - Строка с номерами столбцов (через ","), которые не нужно учитывать.

*FORMAT_VALUES* - Задает, стоит ли записывать значения ячеек с учетом формата эксель,
или без него (raw values). По умолчанию записывает без учета (false). 
	
*CSV_FORMAT* - Формат выходных csv файлов. По умолчанию "CUSTOM"

>Доступные форматы (название и как записывать в кавычках): <br />
>**CUSTOM** - "custom" (индивидуальный формат, с учетом параметров, заданных пользователем.) <br />
>**RFC 4180** - "rfc-4180"   <br />
>**Microsoft Excel** - "excel"<br />
>**Tab-Delimited** - "tdf"<br />
>**MySQL Format** - "mysql"<br />
>**Informix Unload** - "informix-unload"<br />
>**Informix Unload Escape Disabled** - "informix-unload-csv"<br />

## Следующие параметры учитываются только при CSV_FORMAT = "CUSTOM"
*VALUE_SEPARATOR* - Символ-разделитель значений в csv файлах **("," по умолчанию)**<br />
*FIRST_LINE_IS_HEADER* - Параметр отвечает за то, нужно ли включать в csv файл первую строку таблицы. **По умолчанию "true".** <br />
*QUOTE_CHAR* - Символ, который используется для заключения значений в кавычки, чтобы избежать использования escape-символов. **По умолчанию '"'** <br />
*ESCAPE_CHAR* - Используется для обозначения особых символов, которые могли бы повлиять на поведения обработчика данных при считывании. **По умолчанию "\"** <br />
*COMMENT_MAKER* - Символ, обозначающий начало комментария.<br />
*NULL_STRING* - Задает строку, и если она представлена в виде значения в CSV файле,
то должна рассматриваться как пустое поле вместо использования буквального значения.<br />
*TRIM_FIELDS* - Следует ли удалять пробелы в начале и в конце полей. **По умолчанию "true"**. <br />
*QUOTE_MODE* - Указывает, как поля должны заключаться в кавычки при их записи.<br /> 
>### Возможные значения: <br />
> **_Do Not Quote Values_** - **"NONE" (По умолчанию)**. Значения не будут заключаться в кавычки.
> Вместо этого все специальные символы будут экранированы с помощью заданного escape-символа. <br />
> ***Quote All Values*** - "ALL". Все значения будут заключаться в кавычки с использованием заданного символа кавычек (quote char). <br />
> ***Quote Minimal*** - "MINIMAL". Значения будут заключены в кавычки,
> только если они содержат специальные символы, такие как символы новой строки или разделители.<br />
> ***Quote Non-Numeric Values*** - "NON_NUMERIC". Все значения будут заключены в кавычки, кроме числовых.<br />

*RECORD_SEPARATOR* - Указывает символы, используемые для разделения CSV записей. **По умолчанию "\n"** . <br />
*TRAILING_DELIMITER* - Нужно ли записывать завершающий разделитель в конец каждого файла? **По умолчанию "false"**. <br />
*ALLOW_DUPLICATE_HEADER_NAMES* - Отвечает за разрешение записывать
несколько столбцов с одинаковыми названиями. **По умолчанию "false"**. <br />

# Примеры запросов
## Конвертирование файлов по ссылке
### Без параметров
+ **POST** request `<host>:8080/json`
+ header: `Content-Type: application/json`
+ Body:
```json
{
  "URL": "https://filesamples.com/samples/document/xlsx/sample1.xlsx"
}
```
> Успешный ответ: status code 200

> Zip файл с листами, вынесенными в отдельные csv. Стандартный формат.

### С параметрами (с заданым форматом)
+ **POST** request `<host>:8080/json`
+ header: `Content-Type: application/json`
+ Body:
```json
{
  "URL":"https://filesamples.com/samples/document/xlsx/sample1.xlsx",
  "CSV_FORMAT":"Excel",
  "VALUE_SEPARATOR":";",
  "FIRST_LINE_IS_HEADER":"true"
}
```

> Успешный ответ: status code 200

>Zip файл с листами, вынесенными в отдельные csv. Так как задан формат "Excel", остальные параметры, связанные с форматом csv игнорируются.

### С параметрами (без заданого формата)
+ **POST** request `<host>:8080/json`
+ header: `Content-Type: application/json`
+ Body:
```json
{
  "URL":"https://filesamples.com/samples/document/xlsx/sample1.xlsx",
  "VALUE_SEPARATOR":";",
  "FIRST_LINE_IS_HEADER":"true",
  "COMMENT_MAKER":"//",
  "COLUMNS_TO_SKIP":"3"
}
```

>Успешный ответ: status code 200

>Zip файл с листами, вынесенными в отдельные csv. Так как формат не задан, используются все параметры.

## Конвертирование прикрепленных файлов (multipartfile)
### Файл без параметров

+ **POST** request `<host>:8080/multipart`
+ header: `Content-Type: multipart/form-data`
+ Excel multipart file.xlsx 



>Успешный ответ: status code 200 

>Zip файл с листами, вынесенными в отдельные csv. Стандартный формат.

### С параметрами (без заданого формата)

+ **POST** request `<host>:8080/multipart?VALUE_SEPARATOR=%3B&FIRST_LINE_IS_HEADER=true&COMMENT_MAKER=%2F%2F&COLUMNS_TO_SKIP=3`
+ header: `Content-Type: multipart/form-data`
+ Excel multipart file.xlsx



>Успешный ответ: status code 200

>Zip файл с листами, вынесенными в отдельные csv. Так как формат не задан, используются все заданные параметры.

### С параметрами (с заданным форматом)

+ **POST** request `<host>:8080/multipart?VALUE_SEPARATOR=%3B&FIRST_LINE_IS_HEADER=true&COMMENT_MAKER=%2F%2F&COLUMNS_TO_SKIP=3&CSV_TYPE=Excel`
+ header: `Content-Type: multipart/form-data`
+ Excel multipart file.xlsx

> Успешный ответ: status code 200

>Zip файл с листами, вынесенными в отдельные csv. Так как задан формат "Excel", остальные параметры, связанные с форматом csv игнорируются.