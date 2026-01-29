# Send-ToApacheDataset PowerShell Module

## Описание

Модуль PowerShell для загрузки данных из файлов CSV и Excel в датасет Apache Superset через REST API и multipart/form-data. Командлет `Send-ToApacheDataset` выполняет аутентификацию, проверяет наличие указанной базы данных и датасета, затем загружает данные с возможностью замены существующих записей.

## Основные возможности

- **Поддержка двух форматов**: CSV и Excel (XLS/XLSX)
- **Аутентификация в Apache Superset** через LDAP или внутреннюю базу данных
- **Проверка существования** базы данных и датасета перед загрузкой
- **Загрузка данных** с использованием multipart/form-data
- **Гибкая настройка параметров загрузки** через хэш-таблицу
- **Подтверждение успешной загрузки** через сравнение даты изменения датасета
- **Расширенная обработка ошибок** с извлечением сообщений из ответа сервера

## Требования

- **Наименования**: базы данных, датасета и листов Excel должны указываться с учетом регистра
- **Датасет**: не должен содержать русские символы
- **Кодировка CSV**: UTF-8 (без дополнительных спецификаций)
- **Разделитель по умолчанию**: запятая (,) - при необходимости изменения используйте параметр `-FormParameters`

## Установка

Скопируйте файл `Send-ToApacheDataset.psm1` в одну из директорий модулей PowerShell:
- `$env:USERPROFILE\Documents\WindowsPowerShell\Modules\Send-ToApacheDataset\`
- `C:\Program Files\WindowsPowerShell\Modules\Send-ToApacheDataset\`

## Использование

### Базовый пример загрузки CSV

```powershell
# Импорт модуля
Import-Module Send-ToApacheDataset

# Загрузка CSV-файла в датасет
$result = Send-ToApacheDataset `
    -SupersetUrl "http://superseturl" `
    -UserName "user" `
    -Password "pass" `
    -UploadFile "c:\hdd.csv" `
    -DatasetName "HDD"

# Проверка результата
if ($result.Success) {
    Write-Host $result.Message -ForegroundColor Green
} else {
    Write-Host $result.Message -ForegroundColor Red
}
```

### Пример загрузки Excel файла

```powershell
Send-ToApacheDataset `
    -SupersetUrl "http://superset.example.com" `
    -UserName "admin" `
    -Password "secure_password" `
    -UploadFile "C:\data\export.xlsx" `
    -DatasetName "Sales_Data" `
    -DatabaseName "PostgreSQL" `
    -Provider "db"
```

### Пример с пользовательскими параметрами формы

```powershell
$customParams = @{
    "schema" = "analytics"
    "if_exists" = "append"
    "sep" = ";"
    "header" = "1" 
    "mangle_dupe_cols" = "y"
    "skipinitialspace" = "y" 
}

Send-ToApacheDataset `
    -SupersetUrl "http://superset.example.com" `
    -UserName "admin" `
    -Password "secure_password" `
    -UploadFile "C:\data\export.csv" `
    -DatasetName "Sales_Data" `
    -FormParameters $customParams
```
## Параметры командлета

| Параметр | Обязательность | По умолчанию | Описание |
|----------|----------------|--------------|----------|
| `SupersetUrl` | Да | - | URL-адрес Apache Superset |
| `UserName` | Да | - | Имя пользователя для аутентификации |
| `Password` | Да | - | Пароль пользователя |
| `Provider` | Нет | "ldap" | Провайдер аутентификации ("ldap" или "db") |
| `DatasetName` | Да | - | Название датасета для загрузки |
| `UploadFile` | Да | - | Путь к CSV или Excel файлу для загрузки |
| `DatabaseName` | Нет | "PostgreSQL" | Название базы данных в Superset |
| `FormParameters` | Нет | @{...} | Хэш-таблица с параметрами формы |
| `ResultToFile` | Нет | - | Путь для сохранения ответа сервера (для отладки) |

## Параметры формы по умолчанию

```powershell
@{
    "schema" = "public"              # Указание схемы (если БД поддерживает)
    "if_exists" = "replace"          # При существовании таблицы: Fail (ничего), Replace (удалить и создать), Append (добавить данные)
    "sep" = ","                      # Разделитель в CSV файле (для пробелов используйте \s+)
    "header" = "0"                   # Строка с заголовками (0 - первая строка данных). Пусто, если нет строки заголовков
    "mangle_dupe_cols" = "y"         # Обозначать дублирующиеся колонки как "X.0, X.1"
}
```

**Важно**: При указании `-FormParameters` необходимо указывать все параметры, так как значения по умолчанию не используются.

## Дополнительные параметры формы

Для добавления дополнительных параметров формы изучите HTML-код страниц формы в Superset:
- Для CSV: `/csvtodatabaseview/form`
- Для Excel: `/exceltodatabaseview/form`

Наименования параметров содержатся в атрибуте `name` элементов формы.

## Формат результата

Командлет возвращает объект со следующими свойствами:

```powershell
@{
    Success = $true  # или $false
    Message = "Детальное сообщение о результате выполнения"
}
```

## Примеры выходных данных

### При успешной загрузке:
```
Success : True
Message : Загрузка выполнена успешно. Датасет: Sales_Data, дата изменения: 06.10.2025 15:30:25.
```

### При ошибке:
```
Success : False
Message : Не найден датасет Sales_Data
```

## Обработка ошибок

Модуль обрабатывает следующие типы ошибок:
- Файл не найден или неверный тип файла
- Ошибки аутентификации
- База данных или датасет не найдены
- Ошибки при выполнении REST-запросов
- Ошибки из ответа сервера Superset (извлекаются из HTML)

## Примечания

- Командлет автоматически определяет тип файла по расширению (.csv или .xls/.xlsx)
- Для каждого типа файла используется соответствующий эндпойт Superset
- Проверяется дата изменения датасета для подтверждения успешности загрузки
- Поддерживается расширенная отладка через сохранение ответов сервера
- Рекомендуется использовать в автоматизированных сценариях загрузки данных

## Автор

Новгородов Павел  
Дата создания: 06.2025

## Лицензия

Модуль распространяется без гарантий. Используйте на свой страх и риск.
