<#
.SYNOPSIS
    Загрузка данных из файлов CSV и Excel в датасет Apache Superset
.DESCRIPTION
    Командлет выполняет загрузку данных из файлов CSV и Excel в Dataset Apache Superset.
    Загрузка выполняется в существующий датасет.

    Требования:
        - Наименование базы данных, датасета, листа Excel - должны указываться с учетом регистра.
        - Наименование датасета не должно содержать русские символы.
        - Кодировка CSV файлов для загрузки должны быть UTF-8 (без дополнительных спецификаций).
        - По умолчанию Apache Superset используется разделитель запятая (,), если требуется указать иной разделитель, это можно сделать при указании соответсвущего значения в параметре -FormParameters.

    Перед загрузкой проверяется наличие указанной базы данных (параметр -DatabaseName) и определение ID базы данных, после чего проверяется наличие Dataset (параметр -DatasetName).

    Авторизация, получение CSRF токена, получение перечня баз данных и датасетов (в том числе дата изменения датасета) осуществляется через REST API
    (эндпойты /api/v1/security/login, /api/v1/security/csrf_token, /api/v1/database, /api/v1/dataset/)

    Загрузка данных из файлов осущестсвляется с использованием multipart/form-data.
    В зависимости от типа загружаемого файлы используются следующие адреса:
        для CSV   - /csvtodatabaseview/form
        для Excel - /exceltodatabaseview/form
    
    Результат загрузки оцениватеся путем сравнения даты изменения датасета до изменения и после загрузки.
    Командлет формирует результат работы в виде объекта содержащего поля:
        Success - результат выполнения $True\$False
        Message - сообщение

    Параметры запуска передаются в командлет непосредственно отдельными значениями:
        Параметры подключения
        -SupersetUrl  - адрес Superset
        -UserName     - имя пользователя для подключения к Superset
        -Password     - пароль пользователя для подключения к Superset
        -Provider     - провайдер аутентификации в Superset, возможные варинаты "ldap","db". По умолчанию используется значение "ldap"

        Значения параметров формы
        -DatasetName  - имя dataset в который выполняется загрузка (наименование атрибута формы "name")
        -UploadFile   - файл для загрузки в датасет (наименование атрибута формы "csv_file" или "excel_file")
        -DatabaseName - база данных в которой содержится датасет, по умолчанию "PostgreSQL" (наименование атрибута формы "con")
        
    Остальные параметры формы возможно передать через параметр -FormParameters, принимающий на вход хэш-таблицу "Наименование параметра" = "Значение параметра"
    По умолчанию значение данного параметра:
        @{
            "schema" = "public"              - Specify a schema (if database flavor supports this)
            "if_exists" = "replace"          - If table exists do one of the following: Fail (do nothing), Replace (drop and recreate table) or Append (insert data)
            "sep" = ","                      - Delimiter used by CSV file (for whitespace use \s+)
            "header" = "0"                   - Row containing the headers to use as column names (0 is first line of data). Leave empty if there is no header row
            "mangle_dupe_cols" = "y"         - Specify duplicate columns as "X.0, X.1"
         }
    В случае указания значений для -FormParameters требуется указывать все параметры, т.к. даже при указании одного параметра, значения по умолчанию использоваться не будут.

    В случае если необходимо добавить какой-либо параметр формы требуется добавить элемент хэш-таблицы с нименованием параметра и значением.
    Возможные параметры формы можно посмотреть в HTML коде страниц формы (/csvtodatabaseview/form или /exceltodatabaseview/form), наменование параметра содержится в атрибуте name, например:
                <input id="skipinitialspace" name="skipinitialspace" type="checkbox" value="y">
                <span class="help-block">Skip spaces after delimiter.</span>

.PARAMETER ProcessFor

.PARAMETER JobTimeout

.PARAMETER Protocol

.PARAMETER AppendToResult

.PARAMETER Credential

.EXAMPLE

    Send-ToApacheDataset -SupersetUrl http://superseturl -UserName user -Password pass -UploadFile c:\hdd.csv -DatasetName HDD    

.NOTES
    Новгородов Павел 06.2025
#>

function Send-ToApacheDataset {
    [cmdletbinding()]
    Param(
        #адрес Superset
        [Parameter (Mandatory=$False)]
        [string]$SupersetUrl,

        #имя пользователя для подключения к Superset
        [Parameter (Mandatory=$False)]
        [string]$UserName,

        #пароль пользователя для подключения к Superset
        [Parameter (Mandatory=$False)]
        [string]$Password,

        #провайдер проверки подлинности
        [Parameter (Mandatory=$False)]
        [ValidateSet("ldap","db")]
        [string]$Provider = "ldap",

        #имя dataset в который выполняется загрузка
        [Parameter (Mandatory=$False)]
        [string]$DatasetName,

        #файл CSV для загрузки
        [Parameter (Mandatory=$False)]
        [string]$UploadFile,

        #наименование базы данных
        [Parameter (Mandatory=$False)]
        [string]$DatabaseName = "PostgreSQL",

        #параметры формы
        [Parameter (Mandatory=$False)]
        [hashtable]$FormParameters = @{"schema" = "public"; "if_exists" = "replace"; "sep" = ","; "header" = "0"; "mangle_dupe_cols" = "y"},

        #имя файла для вывода резуальтата отправки запроса (используется для отладки) 
        [Parameter (Mandatory=$False)]
        [string]$ResultToFile

    )

    process {

    #проверяем наличие файла для загрузки
    if (!(Test-Path -Path $UploadFile))
    {
        return [PSCustomObject]@{
            Success = $False
            Message = "Файл $UploadFile не найден!"
        }
    }

    #формируем параметры в соответсвии с выбраным типом загружаемого файла
    switch -wildcard ($UploadFile) {

        "*.csv"
        {
            $FormUrl = "$SupersetUrl/csvtodatabaseview/form"           #адрес эндпойта для загрузки CSV
            $PartNameFile = "csv_file"                                 #наменованиe элемента формы с именем файла CSV
            $FileContent = [System.IO.File]::ReadAllText($UploadFile)  #содержимое файла для загрузки
            $ContentType = "text/csv"                                  #тип файла отправляемого в форме
        }

        "*.xls*"
        {
            $FormUrl = "$SupersetUrl/exceltodatabaseview/form"
            $PartNameFile = "excel_file"
            $FileBytes = [System.IO.File]::ReadAllBytes($UploadFile)
            $FileContent = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($FileBytes)
            $ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }

        default {
                    return [PSCustomObject]@{
                                                Success = $False
                                                Data = $null
                                                Error = "Тип файла должен быть CSV или XLSХ!"
                                            }
                }
    } #switch

    #получаем access_token
    $body = @{
            username = $UserName
            password = $Password
            provider = $Provider
        } | ConvertTo-Json

    try {
            $AuthResponse = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/security/login" -Method Post -Body $body -ContentType "application/json"
            $AccessToken = $AuthResponse.access_token
        }
        catch {
                return [PSCustomObject]@{
                                            Success = $False
                                            Message = "$($_.Exception.Message) Аутентификация не выполнена!"
                                        }
              }
    
    Write-Host "AccessToken:$AccessToken"

    #получаем CSRF-токен
    try {
            $CsrfResponse = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/security/csrf_token/" -Method Get -Headers @{ "Authorization" = "Bearer $AccessToken" }
            $CsrfToken = $CsrfResponse.result
        }
        catch {
                return [PSCustomObject]@{
                                            Success = $False
                                            Message = "$($_.Exception.Message) Не получен CSRF токен!"
                                        }
        }

    Write-Host "CsrfToken:$CsrfToken"

    #получение ID базы данных по ее наименованию
    $headers = @{
                    "Authorization" = "Bearer $accessToken"
                    "X-CSRFToken" = $CsrfToken
                    "Accept" = "application/json"
                }


    try {
            $response = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/database/" -Method Get -Headers $headers
            $DatabaseId = ($response.result | select database_name, id | Where-Object database_name -eq $DatabaseName).id
        }
        catch {
                return [PSCustomObject]@{
                                            Success = $False
                                            Message = "$($_.Exception.Message) Ошибка получения ID базы данных $DatabaseName"
                                        }
        }

    if (!$DatabaseId)
    {
        return [PSCustomObject]@{
            Success = $False
            Message = "Не надена база данных $DatabaseName"
        }
    }
    else
    {
        Write-Host "DatabaseID:$DatabaseId"
    }

    #проверка наличия dataset
    #формирование параметра для эндпойта /api/v1/dataset/?q
    $filter = @{
                filters = @(
                                @{
                                    col = "table_name" 
                                    opr = "eq"
                                    value = $DatasetName
                                }
                        )
                } | ConvertTo-Json -Depth 5 -Compress



    try {
            $EncodedFilter = [System.Web.HttpUtility]::UrlEncode($filter)
            $response = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/dataset/?q=$EncodedFilter" -Method Get -Headers $headers 
        }
    catch {
            return [PSCustomObject]@{
                                        Success = $False
                                        Message = "$($_.Exception.Message) Ошибка при получении датасет $DatasetName"
                                    }
          }


    if ($response.count -eq 0)
    {
        return [PSCustomObject]@{
            Success = $False
            Message ="Не найден датасет $DatasetName"
        }
    }
    else
    {
        Write-Host "Найден датасет:$DatasetName"
    }
    
    #формирование multipart/form-data запроса
    #идентификатор границы в форме
    $boundary = [System.Guid]::NewGuid().ToString()

    $LF = "`r`n"
    
    #формирование текста формы
    $BodyLines = New-Object System.Collections.ArrayList

    #добавление Csrf токена в форму
    $BodyLines.AddRange(@(
                            "--$boundary",
                            "Content-Disposition: form-data; name=`"csrf_token`"$LF",
                            "$CsrfToken"
                        ))
    #добавление наименование датсета в форму
    $BodyLines.AddRange(@(
                            "--$boundary",
                            "Content-Disposition: form-data; name=`"name`"",
                            "",
                            "$DatasetName"
                        ))

    #добавление ID базы данных в форме
    $BodyLines.AddRange(@(
                            "--$boundary",
                            "Content-Disposition: form-data; name=`"con`"$LF",
                            $DatabaseId
                        ))


    #добавление параметров формы из хэш-таблицы переданной в командлет
    foreach ($param in $FormParameters.GetEnumerator())
    {
        $BodyLines.AddRange(@(
                                "--$boundary",
                                "Content-Disposition: form-data; name=`"$($param.Key)`"$LF",
                                "$($param.Value)"
                           ))
    }

    #добавление файла в форму и закрытие текста формы
    $BodyLines.AddRange(@(
                            "--$boundary",
                            "Content-Disposition: form-data; name=`"$PartNameFile`"; filename=`"$(Split-Path $UploadFile -Leaf)`"",
                            "Content-Type: $ContentType$LF",
                            $FileContent
                            "--$boundary--$LF"
                        ))

    $body = $BodyLines -join $LF

    #заголовок
    $headers = @{
                    "Content-Type" = "multipart/form-data; boundary=$boundary"
                    "Authorization" = "Bearer $accessToken"
                    "X-CSRFToken" = $CsrfToken
                    "Referer" = $FormUrl
                }

    #получение времени изменения датасета до загрузки
    $response = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/dataset/?q=$EncodedFilter" -Method Get -Headers $headers 
    $DatasetChangedTo = [datetime]::Parse($response.result[0].changed_on_utc)

    #время начала выполнения запроса
    Write-Host "Старт загрузки файла $UploadFile в датасет $DatasetName"

    try {
            #отправляем форму с данными 
            $result = Invoke-RestMethod -Uri "$FormUrl" -Method Post -Headers $headers -ContentType "multipart/form-data; boundary=$boundary" -Body $body #| Out-Null
        }
    catch {
            return [PSCustomObject]@{
                                        Success = $False
                                        Message = $_.Exception.Message
                                    }
          }                                            

    #получение времени изменения датасета после загрузки
    $response = Invoke-RestMethod -Uri "$SupersetUrl/api/v1/dataset/?q=$EncodedFilter" -Method Get -Headers $headers 
    $DatasetChangedAfter = [datetime]::Parse($response.result[0].changed_on_utc)

    #выделяем из ответа сервера возможное сообщение об ошибке
    $pattern = '(?s)<div class="alert alert-danger">(.*?)</div>'

    if ($result -match $pattern)
    {
        Write-Host ("Вероятная проблема:")
        Write-Host ($Matches[1] -replace "<[^>].*?</button>", "").Trim()
    }
    
    #если задан параметр сохранения ответа сервера в файл, сохраняем в файл
    if ($ResultToFile)
    {
        $result | Out-File $ResultToFile
    }

    #проверка выполнения загрузки данных, если дата изменения после загрузки больше чем до загрузки
    if ($DatasetChangedAfter -gt $DatasetChangedTo -and $DatasetChangedAfter -ne $DatasetChangedTo)
    {
        return [PSCustomObject]@{
            Success = $true
            Message = "Загрузка выполнена успешно. Датасет: $DatasetName, дата изменения: $DatasetChangedAfter."
        }
    }
    else
    {
        return [PSCustomObject]@{
            Success = $False
            Message = "Загрузка в датасет $DatasetName не выполнена! Дата изменения датасета до загрузки: $DatasetChangedTo, после зарузки: $DatasetChangedAfter."
        }
    }

    } #process
} #function

