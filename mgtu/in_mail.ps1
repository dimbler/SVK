param (
    $Folder_IN = "folder_to_inbox", #Сетевая папка для сохранения
    $Net_User = "login",	#Имя сетевого пользователя
    $Net_Password = "password", #Пароль сетевого пользователя
    $SQLServer = "SQL Server", #ИМя/адрес MS SQL сервера
    $SQLDatabase = "SQL DATABASE", #Имя базы данных на сервере
    $SQLLogin = "SQL LOGIN",	#Имя пользователя MS SQL
    $SQLPassword = "SQL PASSWORD", #Пароль пользователя MS SQL
    $arj_path = "\arj32.exe", #Путь к файлу архиватора
    $Net_Path = "c:\Windows\System32\net.exe", #Путь к консольной программе net.exe
    $Net_Share = "net_share" #Путь к корню сетевого ресурса
)


$global:VerbosePreference = "Continue"
$program_path = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

#Функция для отправки сообщений из скрипта
function Pandion_Send ($Body) {
	Write-Verbose $Body
}

#Cоздаем временный каталог
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}

#Процедура распаковки из архива arj
function Extract_ARJ ($InputArchiveName){
    $argc = "X -y"
    $arj_command = $program_path + $arj_path
    If (Test-Path $arj_command){
        If (Test-Path $InputArchiveName){
            $archived_files = Get-ChildItem -Filter *.arj -Path $InputArchiveName
            foreach ($file in $archived_files){
                #Если это многотомный архив
                $test = (([io.fileinfo]$file).DirectoryName + '\' + ([io.fileinfo]$file).BaseName + ".a01")
                If (Test-Path  (([io.fileinfo]$file).DirectoryName + '\' + ([io.fileinfo]$file).BaseName + ".a01")){
                    $argc += " -v"
                }
                $extract_folder = New-TemporaryDirectory
                $archive_folder = ([io.fileinfo]$file).DirectoryName
                $argc += " " + ([io.fileinfo]$file).FullName
                Write-Verbose $argc
                $process = Start-Process -FilePath $arj_command -windowstyle Hidden -ArgumentList $argc -WorkingDirectory $extract_folder -PassThru -Wait
                $result = $process.ExitCode
                If ($result -eq 0){
                    return $extract_folder              
                }else{
                    return ("Ошибка распаковки файла $InputFileName")
                }
            }
        }else{
            return ("Файл для распаковки не найден: $InputFileName")
        }
    }else{
        return ("Отсутствует архиватор по пути: $arj_path")
    }
}

Function SelectUnansweredFiles ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword){
    $waiting_files = @{}
    $waiting_emails = @{}
    $waiting_pfr = @{}
    $waiting_pfr_emails = @{}
    
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

        #Set Connection String
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.open()
    
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        #Записываем новый статус в поле
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 600000
        $SqlCmd.CommandText = "SELECT [uniq], [file_name], [FIOOI] FROM [proto].[dbo].[R_mifns] WHERE file_otvet_name IS NULL AND [data_otpr] > CAST(dateadd(day,datediff(day,14,GETDATE()),0) AS date)"
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        If ($DataSet.Tables[0].Rows.Count -gt 0){
            Foreach ($row in $DataSet.Tables[0].Rows){ 
                If (-Not $waiting_files.ContainsKey($row['file_name'].ToString())) {$waiting_files.Add(($row['file_name'].ToString().Trim()).Substring(3), $row['uniq'].ToString()) }
                If (-Not $waiting_emails.ContainsKey($row['file_name'].ToString())) {$waiting_emails.Add(($row['file_name'].ToString().Trim()).Substring(3), $row['FIOOI'].ToString().Trim()) }
            }
        }

        $SqlCmd1 = New-Object System.Data.SqlClient.SqlCommand
        #Записываем новый статус в поле
        $SqlCmd1.Connection = $SqlConnection
        $SqlCmd1.CommandTimeout = 600000
        $SqlCmd1.CommandText = "SELECT [uniq], [file_name], [FIOOI] FROM [proto].[dbo].[R_mifns] WHERE file_otvet_name IS NOT NULL AND [data_otpr] > CAST(dateadd(day,datediff(day,14,GETDATE()),0) AS date)"
        $SqlAdapter1 = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter1.SelectCommand = $SqlCmd1
        $DataSet1 = New-Object System.Data.DataSet
        $SqlAdapter1.Fill($DataSet1)
        If ($DataSet1.Tables[0].Rows.Count -gt 0){
            Foreach ($row1 in $DataSet1.Tables[0].Rows){ 
                If (-Not $waiting_pfr.ContainsKey($row1['file_name'].ToString())) {$waiting_pfr.Add(($row1['file_name'].ToString().Trim()).Substring(3), $row1['uniq'].ToString()) }
                If (-Not $waiting_pfr_emails.ContainsKey($row1['file_name'].ToString())) {$waiting_pfr_emails.Add(($row1['file_name'].ToString().Trim()).Substring(3), $row1['FIOOI'].ToString().Trim()) }
            }
        }
        return $waiting_files, $waiting_emails, $waiting_pfr, $waiting_pfr_emails
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ( "Ошибка чтения на SQL {0} для поиска неотвеченных файлов файла" -f $_.Exception.Message )
    }
    finally
    {
        $SqlAdapter.Dispose()
        $SqlAdapter1.Dispose()
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }
        
}

Function Receive311P ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $file_name, $temp_file, $waiting_files, $waiting_emails, $waiting_pfr, $waiting_pfr_emails){
    #Подключаемся к SQL серверу
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
    $SqlConnection.Open()
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.Connection = $SqlConnection
    $SqlCmd.CommandTimeout = 600000

    #Распаковываем архив если это архив
    If ([System.IO.Path]::GetExtension($file_name) -eq ".arj"){
        $fileToExtract = (Join-Path $env:TEMP $file_name)
        Move-Item $temp_file $fileToExtract -Force
        $status_extract_folder = Extract_ARJ($fileToExtract)
        If ((Test-Path $status_extract_folder) -and (-Not [string]::IsNullOrEmpty($status_extract_folder))){
            Get-ChildItem $status_extract_folder -Force | Foreach-Object {
                Write-Verbose $_.Name
                #Распаковываем архив если это архив
                If ([System.IO.Path]::GetExtension($_.FullName) -eq ".arj"){

                    #Имя файла ответа от ФНС
                    $mifns_filename = $_.Name

                    #Распаковываем вложенный архив
                    $status_extract_subfolder = Extract_ARJ($_.FullName)
                    If ((Test-Path $status_extract_subfolder) -and (-Not [string]::IsNullOrEmpty($status_extract_subfolder))){
                        Get-ChildItem $status_extract_subfolder -Force | Foreach-Object {

                            #Разбираем файлы ответов от ФНС
                            If ([System.IO.Path]::GetExtension($_.FullName) -eq ".xml"){
                                $file_answer_name = $_.Name

                                #Сообщения от ПФР
                                If (($file_answer_name.Substring(0,3) -eq 'SBR') -or ($file_answer_name.Substring(0,3) -eq 'SBP')){
                                    If ($waiting_pfr.ContainsKey($file_answer_name.Substring(3))){
                                        $message_id = $waiting_pfr[$file_answer_name.Substring(3)]
                                        Write-Verbose ("Обрабатываем файл {0} id {1}" -f $file_answer_name, $message_id)

                                        try {
                                            #Перекладываем принятый файл в хранилище
                                            [string] $temp_file_name = [System.Guid]::NewGuid()
                                            Copy-Item $_.FullName (Join-Path $Folder_IN $temp_file_name) -Force
                                            
                                            #Удаляем строки подписи в документе   
                                            [xml] $xml = Get-Content -Path $_.FullName -Raw | %{ [Regex]::Matches($_, "(?smi)(.+)Файл>") } | %{ $_.Value } 
                                            If ($xml.Файл){
                                                $message_kode_obr = $xml.Файл.Документ | Select КодОбр | %{ $_.КодОбр }
                                                $message_rez_obr = $xml.Файл.Документ | Select РезОбр | %{ $_.РезОбр }
                                                $message_date_obr = [datetime]::ParseExact(($xml.Файл.Документ | Select ДатаОбр | %{ $_.ДатаОбр }), 'dd.MM.yyyy', $null)
                                                $message_error_code = $xml.Файл.Документ.Ошибки | Select КодОшибки | %{ $_.КодОшибки }
                                                $message_error_desc = $xml.Файл.Документ.Ошибки | Select НаимОшибки | %{ $_.НаимОшибки }
                    
                                                If ( $message_error_code -ne "000"){
                                                    If ($waiting_emails.ContainsKey($file_answer_name.Substring(3))){
                                                        $hren, $hren2, $hren3, $jid = search_ldap $waiting_emails[$file_answer_name.Substring(3)]
                                                        If (-Not [string]::IsNullOrEmpty($jid)){
                                                            Pandion_Send ("311-П {0} {1} {2}" -f $message_id, $message_rez_obr, $message_error_desc) $jid
                                                        }Else{
                                                            Padion_Send ("311-П внутренняя ошибка не могу получить адрес для пользователя {0} отправки сообщения для файла: {1}" -f $waiting_emails[$file_answer_name.Substring(3)], $file_answer_name)
                                                        }
                                                    }Else{
                                                        Padion_Send ("311-П внутренняя ошибка отсутствует адрес для отправки сообщения для файла: {0}" -f $file_answer_name)
                                                    }
                                                }
                                                
                                                #Записываем файл в мифнс
                                                $SqlCmd.CommandText = "INSERT INTO [proto].[dbo].[R_mifns_pfr_fss] ([dt_z], [data_soobsh], [file_name], [file_content], [mifns_id], [err_code], [err_content], [transport_fromfns], [flie_cont_xml])
                                                VALUES ( CURRENT_TIMESTAMP, convert(datetime,'{0}'), '{1}', CONVERT(varbinary(max), '{2}'), '{3}', '{4}', '{5}', '{6}', '{7}' )" -f $message_date_obr.ToString("yyyy-MM-dd HH:mm:ss"), $file_answer_name, $temp_file_name, $message_id, $message_error_code, $message_error_desc, $file_name, $xml.OuterXml
                                                
                                                #Выполняем запрос
                                                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                                                if ($rowsAffected -gt 0){
                                                    Write-Verbose ("Записываем ПФР для файла {0} {1}" -f $file_answer_name, $message_id)
                                                    #return $false
                                                }Else{
                                                    return "Запись на SQL не обновлена в мифнс для файла {0}" -f $sended_filename
                                                }
                                            }Else{
                                                Pandion_Send ("311-П Нераспознанный формат файла отверта: {0}" -f $_.FullName)
                                            }
                                        }catch [Exception]{
                                            Write-Verbose $_.Exception.Message
                                            Pandion_Send ("311-П Пришел парсинга xml {3} файл {0} в архиве {1} транспорта {2}" -f $file_answer_name, $mifns_filename, $file_name, $_.Exception.Message)
                                        }
                                    }Else{
                                        Pandion_Send ("311-П Пришел неотправленный файл ПФР {0} в архиве {1} транспорта {2}" -f $file_answer_name, $mifns_filename, $file_name)
                                    }

                                #Сообщения от ФНС
                                }Else{
                                    If ($waiting_files.ContainsKey($file_answer_name.Substring(3))){
                                        $message_id = $waiting_files[$file_answer_name.Substring(3)]
                                        Write-Verbose ("Обрабатываем файл {0} id {1}" -f $file_answer_name, $message_id)
                                        
                                        #$input_file_raw = Get-Content -Path $_.FullName 
                                        #[xml] $xml = $input_file_raw[0..($input_file_raw.count - 2)]  

                                        try {
                                            #Удаляем строки подписи в документе   
                                            [xml] $xml = Get-Content -Path $_.FullName -Raw | %{ [Regex]::Matches($_, "(?smi)(.+)Файл>") } | %{ $_.Value } 
                                            If ($xml.Файл){
                                                $message_kode_obr = $xml.Файл.Документ | Select КодОбр | %{ $_.КодОбр }
                                                $message_rez_obr = $xml.Файл.Документ | Select РезОбр | %{ $_.РезОбр }
                                                $message_date_obr = [datetime]::ParseExact(($xml.Файл.Документ | Select ДатаОбр | %{ $_.ДатаОбр }), 'dd.MM.yyyy', $null)
                                                $message_errors = ""
                                                #Если в документе содержатся ошибки
                                                If ($xml.Файл.Документ.Ошибки -and ($message_kode_obr -ne 1)){
                                                    $message_error_code = $xml.Файл.Документ.Ошибки | Select КодОшибки | %{ $_.КодОшибки }
                                                    $message_error_desc = $xml.Файл.Документ.Ошибки | Select НаимОшибки | %{ $_.НаимОшибки } | %{ [Regex]::Replace($_, "'","") } 
                                                    $message_errors = ("[err_code] = '{0}', [err_content] = '{1}'," -f $message_error_code, $message_error_desc)
                                                    
                                                    If ($waiting_emails.ContainsKey($file_answer_name.Substring(3))){
                                                        #$hren, $hren2, $hren3, $jid = search_ldap $waiting_emails[$file_answer_name.Substring(3)]
                                                        If (-Not [string]::IsNullOrEmpty($jid)){
                                                            #Pandion_Send ("311-П {0} {1} {2}" -f $message_id, $message_rez_obr, $message_error_desc) $jid
                                                        }Else{
                                                            #Pandion_Send ("311-П внутренняя ошибка не могу получить адрес для пользователя {0} отправки сообщения для файла: {1}" -f $waiting_emails[$file_answer_name.Substring(3)], $file_answer_name)
                                                        }
                                                    }Else{
                                                        Padion_Send ("311-П внутренняя ошибка отсутствует адрес для отправки сообщения для файла: {0}" -f $file_answer_name)
                                                    }
                                                }
                                                
                                                #Записываем файл в мифнс
                                                $SqlCmd.CommandText = "UPDATE [proto].[dbo].[R_mifns]
                                                SET [data_primotveta] = CURRENT_TIMESTAMP, [data_obr] = convert(datetime,'{0}'), [file_otvet_name] = '{1}', [transport_fromfns] = '{2}', {3}  [file_otvet_xml] = '{4}', [kod_obr] = '{5}', [rez_obr] = '{6}'
                                                WHERE [uniq] = '{7}'" -f $message_date_obr.ToString("yyyy-MM-dd HH:mm:ss"), $file_answer_name, $file_name, $message_errors.subString(0, [System.Math]::Min(300, $message_errors.Length)), ($xml.OuterXml|%{ [Regex]::Replace($_, "'","")}), $message_kode_obr, $message_rez_obr, $message_id
                                                
                                                #Выполняем запрос
                                                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                                                if ($rowsAffected -gt 0){
                                                    Write-Verbose ("Записываем ФНС для файла {0} {1}" -f $file_answer_name, $message_id)
                                                    #return $false
                                                }Else{
                                                    return "Запись на SQL не обновлена в мифнс для файла {0}" -f $sended_filename
                                                }
                                            }Else{
                                                Pandion_Send ("311-П Нераспознанный формат файла отверта: {0}" -f $_.FullName)
                                            }
                                        }catch [Exception]{
                                            Write-Verbose $_.Exception.Message
                                            Pandion_Send ("311-П Пришел парсинга xml {3} файл {0} в архиве {1} транспорта {2}" -f $file_answer_name, $mifns_filename, $file_name, $_.Exception.Message)
                                        }
                                    }Else{
                                        Pandion_Send ("311-П Пришел неотправленный файл {0} в архиве {1} транспорта {2}" -f $file_answer_name, $mifns_filename, $file_name)
                                    }
                                }
                            }Else{
                                Pandion_Send ("Необычный файла 311-П архива: {0}" -f $_.FullName)
                            }
                        }
                        Remove-Item $status_extract_subfolder -Recurse -Force                    
                    }Else{
                        return ("Ошибка распаковки внутри 311-П файла: {0}" -f $_.FullName)
                    }
                }Else{
                    return ("Отсутствует архив внутри транспортного файла 311-П: {0}" -f $_.Fullname)
                }
            }
            Remove-Item $status_extract_folder -Recurse -Force

            #Перекладываем принятый файл в хранилище
            If (-Not (Test-Path (Join-Path $Folder_IN $file_name))){
                Move-Item $fileToExtract (Join-Path $Folder_IN $file_name) -Force
            }Else{
                Padion_Send ("311-П файл {0} уже существует в папке назначения: {1}" -f $file_name, $Folder_IN)
            }
        }Else{
            return "Ошибка распаковки 311-П файла: $file_name"
        }
    }Else{
        return "Неизвестный файл по 311-П: file_name"
    }
    $SqlConnection.Close()
    $SqlConnection.Dispose()

}

Function ReceiveFOR ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $file_name, $temp_file, $mail_from ){
    try {
        #Подключаемся к SQL серверу
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.Open()
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 600000

        #Генерируем ИД файла в архиве
        $temp_file_name = [guid]::NewGuid()

        #Распаковываем архив если это архив
        If ([System.IO.Path]::GetExtension($file_name) -eq ".arj"){
            $fileToExtract = (Join-Path $env:TEMP $file_name)
            Move-Item $temp_file $fileToExtract -Force
            $status_extract_folder = Extract_ARJ($fileToExtract)
            If ((Test-Path $status_extract_folder) -and (-Not [string]::IsNullOrEmpty($status_extract_folder))){
                #Это ФОР
                If (Test-Path (Join-Path $status_extract_folder 'Заголовок.xml')){
                
                    [xml] $xml = Get-Content -Path (Join-Path $status_extract_folder 'Заголовок.xml')
                    
                    $message_kode = $xml.Заголовок.Сведение | Select Код | %{ $_.Код }
                    $message_sender = $xml.Заголовок.Отправитель | Select Наименование | %{ $_.Наименование }
                    $message_id = $xml.Заголовок.ОЭС | Select Идентификатор | %{ $_.Идентификатор }
                    
                    #Записываем файл в мифнс
                    $SqlCmd.CommandText = "INSERT INTO [{6}].[dbo].[R_mgtu_in]
                 ([file_name],[file_path],[sender],[subject], [received_date], [file_id],[received_from]) 
                 VALUES('{0}','{1}','{2}', '{3}', CURRENT_TIMESTAMP, '{4}', '{5}')" -f $file_name, $temp_file_name, $message_sender, $message_kode, $message_id, $mail_from, $SQLDatabase
                
                    #Выполняем запрос
                    $rowsAffected = $SqlCmd.ExecuteNonQuery()
                    if ($rowsAffected -gt 0){
                        #Перекладываем принятый файл в хранилище
                        Move-Item $fileToExtract (Join-Path $Folder_IN $temp_file_name) -Force
                        Write-Verbose ("Записываем ФНС для файла {0} {1} {2}" -f $file_answer_name, $message_id, $temp_file_name)
                        #return $false
                    }Else{
                        return "Запись на SQL не обновлена в ФОР для файла {0}" -f $file_name
                    }
                #Это не ФО
                }Else{
                    If ([System.IO.Path]::GetExtension($file_name) -eq ".arj"){
                        $message_list = Get-ChildItem $status_extract_folder | Select -expand Name
                        $SqlCmd.CommandText = "INSERT INTO [{6}].[dbo].[R_mgtu_in]
                            ([file_name],[file_path],[sender],[subject], [received_date], [file_id],[received_from]) 
                            VALUES('{0}','{1}','{2}', '{3}', CURRENT_TIMESTAMP, '{4}', '{5}')" -f $file_name, $temp_file_name, $message_sender, "FSFM", $message_list, $mail_from, $SQLDatabase
                        
                        #Выполняем запрос
                        $rowsAffected = $SqlCmd.ExecuteNonQuery()
                        if ($rowsAffected -gt 0){
                            #Перекладываем принятый файл в хранилище
                            Move-Item $fileToExtract (Join-Path $Folder_IN $temp_file_name) -Force
                            Write-Verbose ("Записываем ФСФМ для файла {0} {1} {2}" -f $file_name, $message_list, $temp_file_name)
                        
                        }Else{
                            return "Запись на SQL не обновлена в ФОР для файла {0}" -f $file_name
                        }
                    }
                
                }
                Remove-Item $status_extract_folder -Recurse -Force
            }
        }
    
    }catch [Exception]{
        Write-Verbose $_.Exception.Message
        Pandion_Send ("For Sender Пришел парсинга ошибка подключени к SQL {0}" -f $_.Exception.Message)
    }finally{
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }
}

#Подключаемся к сетевому ресурсу
If (-not (Test-Path $Folder_IN)){
    $argc = " use $Net_Share $Net_Password /USER:$Net_User"
    $process = Start-Process -FilePath $Net_Path -windowstyle Hidden -ArgumentList $argc -PassThru -Wait
    $result = $process.ExitCode
    If ($result -ne 0){
        Pandion_Send("Не могу подключиться к $Folder_OUT как пользователь $Net_User")
        Exit
    }
}

$waiting_for_response = @{};

    #Получаем все непринятые файлы
    $test, $test2, $waiting_files, $waiting_emails, $waiting_pfr, $waiting_pfr_emails = SelectUnansweredFiles $SQLServer $SQLDatabase $SQLLogin $SQLPassword

    $waiting_pfr.Keys | % { "key = $_, value = " + $waiting_pfr.Item($_) }

    If ($waiting_for_response){
        #Номер почтового ящика - входящие
        $olFolderInbox = 6
    
        $outlook = new-object -com outlook.application;
        $ns = $outlook.GetNameSpace("MAPI");
        $inbox = $ns.GetDefaultFolder($olFolderInbox)

        #Вначале получаем файлы с внятной темой
        #Каждый файл отправляем в temp информацию заносим в базу
        $inbox.Items.Restrict("[ReceivedTime] > '{0}'" -f (Get-Date ((Get-Date).AddDays(-3)) -UFormat "%d/%m/%Y")) | 
        select -Expand Attachments | Sort-Object ReceivedTime -Descending | % {
            for ($i = $_.Count; $i ; $i--) {
                Write-Verbose $_.Parent.Subject
                Write-Verbose ("'"+ $($_.Item($i).FileName) +"'")
                Write-Verbose $_.Parent.SenderEmailAddress  
                $mail = $_.Parent.SenderEmailAddress  
                $subject = $_.Parent.Subject
                If (($mail -eq 'mifns2@ext-gate.svk.mskgtu.cbr.ru') -and ($subject -eq '311-П')){
                    #Сохраняем вложение
                    $temp_file = [System.IO.Path]::GetTempFileName()
                    $_.Item($i).SaveAsFile($temp_file)

                    #Устанавливаем статус принятого ответа
                    $return_status = Receive311P $SQLServer $SQLDatabase $SQLLogin $SQLPassword $($_.Item($i).FileName) $temp_file $waiting_files $waiting_emails $waiting_pfr $waiting_pfr_emails
                    if ($return_status){
                        Write-Verbose ("Произошла ошибка установки статуса: $return_status для 311-П")
                        Pandion_Send ("Произошла ошибка установки статуса: $return_status для 311-П")
                    }Else{
                        #Удаляем полученное письмо
                        $_.Parent.Unread = $True
                        $_.Parent.Delete()
                    }

                    If (Test-Path $temp_file){ Remove-Item $temp_file -Force }

                #Получаем файлы FOR
                }ElseIf (($mail -eq 'crypt@ext-gate.svk.mskgtu.cbr.ru') -and ($subject -eq 'FOR Sender')){
                    #Сохраняем вложение
                    $temp_file = [System.IO.Path]::GetTempFileName()
                    $_.Item($i).SaveAsFile($temp_file)

                    $return_status = ReceiveFOR $SQLServer $SQLDatabase $SQLLogin $SQLPassword $($_.Item($i).FileName) $temp_file $mail
                    if ($return_status){
                        Write-Verbose ("Произошла ошибка установки статуса: $return_status для ФОР")
                        Pandion_Send ("Произошла ошибка установки статуса: $return_status для ФОР")
                    }Else{
                        #Удаляем полученное письмо
                        $_.Parent.Unread = $True
                        $_.Parent.Delete()
                    }

                    If (Test-Path $temp_file){ Remove-Item $temp_file -Force }
                }ElseIf (($mail -eq 'fsfm349@ext-gate.svk.mskgtu.cbr.ru') -and ($subject -eq 'Передача архивных файлов электронных сообщений, адресованных уполномоченным банкам (филиалам уполномоченных банков)')){
                    #Сохраняем вложение
                    $temp_file = [System.IO.Path]::GetTempFileName()
                    $_.Item($i).SaveAsFile($temp_file)

                    $return_status = ReceiveFOR $SQLServer $SQLDatabase $SQLLogin $SQLPassword $($_.Item($i).FileName) $temp_file $mail
                    if ($return_status){
                        Write-Verbose ("Произошла ошибка установки статуса: $return_status для ФОР")
                        Pandion_Send ("Произошла ошибка установки статуса: $return_status для ФОР")
                    }Else{
                        #Удаляем полученное письмо
                        $_.Parent.Unread = $True
                        $_.Parent.Delete()
                    }

                    If (Test-Path $temp_file){ Remove-Item $temp_file -Force }
                }
            }
        }
    }