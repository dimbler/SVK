param (
    $Folder_IN = "", #Папка для входящих сообщений
    $Folder_OUT = "", #Папка для исходящих сообщений
    $SCSignEx = 'C:\Program Files\ГУ БАНКА РОССИИ по ЦФО\SCSignEx\SCSignEx.exe', #Путь к программе автоматизации распространяемой Банком России
    $subst = 'c:\Windows\System32\subst.exe', #Путь к консольной программе subst.exe
    $SQLServer = "", #Имя MS SQL сервера
    $SQLDatabase = "", #Имя базы данных на сервере
    $SQLLogin = "", #Имя пользователя MS SQL
    $SQLPassword = "", #Пароль пользователя MS SQL
    $arj_path = "\arj32.exe",
    $Net_User = "", #Имя пользователя для подключение к сетевому диску
    $Net_Password = "",  #Пароль пользователя для подключение к сетевому диску
    $Net_Path = "c:\Windows\System32\net.exe", #Путь к консольной программе net.exe
    $CryptCP = "cryptcp.exe",
    $Net_Share = "net_share", #Путь к корню сетевого ресурса
    $archTemplate = "'ARHBG_XXXX_'yyyMMdd_000", #Шаблон имени архива
    $BankName = "Коммерческий Банк", #Имя банка в шаблоне
    $BankNumber = "XXXX" #Номер банка в системе
)

$global:VerbosePreference = "Continue"
$program_path = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$sha1 = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider 

#Подключаемся к сетевому ресурсу
If (-not (Test-Path $Folder_OUT)){
    $argc = " use $Net_Share $Net_Password /USER:$Net_User"
    Write-Verbose $argc
    $process = Start-Process -FilePath $Net_Path -WindowStyle Hidden -ArgumentList $argc -PassThru -Wait
    $result = $process.ExitCode
    If ($result -ne 0){
        Pandion_Send("Не могу подключиться к $Folder_OUT как пользователь $Net_User")
        Exit
    }
}

#Отправка сообщения по Пандиону
function Pandion_Send ($Body) {
        Write-Verbose ("Сообщение {0}" -f $Body)
}


#Создание временной директории
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}

#Процедура упаковки в архив arj
function Compress_ARJ ($InputFileName, $MaxArchivSize = 0){
    $argc = "A -r -y"
    $arj_command = $program_path + $arj_path
    If (Test-Path $arj_command){
        If (Test-Path $InputFileName){
            $archive_folder = New-TemporaryDirectory
            #Если это каталог
            If ((Get-Item $InputFileName) -is [System.IO.DirectoryInfo]){
                $archive_name = Split-Path $InputFileName -Leaf
                $process_folder = $InputFileName
                $InputFileName = "*"
            #Если это файл
            }else{
                $archive_name = [io.path]::GetFileNameWithoutExtension($InputFileName)
                $process_folder = ([io.fileinfo]$InputFileName).DirectoryName
                $InputFileName = ([io.fileinfo]$InputFileName).Name
            }
            #Если есть ограничение по размеру для отправки
            If ($MaxArchivSize -gt 0){
                $argc += " -v" + [int]$MaxArchivSize*1000 + "K"
            }           

            $argc += " " + [string]$archive_folder + "\" + [string]$archive_name + ".arj " + '"' + $InputFileName + '"'
            $process = Start-Process -FilePath $arj_command -windowstyle Hidden -ArgumentList $argc -WorkingDirectory $process_folder -PassThru -Wait
            $result = $process.ExitCode
            If ($result -eq 0){
                return $archive_folder              
            }else{
                return ("Ошибка архивации файла $InputFileName")
            }
        }else{
            return ("Файл для архивации не найден: $InputFileName")
        }
    }else{
        return ("Отсутствует архиватор по пути: $arj_path")
    }
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

#Установить KA
Function KA_File ($InputFolderOrFile, $CHECK = 0, $KeyPath = $SignKeyPath) {
    if ((Test-Path $InputFolderOrFile) -and  ((Get-Item $KeyPath) -is [System.IO.DirectoryInfo])) { 
        if ((Get-Item $InputFolderOrFile).length -gt 0){
            $list_files = [System.IO.Path]::GetTempFileName()
            
            #Опеределяем файл это или папка
            If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                Get-ChildItem $current_folder -force | Foreach-Object {
                    $_.Name | Add-Content $list_files
                }
            }Else{
                $InputFolderOrFile | Add-Content $list_files
            }
            
            #Ожидаем разблокировки мютекса
            try{
                $Mutex = New-Object -TypeName System.Threading.Mutex -ArgumentList $false, "Global\KeyAction"
                $Mutex.WaitOne() | Out-Null

                #Удаляем старое ключи
                $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList "b: /d" -PassThru -Wait
                #Подсовываем ключи
                $arg = "b: " + (([System.IO.DirectoryInfo]$KeyPath).FullName)
                $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList $arg -PassThru -Wait
                $result = $process.ExitCode
                Write-Verbose $result
                If ($result -eq 0){
                    switch ($CHECK){
                        #Без параметров - установка КА
                        0 {$arg = "-s"}
                        #Проверка КА
                        1 {$arg = "-c"}
                        #Снятие КА
                        2 {$arg = "-r"}
                    }
                    #Подписываем КА
                    $log_file = [System.IO.Path]::GetTempFileName()

                    $arg += " -gb:\ -ib:\ -vb:\ -wb:\ -b0 -l$list_files -o$log_file"
                
                    If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                        $process = Start-Process -FilePath $SCSignEx -windowstyle Hidden -ArgumentList $arg -WorkingDirectory $InputFolderOrFile -PassThru -Wait
                    }Else{
                        $process = Start-Process -FilePath $SCSignEx -windowstyle Hidden -ArgumentList $arg -WorkingDirectory ([System.IO.Path]::GetDirectoryName($InputFolderOrFile)) -PassThru -Wait
                    }
                    $result = $process.ExitCode
                
                    $log_result = Get-Content -Path $log_file -Encoding default
                    Remove-Item $log_file -Force
                    Remove-Item $list_files -Force
                    Write-Verbose $result
                    If ($result -eq 0){
                        Write-Verbose "КА подписан успешно для каталога $InputFolderOrFile"
                        #Удаляем старые ключи
                        $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList "b: /d" -PassThru -Wait
                        return $true, $log_result         
                    } Else { 
                        return $false, ("Ошибка подписи файла/каталога {0} {1}" -f $InputFolderOrFile, $log_result)
                    }
                }Else{
                    return $false, "Ошибка подстановки ключей {0}" -f $KeyPath
                }
            }finally{
                #Разблокируем мютекс
                $Mutex.ReleaseMutex();
                $Mutex.Dispose();
            }            
        }Else{
            return $false, "Ошибка размера каталога/файл {0}" -f $InputFolderOrFile
        }
    }Else{
        return $false, "Ошибка! Каталог/файл {0} не существует" -f $InputFolderOrFile
    }
}

#Шифрование/расшифровка на абонента
Function Encrypt_File ($InputFolderOrFile, $Abonent = '0200', $DECRYPT = $false, $KeyPath = $ShifrKeyPath) {
    If ((Test-Path $InputFolderOrFile) -and  ((Get-Item $KeyPath) -is [System.IO.DirectoryInfo])) { 
        if ((Get-Item $InputFolderOrFile).length -gt 0){ 
            $list_files = [System.IO.Path]::GetTempFileName()
            
            #Опеределяем файл это или папка
            If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                Get-ChildItem $current_folder -force | Foreach-Object {
                    $_.Name | Add-Content $list_files
                }
            }Else{
                $InputFolderOrFile | Add-Content $list_files
            }

            #Ожидаем разблокировки мютекса
            try{
                $Mutex = New-Object -TypeName System.Threading.Mutex -ArgumentList $false, "Global\KeyAction"
                $Mutex.WaitOne() | Out-Null

                #Удаляем старые ключи
                $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList "b: /d" -PassThru -Wait
                #Подсовываем ключи
                $arg = "b: " + (([System.IO.DirectoryInfo]$KeyPath).FullName)
                $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList $arg -PassThru -Wait
                $result = $process.ExitCode
                Write-Verbose $result
                If ($result -eq 0){
                    #Без параметров - шифрование            
                    If ($DECRYPT -eq $false)
                    {
                        $arg = "-e"
                    #Расшифровка
                    }else{
                        $arg = "-d"
                    }
                    #Файл отчета
                    $log_file = [System.IO.Path]::GetTempFileName()

                    $arg += " -a$Abonent -gb:\ -ib:\ -vb:\ -wb:\ -b0 -l$list_files -o$log_file"
                    If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                        $process = Start-Process -FilePath $SCSignEx -windowstyle Hidden -ArgumentList $arg -WorkingDirectory $InputFolderOrFile -PassThru -Wait
                    }Else{
                        $process = Start-Process -FilePath $SCSignEx -windowstyle Hidden -ArgumentList $arg -WorkingDirectory ([System.IO.Path]::GetDirectoryName($InputFolderOrFile)) -PassThru -Wait
                    }
                    $result = $process.ExitCode
                    $log_result = Get-Content -Path $log_file -Encoding default
                
                    Remove-Item $log_file
                    Remove-Item $list_files -Force
                    Write-Verbose $result
                    If ($result -eq 0){
                        Write-Verbose "Файл/каталог зашифрован успешно на абонента $Abonent для файла $InputFolderOrFile"
                        #Удаляем старые ключи
                        $process = Start-Process -FilePath $subst -windowstyle Hidden -ArgumentList "b: /d" -PassThru -Wait
                        return $true, $log_result            
                    } Else { 
                        return $false, ("Ошибка шифрования файла/каталога {0} {1}" -f $InputFolderOrFile, $log_result)
                    }
                }Else{
                    return $false, "Ошибка подстановки ключей {0}" -f $KeyPath
                }
            }finally{
                #Разблокируем мютекс
                $Mutex.ReleaseMutex();
                $Mutex.Dispose();
            }
            
        }Else{
            return $false, "Ошибка размера файла {0}" -f $InputFile
        }
    }Else{
        return $false, "Ошибка! Файл {0} не существует" -f $InputFile
    }
}

#Подписание усиленной квалифицированной электронной подписью
function Sign_UKEP ($InputFolderOrFile, $CertName) {
    if (Test-Path $InputFolderOrFile) { 
        if ((Get-Item $InputFolderOrFile).length -gt 0){ 
            $out_path = New-TemporaryDirectory
            #Опеределяем файл это или папка
            If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                $arg += " -signf -dir $out_path -dn $CertName -der -norev -cert $InputFolderOrFile\*"
            }Else{
                $arg += " -signf -dir $out_path -dn $CertName -der -norev -cert $InputFolderOrFile"
            }
            Write-Verbose $arg
            $process = Start-Process -FilePath (Join-Path $program_path $CryptCP) -windowstyle Hidden -ArgumentList $arg -PassThru -Wait
            $result = $process.ExitCode
            If ($result -eq 0){
                Write-Verbose "УКЭП подписан успешно для файла/каталога $InputFolderOrFile"
                Get-ChildItem $out_path -force | Foreach-Object {
                    If ((Get-Item $InputFolderOrFile) -is [System.IO.DirectoryInfo]){
                        Write-Verbose (Join-Path $InputFolderOrFile ([System.IO.Path]::GetFileNameWithoutExtension($_.BaseName) + ".sign"))
                        Move-Item -Path $_.FullName -Destination (Join-Path $InputFolderOrFile ([System.IO.Path]::GetFileNameWithoutExtension($_.BaseName) + ".sign"))
                    }Else{                        
                        Move-Item -Path $_.FullName -Destination (Join-Path ([System.IO.Path]::GetDirectoryName($InputFolderOrFile)) ([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) + ".sign"))
                    }
                }
                Remove-Item $out_path -Recurse -Force
                return $false            
                } Else { 
                    return ("Ошибка подписи файла/каталога {0}" -f $InputFolderOrFile)
                }         
        }Else{
            return "Ошибка размера файла/каталога для подписи {0}" -f $InputFolderOrFile
        }
    }Else{
        return "Ошибка! Файл/каталог для подписи {0} не существует" -f $InputFolderOrFile
    }
}

#Получение списка файлов для отправки
function GetFilesFromSQL ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $SQLQuery) {  
    $SQLfiles = @()
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SQLQuery
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
 
        $SqlConnection.Close()
        $return_data = @()
        foreach ($row in $DataSet.Tables[0].Rows)
        { 
            $return_data += @{filename=$row[0].ToString().Trim(); fileguid=$row[1].ToString().Trim(); name_otch=$row[2].ToString().Trim()}
        }
        return $return_data
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Обновление статусов отправленных файлов
function UpdateSQLStatus ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $FileGUID, $HashFile, $FileMessage, $FileStatus){
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

        #Set Connection String
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.open()
    
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        #Записываем новый статус в поле
        $SqlCmd.CommandText = "UPDATE [{0}].[dbo].[R_mgtu_mail] SET message='{1}', hash_file='{2}', status='{4}', processed=GETDATE() where file_guid='{3}'" -f $SQLDatabase, $FileMessage.subString(0, [System.Math]::Min(254, $FileMessage.Length)), $HashFile, $FileGUID, $FileStatus
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 600000
        $rowsAffected = $SqlCmd.ExecuteNonQuery()
        if ($rowsAffected -gt 0){
            return $false
        }Else{
            return "Запись на SQL не обновлена для файла {0}" -f $FileGUID
        }    
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        return "Ошибка записи на SQL {0} для файла {1}" -f $_.Exception.Message, $FileGUID
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Получение данные справочника на SQL
function GetSPRFromSQL ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $OtchetnostName) {  
    $SQLfiles = @()
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "SELECT [mail],[kvit],[max_files],[max_size],[encr_id],[encr_path],[sign_path],[archiv_encr_id],[archiv_encr_path],[file_ukep],[archiv_ukep],[archive_name],[file_mask],[file_flow],[archiv_flow] FROM [ecdep].[dbo].[R_mgtu_spr] WHERE name_otch='"+$OtchetnostName+"'"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
 
        $SqlConnection.Close()
        $return_data = @()
        foreach ($row in $DataSet.Tables[0].Rows)
        { 
            $return_data = @{mail=$row[0].ToString().Trim(); kvit=$row[1].ToString().Trim(); max_files=$row[2].ToString().Trim(); max_size=$row[3].ToString().Trim(); encr_id=$row[4].ToString().Trim(); encr_path=$row[5].ToString().Trim(); sign_path=$row[6].ToString().Trim(); archiv_encr_id=$row[7].ToString().Trim(); archiv_encr_path=$row[8].ToString().Trim(); file_ukep=$row[9].ToString().Trim(); archiv_ukep=$row[10].ToString().Trim(); archiv_name=$row[11].ToString().Trim(); file_namenov=$row[12].ToString().Trim(); file_flow=$row[13].ToString().Trim() ; archiv_flow=$row[14].ToString().Trim()}
        }
        #$return_data | Sort-Object -Property Value
        return $return_data
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Получаем номер последнего отправляемого файла
function GetFileNameFromSQL ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $increment){
    $cur_date = Get-Date -format yyyyMMdd
    $return_data =  $cur_date + "_"
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "SELECT count([file_name]) FROM [ecdep].[dbo].[R_mgtu_mail] WHERE  [processed] > CAST(GETDATE() AS date)"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
 
        $SqlConnection.Close()
        foreach ($row in $DataSet.Tables[0].Rows)
        {
            $count_files = $row[0] + $increment
        }
        $return_data += $count_files.ToString("000")
        
        return $return_data
        }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Получаем номер последнего отправляемого файла
function GetArchivNameFromSQL ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $increment, $template = $archTemplate){
    If ([string]::IsNullOrEmpty($template)){ $template = $archTemplate }
    $iteract = (Get-Date).ToString($template)

    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "SELECT count([sended_filename]) FROM [ecdep].[dbo].[R_mgtu_queue] WHERE  [processed_time] > CAST(GETDATE() AS date)"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
 
        $SqlConnection.Close()
        foreach ($row in $DataSet.Tables[0].Rows)
        {
            $count_files = $row[0] + $increment
        }
        #If ($count_files -eq 0){
        #    $count_files = 1
        #}
        $return_data = $count_files.ToString($iteract)
        
        return $return_data
        }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Запись протокола обработки на сервер SQL
function WriteQueue ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $queued_files, $sended_filename, $name_otch ){
    try
    {
        #Создаем ИД задания на отправку
        $queue_id = [guid]::NewGuid();

        #Переносим полученные файлы в каталог для отправки
        Move-Item $sended_filename -Destination (Join-Path $Folder_OUT $queue_id) -Force
        
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

        #Set Connection String
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.open()
    
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        #Записываем новый статус в поле
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 600000
        
        #Записываем файл в очередь
        $SqlCmd.CommandText = "set nocount off; INSERT INTO [$SQLDatabase].[dbo].[R_mgtu_queue] ([processed_time],[queue_id],[sended_filename],[file_path],[status],[name_otch]) VALUES(CURRENT_TIMESTAMP, '{0}','{1}','{2}','{3}','{4}')" -f $queue_id,([io.fileinfo]$sended_filename).Name, $queue_id, 'processed', $name_otch
        
        #Выполняем запрос
        $rowsAffected = $SqlCmd.ExecuteNonQuery()
        if ($rowsAffected -gt 0){            

            #Обновляем данные по файлам
            Foreach ($queued_file in $queued_files){

                #Добавляем запись в таблицу сопоставление queue_id и file_id
                $SqlCmd.CommandText = "set nocount off; INSERT INTO [$SQLDatabase].[dbo].[R_mgtu_ref] ([task_guid],[file_guid]) VALUES('{0}','{1}')" -f $queue_id, $queued_file.fileguid
                #Выполняем запрос
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                if ($rowsAffected -eq 0){
                    return "Запись на SQL не обновлена в сопоставление для файла {0}" -f $queued_file.filename
                }
            }
        }Else{
            return "Запись на SQL не обновлена в очередь для файла {0}" -f $sended_filename
        }    
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        return "Ошибка записи на SQL {0} для файла {1}" -f $_.Exception.Message, $sended_filename
    }
    finally
    {
        $SqlConnection.Dispose()
    }
}

#Преобраование кодировок
function ConvertTo-Encoding ([string]$From, [string]$To){
      Begin{
            $encFrom = [System.Text.Encoding]::GetEncoding($from)
            $encTo = [System.Text.Encoding]::GetEncoding($to)
      }
      Process{
            $bytes = $encTo.GetBytes($_)
            $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)
            $encTo.GetString($bytes)
      }
}

#Печатаем полученные файлы
Function PrintReceivedFile($FileContains, $PrinterName){
    
    $printer = Get-WmiObject -Class Win32_Printer -ErrorAction Stop | Where {$_.Network} | Select-Object Name | Where { $_.Name -ieq $PrinterName }
    If ($printer){
        $PrinterName = $printer.Name
        Try { 
            $NetworkObj = New-Object -ComObject WScript.Network 
            $NetworkObj.AddWindowsPrinterConnection("$PrinterName") 
        }Catch {
            $argc = " use \\pft\ipc$ $Net_Password /USER:$Net_User"
            Write-Verbose $argc
            $process = Start-Process -FilePath $Net_Path -WindowStyle Hidden -ArgumentList $argc -PassThru -Wait
            $result = $process.ExitCode
            If ($result -ne 0){
                return("Не могу подключиться к $PrinterName как пользователь $Net_User")
            }
        }
        Try { 
            $FileContains | Out-Printer $PrinterName
        }Catch [Exception]{
            return "Не могу распечатать на принтер {0} {1}" -f $PrinterName, $_.Exception.Message
        }
    }Else{
        return "Принтер {0} не установлен" -f $PrinterName
    }
    
}

#Просматриваем статусы полученных ответов
Function SetStatusResponseKA1($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword){
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.Open()

        #Сначала получаем вариант ошибок с SQL
        $error_code = 0
        $input_errors = @{}
        $SqlCmde = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmde.CommandText = "SELECT [err_message],[err_flag] FROM [ecdep].[dbo].[R_mgtu_err]"
        $SqlCmde.Connection = $SqlConnection
        $SqlAdaptere = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdaptere.SelectCommand = $SqlCmde
        $DataSete = New-Object System.Data.DataSet
        $SqlAdaptere.Fill($DataSete)
        foreach ($rowe in $DataSete.Tables[0].Rows)
        {
            $input_errors.Add($rowe['err_message'].ToString(), $rowe['err_flag'].ToString())
        }
        
        #Теперь получаем необработанные статусы
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "SELECT q.[queue_id],q.[received_file_k1],q.[received_path_k1], spr.[prn] FROM [ecdep].[dbo].[R_mgtu_queue] q INNER JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch WHERE ([status] = 'received_k1' OR [status] = 'received_k2') AND ([status_k1] IS NULL)"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd  
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
       
        foreach ($row in $DataSet.Tables[0].Rows)
        {
            $result_ka1 = ""
            
            #Получаем файл с сервера
            $queue_id = $row['queue_id'].ToString()
            $status_filename = $row['received_file_k1'].ToString()
            $status_filepath = (Join-Path $Folder_IN $row['received_path_k1'].ToString())
            $status_temp_file = [System.IO.Path]::GetTempFileName()
            $printer = $row['prn'].ToString()
            Copy-Item -Path $status_filepath -Destination $status_temp_file -Force

            #Распаковываем архив если это архив внутри архива должен быть xml в кодировке utf8
            If ([System.IO.Path]::GetExtension($status_filename) -eq ".arj"){
                $status_extract_folder = Extract_ARJ($status_temp_file)
                Get-ChildItem $status_extract_folder -Force | Foreach-Object {
                    #Снимаем КА
                    $result_ka, $result_message = KA_File $_.FullName 2
                    If (($result_ka -eq $false) -or (!$result_ka)){
                        Pandion_Send("Ощибка подписи файла статуса $queue_id $result_message")
                        Continue
                    }
                    #Если требуется распечатать отчет
                    If ($printer){
                        [byte[]]$byte = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $_.FullName
                        if ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf ){
                            $file_contains = Get-Content -Path $_.FullName -Encoding UTF8
                        }else{
                            $file_contains = Get-Content -Path $_.FullName
                        }
                        $result_print = PrintReceivedFile $file_contains $printer
                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                    }

                    #Проверяем кодировку полученного документа
                    #[byte[]]$byte = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $_.FullName
                    #if ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf ){
                    If (Get-Content -Path $_.FullName | %{ [Regex]::Matches($_, "utf-8") } ){
                        [xml]$xml = Get-Content -Path $_.FullName -Encoding UTF8
                    }else{
                         [xml]$xml = Get-Content -Path $_.FullName
                    }                    
                    $result = ($xml.UV.REZ_ARH).ToString()
                    If ($result.ToString()){
                        $result_ka1 += $result.ToString()
                    }Else{
                        Pandion_Send ("Ошибка обработки статуса ошибка содержания для файла {0}" -f $queue_id)
                        $result_ka1 = Get-Content -Path $_.FullName -Encoding UTF8 | Out-String
                    }
                }
                Remove-Item -Path $status_extract_folder -Recurse -Force

            #Получена квитанция от ГУЦБ
            }ElseIF ([System.IO.Path]::GetExtension($status_filename) -like ".kvt*"){
                $result_ka, $result_message = KA_File $status_temp_file 2
                If (($result_ka -eq $false) -or (!$result_ka)){
                    Pandion_Send("В файле ответа КА не найдено: $result_message $queue_id")

                    $file_data_result = (Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | Out-String)

                    #Если требуется распечатать отчет
                    If ($printer){
                        $result_print = PrintReceivedFile $file_data_result $printer
                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                    }
                }Else{
                    #Если требуется распечатать отчет
                    If ($printer){
                        $file_contains = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251
                        $result_print = PrintReceivedFile $file_contains $printer
                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                    }
                    $file_data_result = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | %{ [Regex]::Matches($_, "Результат.*") } | %{ $_.Value }

                    #Необычная отчетность
                    If ([string]::IsNullOrEmpty($file_data_result)){
                        $file_data_result = (Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | Out-String)
                        If (-Not ( %{ [Regex]::Matches($file_data_result, "принят") } )){
                            Pandion_Send("Необычный ответ на файл {0} {1}" -f $status_filename, $queue_id)
                        }   
                    }
                }

                #Проверяем данные на ошибки при разборе
                if (-Not [string]::IsNullOrEmpty($file_data_result)){
                    $result_ka1 = $file_data_result 
                }Else{
                    Pandion_Send ("Ошибка разбора файла {0} {1}" -f $status_filename, $queue_id)
                    Continue
                }
            #Получена квитанция от transfer
            }ElseIF ([System.IO.Path]::GetExtension($status_filename) -like ".txt*"){
                $result_ka, $result_message = KA_File $status_temp_file 2
                If (($result_ka -eq $false) -or (!$result_ka)){
                    Pandion_Send("В файле ответа КА не найдено: $result_message $queue_id")
                }                    

                $file_data_result = (Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | %{ [Regex]::Replace($_, "'","") }  | Out-String)
                #Если требуется распечатать отчет
                If ($printer){
                    $result_print = PrintReceivedFile $file_data_result $printer
                    If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                }
                
                #Проверяем данные на ошибки при разборе
                If (-Not [string]::IsNullOrEmpty($file_data_result)){
                    $result_ka1 = $file_data_result
                }Else{
                    Pandion_Send ("Ошибка разбора файла {0} {1}" -f $status_filename, $queue_id)
                    Continue
                }
            #Если нераспознанное расширение для файла
            }Else{
                $result_ka1 = Get-Content -Path $status_temp_file -Encoding UTF8
                Pandion_Send ("Ошибка обработки статуса для файла {0}" -f $queue_id)
            }
            Remove-Item $status_temp_file -Force
            If (-Not [string]::IsNullOrEmpty($result_ka1)){
                #Проверяем на наличие ошибок
                $error_code = 0
                foreach ($input_error in $input_errors.GetEnumerator()){
                    If (%{ [Regex]::Matches($result_ka1, $($input_error.Name))}){
                        $error_code = $($input_error.Value)
                    }
                }

                #Записываем новый статус в поле   
                $SqlCmd.CommandText = ("UPDATE [{0}].[dbo].[R_mgtu_queue] SET status_k1='{1}', err_flag_k1='{3}' WHERE queue_id='{2}'" -f $SQLDatabase, $result_ka1.subString(0, [System.Math]::Min(254, $result_ka1.Length)), $queue_id, $error_code)
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                if ($rowsAffected -gt 0){
                    #return $false
                }Else{
                    Pandion_Send ("Статус на SQL не обновлен для файла {0}" -f $queue_id)
                    Continue
                }    
            }Else{
                Pandion_Send ("Статус с нулевым содержанием для файла {0}" -f $queue_id)
                Continue
            }
        }
    }catch [Exception]{
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения статусов с SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SQLAdapter.Dispose()
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }

}

#Получаем статусы ответов Агентов обмена
Function SetStatusResponseKA2($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword){
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.Open()

        #Сначала получаем вариант ошибок с SQL
        $error_code = 0
        $input_errors = @{}
        $SqlCmde = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmde.CommandText = "SELECT [err_message],[err_flag] FROM [ecdep].[dbo].[R_mgtu_err]"
        $SqlCmde.Connection = $SqlConnection
        $SqlAdaptere = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdaptere.SelectCommand = $SqlCmde
        $DataSete = New-Object System.Data.DataSet
        $SqlAdaptere.Fill($DataSete)
        foreach ($rowe in $DataSete.Tables[0].Rows)
        {
            $input_errors.Add($rowe['err_message'].ToString(), $rowe['err_flag'].ToString())
        }

        #Теперь получаем необработанные статусы
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "SELECT q.[queue_id],q.[received_file_k2],q.[received_path_k2], spr.[prn] FROM [ecdep].[dbo].[R_mgtu_queue] q INNER JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch WHERE ([status] = 'received_k2') AND ([status_k2] IS NULL)"
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
       
        foreach ($row in $DataSet.Tables[0].Rows)
        {
            $result_ka2 = ""
            
            #Получаем файл с сервера
            $queue_id = $row['queue_id'].ToString()
            $status_filename = $row['received_file_k2'].ToString()
            $status_filepath = (Join-Path $Folder_IN $row['received_path_k2'].ToString())
            $status_temp_file = [System.IO.Path]::GetTempFileName()
            $printer = $row['prn'].ToString()
            Copy-Item -Path $status_filepath -Destination $status_temp_file -Force

            #Распаковываем архив если это архив
            If ([System.IO.Path]::GetExtension($status_filename) -eq ".arj"){
                $status_extract_folder = Extract_ARJ($status_temp_file)
                If ((Test-Path $status_extract_folder) -and (-Not [string]::IsNullOrEmpty($status_extract_folder))){
                    Get-ChildItem $status_extract_folder -Force | Foreach-Object {

                        #Если внутри архива архив, тогда распаковываем и его
                        If ([System.IO.Path]::GetExtension($_.FullName) -eq ".arj"){

                            #Проверяем, возможно следует архив разподписать
                            $result_ka, $result_message = KA_File $_.FullName 2
                            
                            #Расшифровываем архив, если требуется
                            $result_ka, $result_message = Encrypt_File $_.FullName "0200" $True
                            If (($result_ka -eq $false) -or (!$result_ka)){
                                Pandion_Send("Ощибка расшифровки файла статуса $queue_id $result_message")
                                Continue
                            }

                            #Извлекаем файлы их архива
                            $file_extract_folder = Extract_ARJ($_.FullName)
                            If ((Test-Path $file_extract_folder) -and (-Not [string]::IsNullOrEmpty($file_extract_folder))){
                                Get-ChildItem $file_extract_folder -Force | Foreach-Object {
                                    
                                    If (Get-Content -Path $_.FullName | %{ [Regex]::Matches($_, "(?i)utf-8") } ){
                                         [xml] $xml = Get-Content -Path $_.FullName -Encoding UTF8 | Out-String
                                         If ($xml.KVIT.ERRORS_ES.ERR_REC.NAM_ERR){
                                            $result_ka2 = ($xml.KVIT.ERRORS_ES.ERR_REC.NAM_ERR).ToString()
                                         }ElseIf ($xml.Уведомление){
                                            $result_ka2 = $xml.OuterXml
                                         }
                                    }Else{
                                        Write-Verbose $_.FullName
                                    }
                                    #Если требуется распечатать отчет
                                    If ($printer){
                                        $file_contains = Get-Content -Path $_.FullName -Encoding UTF8
                                        $result_print = PrintReceivedFile $file_contains $printer
                                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                                    }                                
                                }
                                Remove-Item $file_extract_folder -Recurse -Force
                            }

                        #Если внутри архива нет архива, а лежит файл в формате xml                    
                        }Else{

                            #Снимаем КА
                            Write-Verbose $_.FullName
                            $result_ka, $result_message = KA_File $_.FullName 2
                            If (($result_ka -eq $false) -or (!$result_ka)){
                                Pandion_Send("Ощибка разподписи файла статуса $queue_id $result_message")
                                Continue
                            }
                            #Если требуется распечатать отчет
                            If ($printer){
                                $file_contains = Get-Content -Path $_.FullName -Encoding UTF8
                                $result_print = PrintReceivedFile $file_contains $printer
                                If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                            }

                            #UTF8 стандартная кодировка для принимаемых сообщений из ФНС и Таможни
                            If (Get-Content -Path $_.FullName | %{ [Regex]::Matches($_, "(?i)utf-8") } ){
                                [xml]$xml = Get-Content -Path $_.FullName -Encoding UTF8
                                $result = ($xml.UV.REZ_ARH).ToString()
                            }Else{
                                $result = Get-Content -Path $_.FullName
                            }
                            If ($result.ToString()){
                                $result_ka2 += $result.ToString()
                            }Else{
                                Pandion_Send ("Ошибка обработки статуса ошибка содержания для файла {0}" -f $queue_id)
                                $result_ka2 = Get-Content -Path $_.FullName -Encoding UTF8 | Out-String
                            }
                        }
                    }
                    Remove-Item -Path $status_extract_folder -Recurse -Force
                }
            #Если пришла квитанция в кодировке dos
            }ElseIF ([System.IO.Path]::GetExtension($status_filename) -like ".kvt*" -or [System.IO.Path]::GetExtension($status_filename) -eq ".txt"){

                #Первым делом снимаем КА
                $result_ka, $result_message = KA_File $status_temp_file 2

                #Если отсутствует КА
                If (($result_ka -eq $false) -or (!$result_ka)){
                    Pandion_Send("В файле ответа КА не найдено: $result_message $queue_id")
                    $file_data_result = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251
            
                    #Если требуется распечатать отчет
                    If ($printer){
                        $result_print = PrintReceivedFile $file_data_result $printer
                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                    }
                
                #Если КА присутствует
                }Else{
                    #Если требуется распечатать отчет
                    If ($printer){
                        $file_contains = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251
                        $result_print = PrintReceivedFile $file_contains $printer
                        If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                    }

                    #Ответ от ГУ ЦБ
                    $file_data_result = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | %{ [Regex]::Matches($_, "Результат.*") } | %{ $_.Value }
                    If (-Not $file_data_result){
                        $file_data_result = (Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251 | Out-String)
                        If ([System.IO.Path]::GetExtension($status_filename) -like ".kvt*") {
                            Pandion_Send("Необычный ответ на файл {0} {1}" -f $status_filename, $queue_id)
                        }
                    }
                }

                #Проверяем на существования данных
                if (-Not [string]::IsNullOrEmpty($file_data_result)){
                    $result_ka2 = $file_data_result
                }Else{
                    Pandion_Send ("Ошибка разбора файла {0} {1}" -f $status_filename, $queue_id)
                    Continue
                }

            #Ответ в формате XML
            }ElseIf([System.IO.Path]::GetExtension($status_filename) -eq ".xml"){
                [xml]$xmlresp = Get-Content -Path $status_temp_file
                $result = ($xmlresp.ТКвит.Пояснение).ToString()
                If ($result){
                    $result_ka2 += $result
                }Else{
                    $result_ka2 = Get-Content -Path $status_temp_file
                }

                #Если требуется распечатать отчет
                If ($printer){
                    $file_contains = Get-Content -Path $status_temp_file
                    $result_print = PrintReceivedFile $file_contains $printer
                    If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                }

            #Если получена отчетность номерная КФМ
            }ElseIf (%{ [Regex]::Matches([System.IO.Path]::GetExtension($status_filename), "\d{3}") } ){
                #Снимаем КА
                $result_ka, $result_message = KA_File $status_temp_file 2
                If (($result_ka -eq $false) -or (!$result_ka)){
                    Pandion_Send("Ощибка разподписи файла статуса $queue_id $result_message")
                    Сontinue
                }

                $file_data_result = Get-Content -Path $status_temp_file | ConvertTo-Encoding cp866 windows-1251
                #Если требуется распечатать отчет
                If ($printer){
                    $result_print = PrintReceivedFile $file_data_result $printer
                    If ($result_print){ Pandion_Send ("Ошибка печати файла {0} на принтер {1} {2} {3}" -f $status_filename, $printer, $queue_id, $result_print) }
                }

                #Определяем количество принятых записей
                if ($file_data_result){
                    $result_ka2 = (%{ [Regex]::Matches($file_data_result, "Всего записей:(.+)Принятых:\s?\d+") } | %{ $_.Value})
                #Если прислалали неожиданных результат
                }Else{
                    $result_ka2 = $file_data_result
                    Pandion_Send ("Ошибка разбора файла {0} {1}" -f $status_filename, $queue_id)
                }

            #Ошибка обработки статуса для файла
            }Else{
                $result_ka2 = Get-Content -Path $status_temp_file
                Pandion_Send ("Ошибка обработки статуса для файла {0}" -f $queue_id)
                Continue
            }
            Remove-Item $status_temp_file -Force

            #Полученные данные записываем в базу
            If (-Not [string]::IsNullOrEmpty($result_ka2)){
                #Проверяем на наличие ошибок
                $error_code = 0
                foreach ($input_error in $input_errors.GetEnumerator()){
                    If (%{ [Regex]::Matches($result_ka2, $($input_error.Name))}){
                        $error_code = $($input_error.Value)
                    }
                }

                #Записываем новый статус в поле   
                $SqlCmd.CommandText = ("UPDATE [{0}].[dbo].[R_mgtu_queue] SET status_k2='{1}', err_flag_k2='{3}' WHERE queue_id='{2}'" -f $SQLDatabase, $result_ka2.subString(0, [System.Math]::Min(254, $result_ka2.Length)), $queue_id, $error_code)
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                if ($rowsAffected -gt 0){
                    #return $false
                }Else{
                    Pandion_Send ("Статус на SQL не обновлен для файла {0}" -f $queue_id)
                    Continue
                }    
            }Else{
                Pandion_Send ("Статус с нулевым содержанием для файла {0}" -f $queue_id)
                Continue
            }
        }
    }catch [Exception]{
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Ошибка чтения статусов с SQL {0}" -f $_.Exception.Message)
    }
    finally
    {
        $SQLAdapter.Dispose()
        $SQLAdaptere.Dispose()
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }

}

#Создание и добавление файла информации в архив
Function AddInfToArchive($cur_archive_folder, $inf_file){
    If ((Test-Path $cur_archive_folder) -and $cur_archive_folder){
        #Если заголовок мы где-то уже взяли
        If ((Test-Path $inf_file) -and $inf_file){
            If ([bool]((Get-Content $inf_file) -as [xml])){
                Move-Item -Path $inf_file -Destination (Join-Path $cur_archive_folder "Заголовок.xml")
            }Else{
                return "Плохое содержание xml файла $inf_file"
            }
        #Если заголовок требуется создать
        }ElseIf ($inf_file){
            $kvit_temp_file = [System.IO.Path]::GetTempFileName()
            $xmlWriter = New-Object System.XMl.XmlTextWriter($kvit_temp_file,[System.Text.Encoding]::GetEncoding("windows-1251"))
            $xmlWriter.WriteStartDocument()
            $xmlWriter.WriteStartElement('Заголовок')
            $xmlWriter.WriteStartElement('Сведения')
            $XmlWriter.WriteAttributeString('Код', $inf_file)
            $xmlWriter.WriteStartElement('Отправитель')
            $xmlWriter.WriteAttributeString('Наименование', $BankName)
            $xmlWriter.WriteAttributeString('РегНомер', $BankNumber)
            $xmlWriter.WriteEndElement()
            $xmlWriter.WriteStartElement('Данные')
            $XmlWriter.WriteAttributeString('ОтчетнаяДата', (Get-Date).ToString("yyyy'-'MM'-'dd"))
            Get-Childitem $cur_archive_folder | Select-Object  | Where-Object {$_.Extension -notlike '*.doc' -and $_.Extension -notlike '*.pdf' -and $_.Extension -notlike '*.tif' -and $_.Extension -notlike '*.tiff' -and $_.Extension -notlike '*.xls' -and $_.Extension -notlike '*.xlsx' -and $_.Extension -notlike '*.docx' -and $_.Extension -notlike '*.jpg' -and $_.Extension -notlike '*.jpeg' -and $_.Extension -notlike '*.png' -and $_.Extension -notlike '*.xml' -and $_.Extension -notlike '*.arj'}| Foreach-Object {
                $xmlWriter.WriteStartElement('Отчет')  
                $XmlWriter.WriteAttributeString('Имя', $_.Name)  
                $xmlWriter.WriteEndElement()
            }
            Get-ChildItem $cur_archive_folder | Select-Object  | Where-Object {$_.Extension -like '*.doc' -or $_.Extension -like '*.pdf' -or $_.Extension -like '*.tif' -or $_.Extension -like '*.tiff' -or $_.Extension -like '*.xls' -or $_.Extension -like '*.xlsx' -or $_.Extension -like '*.docx' -or $_.Extension -like '*.jpg' -or $_.Extension -like '*.jpeg' -or $_.Extension -like '*.png' -or $_.Extension -like '*.xml' -or $_.Extension -like '*.arj'}| Foreach-Object {
                $xmlWriter.WriteStartElement('Файл')  
                $XmlWriter.WriteAttributeString('Имя', $_.Name)  
                $xmlWriter.WriteEndElement()
            }
            $xmlWriter.WriteEndElement()
            $xmlWriter.WriteEndElement()
            $xmlWriter.WriteEndDocument()
            $xmlWriter.Flush()
            $xmlWriter.Close()

            Move-Item $kvit_temp_file (Join-Path $cur_archive_folder "Заголовок.xml")
        }Else{
            return "Невозможно постоить заголовок с пустыми данными $inf_file"
        }
    }Else{
        return "Отсутствует папка для добавления информации $cur_archive_folder"
    }
}

#Для ФНС запись файлов в таблицу
Function MIFNS_File ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $input_file, $input_file_guid){

    If (Test-Path $input_file){
        try
        {
            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = “Server=tcp:$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
            $SqlConnection.Open()
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.Connection = $SqlConnection
            $SqlCmd.CommandTimeout = 600000

            [xml] $xml = Get-Content -Path $input_file
            If ($xml.Файл){
                $message_file_name = [System.IO.Path]::GetFileName($input_file)
                $message_sender_name = $xml.Файл | Select ФамОтпр | %{ $_.ФамОтпр }
                $message_date = [datetime]::ParseExact(($xml.Файл.Документ | Select ДатаСооб | %{ $_.ДатаСооб }), 'dd.MM.yyyy', $null)
                $message_sch = $xml.Файл.Документ.СвСчет | Select НомСч | %{ $_.НомСч }

                #Записываем файл в мифнс
                $SqlCmd.CommandText = "INSERT INTO [proto].[dbo].[R_mifns]
                 ([data_soobsh], [FIOOI], [Sch], [data_otpr], [file_name], [key_doc]) 
                 VALUES(convert(datetime,'{0}'),'{1}','{2}', CURRENT_TIMESTAMP,'{3}', '{4}')" -f $message_date.ToString("yyyy-MM-dd HH:mm:ss"), $message_sender_name, $message_sch, $message_file_name, $input_file_guid
                
              
                #Выполняем запрос
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                if ($rowsAffected -gt 0){
                    Write-Verbose "Записываем ФНС"
                    return $false

                }Else{
                    return "Запись на SQL не обновлена в мифнс для файла {0}" -f $sended_filename
                }  
            }Else{
                return "Ошибка разбора xml файла $input_file"
            }

        }catch [Exception]{
            Write-Verbose $_.Exception.Message
            Pandion_Send ("Ошибка чтения статусов с SQL {0}" -f $_.Exception.Message)
        }finally{
            $SqlConnection.Close()
            $SqlConnection.Dispose()
        }

    }Else{
        return "Ошибка чтения файла $input_file"
    }
}

#Разделение массива на несколько массивов по размеру
Function Split-array {

  param($inArray,[int]$parts,[int]$size)
  
  if ($parts) {
    $PartSize = [Math]::Ceiling($inArray.count / $parts)
  } 
  if ($size) {
    $PartSize = $size
    $parts = [Math]::Ceiling($inArray.count / $size)
  }

  $outArray = @()
  for ($i=1; $i -le $parts; $i++) {
    $start = (($i-1)*$PartSize)
    $end = (($i)*$PartSize) - 1
    if ($end -ge $inArray.count) {$end = $inArray.count}
    $outArray+=,@($inArray[$start..$end])
  }
  return ,$outArray

}

#Получаем статусы ответов ГУ ЦБ
SetStatusResponseKA1 $SQLServer $SQLDatabase $SQLLogin $SQLPassword

#Получаем статусы ответов Агентов обмена
SetStatusResponseKA2 $SQLServer $SQLDatabase $SQLLogin $SQLPassword

#Получение данных с сервера и обработка
$sqlfiles = @();
$tasks_lists = @();
#$test,$sqlfiles = GetFilesFromSQL $SQLServer $SQLDatabase $SQLLogin $SQLPassword "select file_name, file_guid, name_otch from [ecdep].[dbo].[R_mgtu_mail] WHERE processed IS NOT NULL AND operator_name IS NOT NULL AND [status] IS NULL AND [dt_z] > CAST(dateadd(day,datediff(day,5,GETDATE()),0) AS date) order by dt_z asc"
$test,$sqlfiles = GetFilesFromSQL $SQLServer $SQLDatabase $SQLLogin $SQLPassword "select file_name, file_guid, name_otch from [ecdep].[dbo].[R_mgtu_mail] WHERE processed IS NULL AND operator_name IS NOT NULL order by dt_z asc"
#Получаем количество разнообразных форм отчетности
$otchetnosti = $sqlfiles |  ForEach-Object {$_.name_otch} | Select-Object -Unique

#Выбираем отчетность из списка отчетностей на отправку
Foreach ($otchetnost in $otchetnosti){

    #Получаем данные из справочника SQL
    $test ,$spr_data = GetSPRFromSQL $SQLServer $SQLDatabase $SQLLogin $SQLPassword $otchetnost
    If (-Not $spr_data){
        Pandion_Send("Ошибка получения справочника отчетности по форме $otchetnost")
        Exit
    }

    #Максимальное количество файлов в архиве, если не указано иное
    $max_files = 99
    #Получаем максимальное количество файлов в одном архиве
    If ($spr_data["max_files"]){
        $max_files = $spr_data["max_files"]
    }

    #Заголовочный файл
    $inf_file = ""

    #Получаем последовательность действий с файлом
    [string[]]$file_flow = $spr_data.file_flow.Split('; ',[System.StringSplitOptions]::RemoveEmptyEntries)
    If ($file_flow.Contains('UNPACK')){

        #Распаковываем существующие архивы, если их надо распаковать
        $archived_input_files = $sqlfiles |  ForEach-Object {$_} | Where-Object {$_.filename -like "*.arj"}
        $remain_sqlfiles = $sqlfiles |  ForEach-Object {$_} | Where-Object {-not ($_.filename -like "*.arj")}
        if ($remain_sqlfiles){
            $sqlfiles = $remain_sqlfiles
        }else{
            $sqlfiles = @();
        }
        Foreach ($archived_input_file in $archived_input_files){
            $temp_archived_file = [System.IO.Path]::GetTempFileName()
            Copy-Item -Path (Join-Path $Folder_OUT $archived_input_file.fileguid) -Destination $temp_archived_file
            $extract_folder = Extract_ARJ($temp_archived_file)
            Remove-Item $temp_archived_file -Force
            Get-ChildItem $extract_folder -Force | Foreach-Object {
                If ($_.Name -ne "Заголовок.xml"){
                    $sqlfiles += @{filename=$_.Name; fileguid=$archived_input_file.fileguid; name_otch=$archived_input_file.name_otch}
                    Copy-Item $_.FullName (Join-Path $current_folder $_.Name) -Force
                }Else{
                    $inf_file = [System.IO.Path]::GetTempFileName()
                    Move-Item -Path $_.FullName -Destination $inf_file -Force
                }
            }
            Remove-Item $extract_folder -Recurse -Force
        }
    }

    $cotch_files = $sqlfiles |  ForEach-Object {$_} | Where-Object {$_.name_otch -eq $otchetnost}
    

    #Записываем на SQL принято в обработку
    $cotch_files |  ForEach-Object {
        $set_status = UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid "" "Принято в обработку" "S"
        If ($set_status -ne $false){
            Pandion_Send("Ошибка обновления статуса файла {0} GUID: {1} для отчетности {2}" -f $_.filename, $_.fileguid, $otchetnost)
            Continue
        }
    }
    Write-Verbose $cotch_files.Length
    #Добавляем в задание на обработку
    If ($cotch_files.Length -gt $max_files){
        #Если максимальное количество файлов превышает возможности - разбиваем на несколько заданий
        $split_files = Split-array -inArray $cotch_files -size $max_files
        $split_files |  ForEach-Object {
            $tasks_lists += @{ files=$_; inf_file=$inf_file; spr_data=$spr_data; name_otch=$otchetnost }
        }
    }Else{
        $tasks_lists += @{ files=$cotch_files; inf_file=$inf_file; spr_data=$spr_data; name_otch=$otchetnost }
    }
}

#Начинаем работать с полученным заданием
$tasks_lists |  ForEach-Object { 
    #Создаем директорию для временных файлов
    $current_folder = New-TemporaryDirectory
    
    #Сообщения в процессе обработки
    $processed_message = ""
    
    #Текущее задание
    $current_task = $_.files

    #Текущая отчетность
    $current_otch = $_.name_otch

    #Текущий справочник
    $spr_data = $_.spr_data

    #Очередь файлов в работу
    $queued_files = @()

    #Номер файла при переименовании
    $increment_j = 0
        
    Write-Verbose "Выполняю работы по отчетности $current_otch"

    #Получаем последовательность действий с файлом
    [string[]]$file_flow = $spr_data.file_flow.Split('; ',[System.StringSplitOptions]::RemoveEmptyEntries)

    #Готовим файлы к дальшейшим операциям
    $current_task | ForEach-Object { 
        
        #Текущий файл в задании
        $current_file = $_        

        #Собираем файлы для работы
        If ((Test-Path (Join-Path $Folder_OUT $current_file.fileguid)) -and $current_file){

            #Получаем имя для файла
            If (-not $spr_data.file_namenov){
                $inputFileName = (Join-Path $current_folder $current_file.filename)
            }else{
                #Получаем имя архивного файла с сервера
                $increment_j ++
                $test, $otch_file_name = (GetFileNameFromSQL $SQLServer $SQLDatabase $SQLLogin $SQLPassword $increment_j)
                If (-Not $otch_file_name){
                    Pandion_Send ("Ошибка при получении имени файла с SQL $otch_file_name")
                    Continue
                }Else{
                    $otch_file_name = $spr_data.file_namenov + $otch_file_name
                    $inputFileName = (Join-Path $current_folder ($otch_file_name+[System.IO.Path]::GetExtension($current_file.filename)))
                }
            }            
            
            #Если файл не из распакованного архива - требуется его копирование в локальную папку
            If (-Not $file_flow.Contains('UNPACK')){
                Copy-Item -Path (Join-Path $Folder_OUT $current_file.fileguid) -Destination $inputFileName
            }
            
            #Если требуется обновить данный файл в таблице [proto].[dbo].[R_mifns]
            If (($current_otch -eq 'mifns_f') -or ($current_otch -eq 'mifns_j')){
                $result_mifns = MIFNS_File $SQLServer $SQLDatabase $SQLLogin $SQLPassword $inputFileName $current_file.fileguid
                If ($result_mifns) { Pandion_Send("Ощибка записи файла в базу МИФНС $inputFileName $result_mifns") }
            }

            #Вначале считаем контрольную сумму файла
            $begin_hash = [System.BitConverter]::ToString( $sha1.ComputeHash([System.IO.File]::ReadAllBytes($inputFileName)))
            
            #Записываем всю информацию о файле
            $queued_files += @{fileguid=$current_file.fileguid; filename=$inputFileName; filehash=$begin_hash; filemessage="Добавлен"}

        }Else{
            Pandion_Send("Не найден файл {0} GUID: {1} для отчетности {2}" -f $current_file.filename, $current_file.fileguid, $otchetnost)
            Continue
        }
    }

    #Разбираем последовательность действий с файлами
    Foreach ($file_action in $file_flow){
        Write-Verbose "Следующее действие с файлом: $file_action"

        switch ($file_action){
            UNPACK {
            }
            #Подписываем файл/каталог
            KA {
                If (Test-Path $spr_data.sign_path){
                    $result_ka, $result_message = KA_File $current_folder 0 $spr_data.sign_path
                    If (($result_ka -eq $false) -or (!$result_ka)){
                        Pandion_Send("Ощибка подписи файла/каталога {0} $result_message" -f $current_otch)
                        $current_task | ForEach-Object { UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid "Произошла ошибка при подписи файла" $result_message "Err"}
                        Exit
                    }
                    #Записываем все результаты отдельно для каждого файла
                    $queued_files | ForEach-Object { $_.filemessage += $result_message | Select-String ([io.path]::GetFileName($_.filename)) }
                    Write-Verbose " Подписано КА успешно "
                }else{
                    Pandion_Send("Ощибка справочника настройки подписи файла $inputFileName форма {0}" -f $current_otch)
                    Exit
                }
            }
            #Шифрование файла/каталога
            ENCR {
                If (($spr_data.encr_id -ne '') -and ($spr_data.encr_id)){
                    $result_enc, $result_message = Encrypt_File $current_folder $spr_data.encr_id $false $spr_data.encr_path
                    If ($result_enc -eq $false){
                        Pandion_Send("Ощибка шифрования файла/каталога {0} $result_message" -f $current_otch)
                        $current_task | ForEach-Object { UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid "Произошла ошибка при шифровании файла" $result_message "Err"}
                        Exit
                    }
                    #Записываем все результаты отдельно для каждого файла
                    $queued_files | ForEach-Object { $_.filemessage += $result_message | Select-String ([io.path]::GetFileName($_.filename)) }
                    Write-Verbose " Зашифровано успешно "
                }Else{
                    Pandion_Send("Ощибка справочника настройки шифрования файла $inputFileName форма {0}" -f $current_otch)
                    Exit
                }
            }
            #Создаем усиленную квалифицированную электронную подпись
            UKEP {
                If ($spr_data.file_ukep){
                    $result_ukep = Sign_UKEP $current_folder $spr_data.file_ukep
                    If ($result_ukep){
                        Pandion_Send("Ощибка усиленной квалифицированной электронной подписи файла/каталога $current_folder $result_ukep")
                        $current_task | ForEach-Object { UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid "Произошла ошибка при ЭЦП файла/каталога" $result_ukep "Err"}
                        Exit
                    }
                    #Записываем все результаты отдельно для каждого файла
                    $queued_files | ForEach-Object { $_.filemessage += " УКЭП подписано успешно " }
                    Write-Verbose " УКЭП подписано успешно "
                }else{
                    Pandion_Send("Ощибка справочника настройки усиленной квалифицированной электронной подписи файла $inputFileName форма {0}" -f $current_otch)
                    Exit
                }
            }
            #Если выбор не распознан
            default {
                Pandion_Send("Ощибка распознования последовательности действий справочника для файла $inputFileName форма {0}" -f $current_otch)
                Exit
            }
                    
        }
    }

    #Упаковываем все сделанное в архив, если требуется
    $cdirectoryInfo = (Get-ChildItem $current_folder | Measure-Object).Count
    If ($cdirectoryInfo -ne 0){
        #Переменная сборки информации в процессе работы
        $processed_message = ""

        #Добавление в архив файла информации
        $add_inf = ""
        If ($inf_file){
            $add_inf = AddInfToArchive $current_folder $inf_file
        }ElseIf($spr_data.kvit){
            $add_inf = AddInfToArchive $current_folder $spr_data.kvit
        }
        If ($add_inf){
            Pandion_Send("Ошибка добавления информационного файла $add_inf для отчетности {0}" -f $current_otch)
            Continue
        }
        
        #Архивирование файла
        If ($spr_data.archiv_flow){
            $return_folder = Compress_ARJ $current_folder $spr_data.max_size
            If (-not (Test-Path $return_folder)){
                Pandion_Send("Ощибка создания архива для каталога $current_folder $return_folder")
                $queued_files | ForEach-Object { UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid $_.filehash "Ошибка создания архива для каталога $current_folder $return_folder" "Err"}
                Exit
            }Else{
                #Удаляем папку после архивирования
                Remove-Item $current_folder -Recurse -Force 
                $processed_message += " Каталог заархивирован "

                $increment_i = 0;

                $current_folder = $return_folder
                $adirectoryInfo = (Get-ChildItem $current_folder | Measure-Object).Count
                If ($adirectoryInfo -ne 0){
                    Get-ChildItem $current_folder | Foreach-Object {

                        #Получаем имя для архива
                        Write-Verbose ([System.IO.Path]::GetExtension($spr_data.archiv_name))
                        If ([System.IO.Path]::GetExtension($spr_data.archiv_name) -eq '.arj'){
                            $archive_name = (Join-Path ([io.fileinfo]$_.FullName).DirectoryName $spr_data.archiv_name)
                        }else{
                            #Получаем имя архивного файла с сервера
                            $increment_i++ 
                            $test, $cbr_file_name = (GetArchivNameFromSQL $SQLServer $SQLDatabase $SQLLogin $SQLPassword $increment_i $spr_data.archiv_name)
                            $archive_name = (Join-Path ([io.fileinfo]$_.FullName).DirectoryName ($cbr_file_name+[System.IO.Path]::GetExtension($_.FullName)))
                        }

                        #Переименовываем полученный архив
                        Move-Item $_.FullName $archive_name -Force

                        #Разбираем последовательность действий с архивом
                        [string[]]$archiv_flow = $spr_data.archiv_flow.Split('; ',[System.StringSplitOptions]::RemoveEmptyEntries)
                        Foreach ($archiv_action in $archiv_flow){
                            Write-Verbose "Следующее действие с архивом: $archiv_action"

                            switch ($archiv_action){
                                #Создаем информационное сообщение для архива
                                KVIT {
                                    #Добавление файла информации
                                    $add_inf = ""
                                    If($spr_data.kvit){
                                        $add_inf = AddInfToArchive $current_folder $spr_data.kvit
                                    }else{
                                        Pandion_Send("Не найдены данные для информацинного файла для архива $archive_name")
                                    }
                                    If ($add_inf){
                                        Pandion_Send("Ошибка добавления информационного файла $add_inf для архива $archive_name")
                                        Continue
                                    }
                                }
                                #Создаем усиленную квалифицированную электронную подпись
                                UKEP {
                                    If ($spr_data.archiv_ukep){
                                        $result_ukep = Sign_UKEP $archive_name $spr_data.archiv_ukep
                                        If (-not (Test-Path $result_ukep)){
                                            Pandion_Send("Ощибка усиленной квалифицированной электронной подписи архива $archive_name $result_ukep")
                                            Exit
                                        }else{
                                            Move-Item $result_ukep ($archive_name +".sign") -Force
                                            Write-Verbose $result_ukep
                                            $processed_message += " Архив подписан УКЭП "
                                        }
                                    }else{
                                        Pandion_Send("Ощибка справочника настройки усиленной квалифицированной электронной подписи архива $archive_name форма {0}" -f $current_otch)
                                        Exit
                                    }
                                }
                                #Подпись архива
                                KA {
                                    If (Test-Path $spr_data.sign_path){
                                        $result_ka, $result_message = KA_File $archive_name 0 $spr_data.sign_path
                                        If ($result_ka -ne $True){
                                            Pandion_Send("Ощибка подписи файла архива $archive_name $result_message")
                                            Exit
                                        }else{
                                            $processed_message += " Архив подписан КА " + [string]($result_message -join '-')
                                        }
                                    }else{
                                        Pandion_Send("Ощибка справочника настройки подписи архива $archive_name форма {0}" -f $current_otch)
                                        Exit
                                    }
                                }
                                #Шифрование архива
                                ENCR {
                                    If ((Test-Path ($spr_data.archiv_encr_path)) -and ($spr_data.archiv_encr_id)){
                                        $result_enc, $result_message = Encrypt_File $archive_name $spr_data.archiv_encr_id $false $spr_data.archiv_encr_path
                                        If ($result_enc -ne $True){
                                            Pandion_Send("Ощибка проверки подписи файла $archive_name $result_message")
                                            Exit
                                        }else{
                                            $processed_message += " Архив зашифрован " + [string]($result_message -join '-')
                                        }
                                    }else{
                                        Pandion_Send("Ощибка справочника настройки подписи архива $archive_name форма {0}" -f $current_otch)
                                        Exit
                                    }                        
                                }
                                #Упаковка подписанного архива в транспортный файл
                                TRANSPORT {
                                    $file_to_archive = $archive_name
                                    #Если кроме архива присутствует еще файл подписи - переносим их в специально созданный каталог
                                    If ($spr_data.archiv_ukep){
                                        $temp_folder = New-TemporaryDirectory
                                        Move-Item $archive_name -Destination $temp_folder -Force
                                        Move-Item ($archive_name +".sign") -Destination $temp_folder -Force
                                        $file_to_archive = $temp_folder
                                    }
                                    $return_folder = Compress_ARJ $file_to_archive
                                    If (-not (Test-Path $return_folder)){
                                        Pandion_Send("Ошибка создания архива для файла $archive_name $return_folder")
                                        Exit
                                    }
                                    If (Test-Path $file_to_archive){ Remove-Item $file_to_archive -Recurse -Force }
                                    Move-Item -Path "$return_folder\*" -Destination $current_folder -Force
                                    Remove-Item $return_folder
                                }
                                #Если выбор не распознан
                                default {
                                    Pandion_Send("Ощибка распознования последовательности действий справочника для архива $archive_name форма {0}" -f $current_otch)
                                    Exit
                                }
                            }
                        }
                    }
                }Else{
                    Pandion_Send("Ошибка создания архива для папки $current_folder")
                    $queued_files | ForEach-Object { UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid $_.filehash "Ошибка создания архива для папки $current_folder $return_folder" "Err"}
                    Exit
                }
            }
        }
    
        #Обновляем данные по файлам
        $queued_files |  ForEach-Object {
            $set_status = UpdateSQLStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $_.fileguid $_.filehash $_.filemessage "F"
            If ($set_status -ne $false){
                Pandion_Send("Ошибка обновления статуса файла {0} GUID: {1}" -f $_.filename, $_.fileguid)
                Continue
            }
        }

        Get-ChildItem $current_folder -force | Foreach-Object {
            Write-Verbose ("Подготовлен к отправке файл {0}" -f $_.Name)

            #Записываем все в очередь на отправку
            $queue_send_status = WriteQueue $SQLServer $SQLDatabase $SQLLogin $SQLPassword $queued_files $_.FullName $current_otch
            If ($queue_send_status){
                Pandion_Send ("Ошибка записи файла {0} в очередь {1}" -f $_.Name, $queue_send_status )
                Write-Verbose $queue_send_status
            }
        }
        #Удаляем из временной папки
        Remove-Item $current_folder -Force -Recurse


    }Else{
        Pandion_Send("Папка {0} с файлами для отчетности {1} пуста!" -f $current_folder, $otchetnost)
    }
}
 