param (
    $smtp_server = "192.168.19.3", #Адрес сервера электронной почты в сети МГТУ
    $smtp_port = 25,
    $mail_from = "svkIDn2@svk.mskgtu.cbr.ru", #EMAIL для исходящих писем
    $slogin = "svkIDn2", #Логин электронной почты
    $spass = "", #Пароль
    $Folder_IN = "", #Папка для входящих сообщений
    $Folder_OUT = "out", #Папка для исходящих сообщений
    $Net_User = "", #Имя пользователя для подключение к сетевому диску
    $Net_Password = "", #Пароль пользователя для подключение к сетевому диску
    $SQLServer = "", #Имя MS SQL сервера
    $SQLDatabase = "", #Имя базы данных на сервере
    $SQLLogin = "", #Имя пользователя MS SQL
    $SQLPassword = "", #Пароль пользователя MS SQL
    $Net_Path = "c:\Windows\System32\net.exe", #Путь к консольной программе net.exe
    $Net_Share = "net_share", #Путь к корню сетевого ресурса
    $SVK_Login = "svkXXXXn2", #Логин СВК
    $SVK_Pass = '', #Пароль СВК
    $VPN_Name = '"Подключение к СВК" VPN_Name VPN_Pass' #Строка VPN подключения
)

$global:VerbosePreference = "Continue"
$program_path = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

#Cоздаем временный каталог
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}

#Отправка сообщения по Пандиону
function Pandion_Send ($Body) {
        
        Write-Verbose ("Сообщение {0}" -f $Body)
}

# Function to loop through array of commands and write them to console.
Function Send-Command ($Command, $Wait){
	$Writer.WriteLine($Command) 
	Write-Verbose "Sent: $Command"
	try 
	{
		$Writer.Flush()
	} 
		
	catch 
	{
		Write-Error "Writer Flush Error"
		Write-Debug "DEBUG Flush: $Error "
	}
    Start-Sleep -Milliseconds $WaitTime
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

Function TestPOP ( $Username, $Password, [System.Net.IPAddress]$IP, [int]$Port = 110 ){

    If ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($Port) ){
        return "Не могу проверить адрес или порт не указан"
    }Else{
        $test = New-Object System.Net.Sockets.TcpClient;
        Try
        {
            # Suppress error messages
            $ErrorActionPreference = 'SilentlyContinue'
            $test.Connect($IP, $Port);
            if ($test.Connected) {
                Write-Verbose "Connection successful";
                $NetStream = $test.GetStream();
                $Reader = new-object -TypeName System.IO.StreamReader($NetStream);
                $Writer = new-object -TypeName System.IO.StreamWriter($NetStream);
                $Buffer = $Reader.ReadLine();
                $Writer.WriteLine("USER $Username");
                $Writer.Flush(); $Buffer = $Reader.ReadLine();
                $Writer.WriteLine("PASS $Password");
                $Writer.Flush();
                If ($Reader.ReadLine() -match "OK")
                {
                    Write-Verbose "Auth OK"
                }Else{
                    Write-Verbose "Authentication Error"
                    return "Authentication Error"
                }
                $Reader.Dispose();
                $Writer.Dispose();
                $NetStream.Dispose();
                $test.Close()
            }else {
                Write-Vebose "Port is closed or filtered"
            }
        }
        Catch
        {
            return "Connection failed"
        }
        Finally
        {
            $test.Dispose();
        }
    }
}

#Оргазация доступа к МГТУ для отправки почты
Function Telnet_Login($Login, $Password, $RHost, $TPort = 23, $WaitTime = 2000, $SuccessFlag = "Authentication successful"){

    # Create a new connetion with host
    Write-Output "Processing $RHost ..."
    Try { $Socket = New-Object System.Net.Sockets.TcpClient($RHost, $TPort)}
    Catch { Write-Error "Unable to connect to host: $($RHost):$TPort"; Exit } 

    # Check to make sure the connection is active
    If ($Socket){   
        $Stream = $Socket.GetStream()
        $Writer = New-Object System.IO.StreamWriter($Stream)
        $Buffer = New-Object System.Byte[] 2048 
        $Encoding = New-Object System.Text.AsciiEncoding
        Start-Sleep -Milliseconds 4000 # Four seconds to allow the connection to establish
        
        $Commands = @($Login, $Password)	    

	    # Send each command to Send-Command function
        ForEach ($Command in $Commands)
        {   
	        If (!$Socket){ Write-Output "Connection Failed"; break}
		        Send-Command $Command $WaitTime
        }
		
        # Read console output and save it to $ConsoleOutput variable
        While($Stream.DataAvailable) 
        {   
            $Read = $Stream.Read($Buffer, 0, 2048) 
            $ConsoleOutput += ($Encoding.GetString($Buffer, 0, $Read))
        }
		
        # Use -debug switch to view the console output.
		Write-Debug $ConsoleOutput
		
		# If SuccessFlag variable provided, check console output for string.
		If ($SuccessFlag)
		{
			Write-Verbose "Checking for SuccessFlag: ${SuccessFlag} "
			If ($ConsoleOutput -like "*${SuccessFlag}*")
			{
				Write-Verbose "Process Successful!"
                return $true
			}
			Else
			{
				Write-Verbose "Process Failed!"
                return $false
			}
		}
			
		# Close socket connections
		$Writer.Close()
		$Stream.Close()
	}	
	Else 
	{ 
		$ConsoleOutput = "Unable to connect to host: ${RHost}:${Port}" 
        return $false
    }
}

#Подключаемся к сетевому ресурсу
If (-not (Test-Path $Folder_OUT)){
    $argc = " use $Net_Share $Net_Password /USER:$Net_User"
    $process = Start-Process -FilePath $Net_Path -windowstyle Hidden -ArgumentList $argc -PassThru -Wait
    $result = $process.ExitCode
    If ($result -ne 0){
        Pandion_Send("Не могу подключиться к $Folder_OUT как пользователь $Net_User")
        Exit
    }
}

Function SelectUnansweredFiles ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword){
    $waiting_queues = @{}
    $waiting_files = @{}
    $waiting_pandions = @{}

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
        #$SqlCmd.CommandText = "SELECT q.sended_filename, q.[queue_id], spr.mail FROM [ecdep].[dbo].[R_mgtu_queue] q LEFT JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch WHERE (status = 'sended' OR status='received_k1') and [processed_time] > CAST(dateadd(day,datediff(day,14,GETDATE()),0) AS date)"
        $SqlCmd.CommandText = "SELECT mail.[file_name], q.[queue_id], spr.mail, spr.pandion FROM [ecdep].[dbo].[R_mgtu_queue] q INNER JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch INNER JOIN [ecdep].[dbo].[R_mgtu_ref] ref ON q.[queue_id] = ref.[task_guid] INNER JOIN [ecdep].[dbo].[R_mgtu_mail] mail ON ref.[file_guid] = mail.[file_guid] WHERE (q.status = 'sended' OR q.status='received_k1') and [processed_time] > CAST(dateadd(day,datediff(day,14,GETDATE()),0) AS date)"
        #$SqlCmd.CommandText = "SELECT q.[queue_id], spr.mail FROM [ecdep].[dbo].[R_mgtu_queue] q INNER JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch WHERE (status = 'sended' OR status='received_k1') and [processed_time] > CAST(dateadd(day,datediff(day,14,GETDATE()),0) AS date)"
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        If ($DataSet.Tables[0].Rows.Count -gt 0){
            Foreach ($row in $DataSet.Tables[0].Rows){ 
                If (-Not $waiting_queues.ContainsKey($row['queue_id'].ToString())) {$waiting_queues.Add($row['queue_id'].ToString(), $row['mail'].ToString()) }
                If (-Not $waiting_files.ContainsKey($row['file_name'].ToString())) { $waiting_files.Add($row['file_name'].ToString(), $row['queue_id'].ToString()) }
                If (-Not $waiting_pandions.ContainsKey($row['queue_id'].ToString())) { $waiting_pandions.Add($row['queue_id'].ToString(), $row['pandion'].ToString()) }
            }
        }
        return $waiting_queues, $waiting_files, $waiting_pandions
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ( "Ошибка чтения на SQL {0} для поиска неотвеченных файлов файла" -f $_.Exception.Message )
    }
    finally
    {
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }
        
}

function WriteQueueStatus ($SQLServer, $SQLDatabase, $SQLLogin, $SQLPassword, $queued_file, $response_file_name, $response_file_path ){
    
    [string] $temp_file_name = [System.Guid]::NewGuid()

    If (Test-Path $response_file_path){
        Move-Item -Path $response_file_path -Destination (Join-Path $Folder_IN $temp_file_name) -Force
    }Else{
        return "Файл $response_file_name не найден по пути $response_file_path"
    }
    #$response_file_path = (Split-Path $response_file_path -Leaf)
    $response_file_path = $temp_file_name

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
        $SqlCmd.CommandText = "SELECT [queue_id],[received_file_k1],[received_file_k2] FROM [ecdep].[dbo].[R_mgtu_queue] WHERE [queue_id]='{0}'" -f $queued_file
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        
        If ($DataSet.Tables[0].Rows.Count -gt 0){
            Foreach ($row in $DataSet.Tables[0].Rows)
            {
                #Переменная указывает в какое поле записывать
                $set_to_answer = 'k1'
                $received_status = 'k1'

                #Если Файла К1 не было
                If([string]::IsNullOrEmpty($row['received_file_k1']) -and [string]::IsNullOrEmpty($row['received_file_k2'])){
                    $set_to_answer = 'k1'
                    $received_status = 'k1'
                #Если файл К1 уже приходил
                }Else{
                    $set_to_answer = 'k2'
                    $received_status = 'k2'
                }

                #Проверяем чтобы ответы от ГУЦБ складывались правильно
                If ([System.IO.Path]::GetExtension($response_file_name) -like ".kvt*"){
                    $file_data_result = Get-Content -Path (Join-Path $Folder_IN $response_file_path) | ConvertTo-Encoding cp866 windows-1251 | %{ [Regex]::Matches($_, "ИЭС(\d{1})") } | %{ $_.Value } 
                    If (-Not [string]::IsNullOrEmpty($file_data_result)){
                        Write-Verbose $file_data_result 
                        If ($file_data_result -eq "ИЭС1"){
                            $set_to_answer = 'k1'
                        }ElseIf ($file_data_result -eq "ИЭС2"){
                            $set_to_answer = 'k2'
                        }Else{
                            return "Ошибка разбора файла {0}" -f $response_file_name
                        }
                    }
                }

                $SqlCmd.CommandText = "UPDATE [{4}].[dbo].[R_mgtu_queue] SET status='received_{5}', received_time_{0}=GETDATE(), received_path_{0}='{1}', received_file_{0}='{2}' where [queue_id]='{3}'" -f $set_to_answer, $response_file_path, $response_file_name, $queued_file,  $SQLDatabase, $received_status
                Write-Verbose ("UPDATE [{4}].[dbo].[R_mgtu_queue] SET status='received_{5}', received_time_{0}=GETDATE(), received_path_{0}='{1}', received_file_{0}='{2}' where [queue_id]='{3}'" -f $set_to_answer, $response_file_path, $response_file_name, $queued_file,  $SQLDatabase, $received_status)
                #Выполняем запрос
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                if ($rowsAffected -gt 0){            
                    return $false
                }Else{
                    return "Запись на SQL не обновлена в очередь для файла {0}" -f $response_file_name
                }    
            }
        }Else{            
            $SqlCmd.CommandText = "set nocount off; INSERT INTO [$SQLDatabase].[dbo].[R_mgtu_queue] ([received_time_k1],[received_path_k1],[received_file_k1],[status]) VALUES(CURRENT_TIMESTAMP, '{0}','{1}','{2}')" -f ([io.fileinfo]$response_file_path).Basename, $response_file_name, 'received_k1'
            $rowsAffected = $SqlCmd.ExecuteNonQuery()
            If ($rowsAffected -gt 0){            
                return "Пришел ответ по неотправленному файлу: " -f $response_file_name
            }Else{
                return "Запись на SQL не обновлена в очередь для файла {0}" -f $response_file_name
            }            
        }
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        return "Ошибка записи на SQL {0} для файла {1}" -f $_.Exception.Message, $response_file_name
    }
    finally
    {
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }
}

#Проверка входящей почты и разбор набежавших сообщений
Function Get_Sended_Status {
    
    $waiting_for_response = @{}
    $waiting_files = @{}
    $waiting_pandions = @{}

    #Получаем все непринятые файлы
    $test, $waiting_for_response, $waiting_files, $waiting_pandions = SelectUnansweredFiles $SQLServer $SQLDatabase $SQLLogin $SQLPassword

    $waiting_files.Keys | % { "key = $_, value = " + $waiting_files.Item($_) }

    If ($waiting_for_response){
        #Номер почтового ящика - входящие
        $olFolderInbox = 6
    
        $outlook = new-object -com outlook.application;
        $ns = $outlook.GetNameSpace("MAPI");

        # Для отправки сообщения требуется авторизация
        If (TestPop $SVK_Login $SVK_Pass '192.168.19.4'){
            If (-Not (Telnet_Login $SVK_Login $SVK_Pass '192.168.19.20') -eq $true){
                Exit
            }Else{
                Write-Verbose "Connected via Telnet"
            }
        }Else{
            Write-Verbose "Exist Connection"
        }
        $ns.SendAndReceive(1)
        $inbox = $ns.GetDefaultFolder($olFolderInbox)

        #Вначале получаем файлы с внятной темой
        #Каждый файл отправляем в temp информацию заносим в базу
        $inbox.Items.Restrict("[ReceivedTime] > '{0}'" -f (Get-Date ((Get-Date).AddDays(-5)) -UFormat "%d/%m/%Y")) | 
        select -Expand Attachments | Sort-Object ReceivedTime -Descending | % {
            for ($i = $_.Count; $i ; $i--) {
                Write-Verbose $_.Parent.Subject
                        Write-Verbose ("'"+ $($_.Item($i).FileName) +"'")
                        Write-Verbose $_.Parent.SenderEmailAddress  
                #Проверяем чтобы это был ответ на посланое нами письмо
                IF (-Not [string]::IsNullOrEmpty($_.Parent.Subject)){
                    IF ($waiting_for_response.ContainsKey($_.Parent.Subject)){
                        $sended_queue_id = $_.Parent.Subject
                    }Else{
                        $sended_queue_id = (%{ [Regex]::Matches($_.Parent.Subject, "(?i)^Re:\s?(.+)$") } | %{ $_.Captures.Groups[1].Value })
                    }
                }
                If (-Not [string]::IsNullOrEmpty($sended_queue_id)){

                    #Проверяем чтобы ответ был на содержащееся в очереди письмо
                    if ($waiting_for_response.ContainsKey($sended_queue_id) -and (-Not [string]::IsNullOrEmpty($sended_queue_id))){
                        
                        Write-Verbose $_.Parent.Subject
                        Write-Verbose $($_.Item($i).FileName)
                        Write-Verbose $_.Parent.SenderEmailAddress 
                        
                        #Сохраняем вложение
                        $temp_file = [System.IO.Path]::GetTempFileName()
                        $_.Item($i).SaveAsFile($temp_file)

                        #Удаляем полученное письмо
                        $_.Parent.Unread = $True
                        $_.Parent.Delete()
            
                        #Устанавливаем статус принятого ответа
                        $test, $return_status = WriteQueueStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $sended_queue_id $($_.Item($i).FileName) $temp_file
                        if ($return_status){
                            Pandion_Send ("Произошла ошибка установки статуса принятого ответа: $return_status")
                        }

                        if (-Not [string]::IsNullOrEmpty($waiting_pandions[$sended_queue_id])){ Pandion_Send (("Принята отчетность {0}" -f $_.Parent.Subject), $waiting_pandions[$sended_queue_id]) }
                    }                                   
                }
            }
        }

        #!!!Это спорная логика!!!
        #Затем получаем файлы не содержащие ответ Re:
        $inbox.Items.Restrict("[ReceivedTime] > '{0}'" -f (Get-Date ((Get-Date).AddDays(-5)) -UFormat "%d/%m/%Y")) | 
        select -Expand Attachments | Sort-Object ReceivedTime -Descending | % {
            for ($i = $_.Count; $i ; $i--) {
                Write-Verbose $_.Parent.SenderEmailAddress
                Write-Verbose $_.Parent.Subject
                #Проверяем чтобы это не был ответ на посланое нами письмо
                If (-Not (%{ [Regex]::Matches($_.Parent.Subject, "^Re:\s?(.+)$") })){

                    $sended_address = $_.Parent.SenderEmailAddress

                    #Сохраняем вложение
                    $temp_file = [System.IO.Path]::GetTempFileName()
                    $_.Item($i).SaveAsFile($temp_file)

                    #Для файлов от КФМ проверяем имя файла в содержимом
                    $fileid_in_file = (Get-Content -Path $temp_file | ConvertTo-Encoding cp866 windows-1251 | %{ [Regex]::Matches($_, "Файл:\s?(.+)?,\s?Размер") } | %{ $_.Captures.Groups[1].Value })
                    If (-Not [string]::IsNullOrEmpty($fileid_in_file)){
                        If ($waiting_files.ContainsKey($fileid_in_file)){
                            Write-Verbose $_.Parent.Subject
                            Write-Verbose $waiting_files[$fileid_in_file]

                            #Удаляем полученное письмо
                            $_.Parent.Unread = $True
                            $_.Parent.Delete()   

                            #Устанавливаем статус принятого ответа
                            $test, $return_status = WriteQueueStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $waiting_files[$fileid_in_file] $($_.Item($i).FileName) $temp_file
                            if ($return_status){
                                Pandion_Send ("Произошла ошибка установки статуса принятого ответа: $return_status")
                            } 
                            if (-Not [string]::IsNullOrEmpty($waiting_pandions[($waiting_files[$fileid_in_file])])){ Pandion_Send (("Принята отчетность {0}" -f $_.Parent.Subject), $waiting_pandions[($waiting_files[$fileid_in_file])]) }
                        }
                    #Колхозная часть для приема ответов от fts545@ext-gate.svk.mskgtu.cbr.ru
                    }ElseIf($waiting_for_response.ContainsValue($sended_address) -and (-Not [string]::IsNullOrEmpty($sended_address)) -and ($sended_address -eq 'fts545@ext-gate.svk.mskgtu.cbr.ru')){
                    #Прицепляем к существующем письму если отправитель из списка неполученных ответов
                        Foreach ($message_id in ($waiting_for_response.GetEnumerator() | Where-Object {$_.Value -eq $sended_address})){
                            Write-Verbose $_.Parent.Subject
                            Write-Verbose $($_.Item($i).FileName)
                            Write-Verbose $_.Parent.SenderEmailAddress
                            Write-Verbose $message_id.Key                        

                            #Удаляем полученное письмо
                            $_.Parent.Unread = $True
                            $_.Parent.Delete()
                
                            #Устанавливаем статус принятого ответа
                            $test, $return_status = WriteQueueStatus $SQLServer $SQLDatabase $SQLLogin $SQLPassword $message_id.Key $($_.Item($i).FileName) $temp_file
                            if ($return_status){
                                Pandion_Send ("Произошла ошибка установки статуса принятого ответа: $return_status")
                            }

                            
                        }  
                    }

                    If(Test-Path $temp_file){
                        Remove-Item $temp_file -Force
                    }
                }
            }
        }
    }
    

}

#Проверяем подключение к VPN
If (-NOT [bool](Ipconfig | ConvertTo-Encoding cp866 windows-1251 | Select-String "Подключение к СВК")){
    (Start-Process rasdial -NoNewWindow -ArgumentList $VPN_Name -PassThru -Wait).ExitCode
}
#Проверяем авторизацию на POP
If (TestPop $SVK_Login $SVK_Pass '192.168.19.4'){
    If (-Not (Telnet_Login $SVK_Login $SVK_Pass '192.168.19.20') -eq $true){
        Exit
    }Else{
        Write-Verbose "Connected via Telnet"
    }
}


Get_Sended_Status

    #Получаем данные по отправляемым файлам
    try
    {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

        #Set Connection String
        $SqlConnection.ConnectionString = “Server=$SQLServer;Database=$SQLDatabase;User ID = $SQLLogin; Password = $SQLPassword;”
        $SqlConnection.open()
    
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        #Записываем новый статус в поле
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandTimeout = 600000
        $SqlCmd.CommandText = "SELECT q.[queue_id],q.[sended_filename],q.[file_path],spr.[mail] FROM [ecdep].[dbo].[R_mgtu_queue] q INNER JOIN [ecdep].[dbo].[R_mgtu_spr] spr ON q.name_otch = spr.name_otch WHERE q.status='processed' order by q.id asc"
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
    
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet)
        
        If ($DataSet.Tables[0].Rows.Count -gt 0){

        # Для отправки сообщения требуется авторизация
        If (TestPop $SVK_Login $SVK_Pass '192.168.19.4'){
            If (-Not (Telnet_Login $SVK_Login $SVK_Pass '192.168.19.20') -eq $true){
                Exit
            }Else{
                Write-Verbose "Connected via Telnet"
            }
        }Else{
            Write-Verbose "Exist Connection"
        }

            Foreach ($row in $DataSet.Tables[0].Rows){
                
                #Готовим файлы
                $queue_id = $row['queue_id']
                $mail_to = $row['mail']
                $filename = $row['sended_filename']
                $current_folder = New-TemporaryDirectory
                $sended_filename = (Join-Path $current_folder $filename)
                Copy-Item -Path (Join-Path $Folder_OUT $row['file_path']) -Destination $sended_filename -Force
            
                #Создаем два экземпляра класса
                $att = New-object Net.Mail.Attachment($sended_filename)
                $att.ContentDisposition.Filename = $filename
                $att.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
                $mes = New-Object System.Net.Mail.MailMessage

                #Формируем данные для отправки
                $mes.From = New-Object System.Net.Mail.MailAddress($mail_from)
                $mes.To.Add($mail_to) 
                $mes.Subject = $queue_id 
                $mes.IsBodyHTML = $false 
                $mes.Body = $filename

                $mes.Attachments.Add($att)  

                #Создаем экземпляр класса подключения к SMTP серверу 
                $smtp = New-Object Net.Mail.SmtpClient($smtp_server, $smtp_port)

                #Создаем экземпляр класса для авторизации на сервере яндекса
                $smtp.Credentials = New-Object System.Net.NetworkCredential($slogin, $spass);

                #Отправляем письмо, освобождаем память
                $smtp.Send($mes) 

                Write-Verbose ("Отправлено {0} {1}" -f $queue_id, $filename)
                $att.Dispose()

                #Подчищаем за собой
                Remove-Item $current_folder -Recurse -Force

                #Записываем результат отправки
                $SqlCmd.CommandText = "UPDATE [{0}].[dbo].[R_mgtu_queue] SET status='sended', sended_time=GETDATE() where [queue_id]='{1}'" -f $SQLDatabase, $queue_id
                $rowsAffected = $SqlCmd.ExecuteNonQuery()
                If ($rowsAffected -gt 0){            
                    #Можно заканчивать работу
                }Else{
                    Pandion_Send ("Записи о результатах не выполнена для файла {0}" -f $filename)
                }     
            }
        }
    }  
    catch [Exception]
    {
        Write-Verbose $_.Exception.Message
        Pandion_Send ("Отчетность не отправлена {0} для файла {1}" -f $_.Exception.Message, $filename)
    }
    finally
    {
        $SqlConnection.Close()
        $SqlConnection.Dispose()
    }

