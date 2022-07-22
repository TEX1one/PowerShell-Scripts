Clear-Host
Install-Module SqlServer

$fi = Get-Content "C:\scripts\conversor_xml_csv\dados.txt" #Ficheiro que contém os dados necessários
$nx = $fi[0] #caminho para os xmls
$lc = $fi[1] #caminho para onde vai o csv
$nc = $fi[2] #Nome do csv
$sn = $fi[3] #Nome da tabela Sql
$ToProcess = Get-ChildItem -Path $nx -Filter '*.xml' -Recurse #vai buscar todos os ficheiros xml dentro da diretoria com recursividade
$DataExec = Get-Date -Format "dd/MM/yyyy"
$horaExec = Get-Date -Format "HH:mm"
$SqlFile = "$lc\$($sn)_SQL.csv"
$CF = "$lc\" + $nc + ".csv"
$CsvFilePath = "$lc\"
$csvDelimiter = ";"

#Se o ficheiro csv existir elimina o ficheiro
if((Test-Path $CF) -eq $true){
Remove-Item $CF
}

#Verifica se os ficheiros csv com as tabelas sql já foram criadas
do{   
    if (!(Test-Path $SqlFile)){$SqlFile = "$lc\$($sn)_SQL.csv"}
    else{$SqlFile = "$lc\$($sn)_SQL" + "$($i)" + ".csv"}
    $i = $i + 1
}until(!(Test-Path $SqlFile))

#dados da tabela sql e da Database
$serverName = "srv-sigma\sqlexpress"
$databaseName = "FATURAS"
$tableSchema = "dbo" 
$tableName = $sn
$username = "sa"
$password = "srvsigma"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd

#Cria o ficheiro csv com as informações dos ficheiros xmls
$CsvFiles = $ToProcess.Name -replace('.xml','.csv') | ForEach-Object {"$CsvFilePath$_"}

ForEach ($File in $ToProcess){

    [Xml]$Xml = Get-Content -Path $File.FullName

    $Xml.message.invoice | Select-Object -Property  @(
            @{Name='File_Path';Expression={($File).DirectoryName + '\'}}
            @{Name='FileName';Expression={$File.Name -split('.xml')}},
            @{Name='Fornecedor'; Expression={$Xml.message.invoice.seller.vatNumber}},
            @{Name='documentNumber'; Expression={$Xml.message.invoice.documentNumber}},
            @{Name='documentDate';Expression={$_.documentDate}},
            @{Name='CreationDate';Expression={$Xml.message.creationDateTime}},
            @{Name='DocType'; Expression={$Xml.message.invoice}},
            @{Name='Data_Exec';Expression={$DataExec}},
            @{Name='Hora_Exec';Expression={$horaExec}},
            @{Name='BARCODE'; Expression={$Xml.message.invoice.reference.InnerText}},
            @{Name='Description'; Expression={$Xml.message.invoice.lineItem.description}},
            @{Name='Total'; Expression={$Xml.message.invoice.totalPayableAmount}}

    ) | Export-Csv -Path "$CsvFilePath$($File.BaseName).csv" -NoTypeInformation -Encoding Unicode
}

if(!(Test-Path $CF)){

     Import-Csv -LiteralPath $CsvFiles | Export-Csv -Path $CF -NoTypeInformation -Delimiter $csvDelimiter -Encoding UTF8  
    
}else{

    Write-Host "Boas"
}

#Importa o ficheiro Csv criado
$data = Get-Content $CF -Encoding UTF8

#Verifica se a tabela sql já existe
$Return = $SQL = $dataTable = $null
$sql = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$sn'"
$dataTable = Invoke-Sqlcmd –ServerInstance $serverName -Credential $creds –Database $databaseName –Query $sql
if ($dataTable) {$return = $true}
else {$return = $false}
if($Return -eq $true){

    #Exporta a tabela sql para um ficheiro csv
    $checksql = Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName  -Credential $creds -Query "Select* from $sn"
    $checksql | Export-Csv $SqlFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation

    #Compara a tabela exportada com o ficheiro csv criado
    $compare = Compare-Object -ReferenceObject ($data | Select-Object -Skip 1 | Sort-Object ) -DifferenceObject (Get-Content $SqlFile | Select-Object -Skip 1 | Sort-Object) -IncludeEqual
    $CsvLines = $compare | Where-Object {$_.SideIndicator -eq '<='}
    
    #Se o ficheiro csv não tiver nenhuma linha que a tabela já tenha o script para
    if($CsvLines -eq $null){
        Write-Warning "Não foi encontrado nenhuma linha no ficheiro $CF que a tabela $sn não tivesse"
        foreach($File in $CsvFiles){
            if((Test-Path $File) -eq $true){
                Remove-Item -Path $File
            }
        }
        EXIT
    }

    #Se o ficheiro csv tiver linhas que a tabela não tenha exporta essa linha(s) para a tabela sql
    $CsvLinesToSql = $CsvLines.InputObject | ForEach-Object {$_.replace('";"',"','")} | ForEach-Object{$_.replace('"','')} | ForEach-Object{("'$_'")} 
    
    foreach($line in $CsvLinesToSql){
        $FileName = $line.Substring(0, $line.IndexOf(',')) 
        $FileName = $line.Replace("$FileName,",'')
        $FileName = $FileName.Substring(0, $FileName.IndexOf(',')) 
        $FileName
        Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName  -Credential $creds -Query "IF EXISTS (SELECT * FROM $sn WHERE FileName = $FileName)
                                                                                                      delete from $sn WHERE FileName = $FileName
                                                                                                      Insert Into dbo.$($sn) (File_Path, FileName, Fornecedor, documentNumber, documentDate, CreationDate, DocType, Data_Exec, Hora_Exec, BARCODE, Description, Total) VALUES ($line)

                                                                                                      IF NOT EXISTS (SELECT * FROM $sn WHERE FileName = $FileName)
                                                                                                      Insert Into dbo.$($sn) (File_Path, FileName, Fornecedor, documentNumber, documentDate, CreationDate, DocType, Data_Exec, Hora_Exec, BARCODE, Description, Total) VALUES ($line)"
    }

    Remove-Item $SqlFile
}

#Se a tabela não existir cria a tabela com toda a informação do csv
if($Return -eq $false){
    Import-Csv -Path $CF -Delimiter ';' -Encoding UTF8 | Write-SqlTableData -ServerInstance $serverName -DatabaseName $databaseName -SchemaName $tableSchema -TableName $tableName -Credential $creds -Force

}

#Elimina todas as linhas nulas da tabela sql
Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName  -Credential $creds -Query "DELETE FROM $sn where FileName is null"

 foreach($File in $CsvFiles){
    if((Test-Path $File) -eq $true){
        Remove-Item -Path $File
    }
}


