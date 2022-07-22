Clear-Host
$fi = Get-Content "C:\scripts\xml_para_csv_sql\dados.txt" #Ficheiro que contém os dados necessários
$nx = $fi[0] #caminho para os xmls
$lc = $fi[1] #caminho para onde vai o csv
$nc = $fi[2] #Nome da 1ª tabela sql
$sn = $fi[3] #Nome do 2 Csv
$SqlFile = "$lc\_$($nc)_SQLI.csv"

#Verifica se os ficheiros csv com as tabelas sql já foram criadas
$i = 1
do{   
    if (!(Test-Path $SqlFile)){$SqlFileI = "$lc\_$($nc)_SQLI.csv"}
    else{
        $SqlFileI = "$lc\_$($nc)_SQLI" + "_$($i)" + ".csv"
    }
    $i = $i + 1
}until(!(Test-Path $SqlFileI))

$i = 1
$SqlFile = "$lc\_$($sn)_SQLII.csv"

do{   
    if (!(Test-Path $SqlFile)){$SqlFileII = "$lc\_$($sn)_SQLII.csv"}
    else{
        $SqlFileII = "$lc\_$($sn)_SQLII" + "_$($i)" + ".csv"
    }
    $i = $i + 1
}until(!(Test-Path $SqlFileII))

#Convert os XMl's para csv para depois comparar com a tabela MAIN
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

#dados da tabela sql e da Database
$SQLServer = "srv-sigma\sqlexpress"
$SQLDBName = "FATURAS"
$tableSchema = "dbo" 
$tableName = $sn
$uid = "sa"
$pwd = "srvsigma"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd   
$delimiter = ";"

#Exporta as tabelas para csv 
$SqlQuery = "SELECT * from $SQLDBName.dbo.$nc;"  
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection  
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; user id = $uid; password = $pwd;"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand  
$SqlCmd.CommandText = $SqlQuery  
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter  
$SqlAdapter.SelectCommand = $SqlCmd   
$DataSet = New-Object System.Data.DataSet  
$SqlAdapter.Fill($DataSet)  
$DataSet.Tables[0] | export-csv -Delimiter $delimiter -Path $SqlFileI -NoTypeInformation -Encoding UTF8


Clear-Host

#compara as duas tabelas exportadas
$1sql = Import-Csv -Path $SqlFileI -Delimiter $delimiter -Encoding UTF8 
$2sql = Import-Csv -Path "$CsvFilePath$($File.BaseName).csv" -Delimiter $delimiter -Encoding UTF8 

$CompareCsvSql = Compare-Object -ReferenceObject $1sql -DifferenceObject $2sql -Property File_Path ,FileName, Fornecedor, Total 


#Exporta todas as linhas que a segunda tabela sql tem e a primeira tabela não tem
if($CompareCsvSql.sideindicator -eq '=>'){
Write-Warning "A tabela '$sn' tem $(($CompareCsvSql.sideindicator | Where-Object{$_ -eq '=>'}).count) linha(s) diferente(s) da tabela '$nc'."
($CompareCsvSql | Where-Object{$_.SideIndicator -eq '=>'}) | Select-Object -Property File_Path ,FileName, Fornecedor, Total | 
Export-Csv -LiteralPath "$lc\LinhasDaTabela_$sn.csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation 
}

#Avisa se não houver diferenças 
if($CompareCsvSql.sideindicator -eq $null){Write-Host "As tabelas '$nc' e '$sn' não têm diferenças."}

#Elimina os csv que foram criados com as tabelas sql
Remove-Item $SqlFileI


