
Clear-Host
#---------------------------(A)-------------------------------
#Dados do Ficheiro que vai ser exportado com linhas repetidas
$Path = "C:\scripts\compara_Tabelas2\"
$FileName = "FicheiroComLinhasRepetidas"
$FilePath = $Path + "\" + $FileName + ".csv"

#Verifica se o ficheiro csv com as linhas repetidas exportadas da tabela sql já foi criada
[Int]$i = 0  
do{   
    if (!(Test-Path $FilePath)){}
    else{$FilePath = $Path + "\" + $FileName + "_$($i)" + ".csv"}
    $i++
}until(!(Test-Path $FilePath))

#dados da tabela sql e da Database
$SQLServer = "srv-sigma\sqlexpress"
$SQLDBName = "FATURAS"
$SQLTableName = "XmlCsx"
$uid = "sa"           #login
$pwd = "srvsigma"     #password
$secureStringPwd = $pwd | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $uid, $secureStringPwd   
$delimiter = ";"

#----------------------------(B)-------------------------------
$SqlQuery = "select *
  from $($SQLTableName) a
  join ( select Fornecedor, documentDate, Total
           from $($SQLTableName) 
          group by Fornecedor, documentDate, Total
         having count(*) > 1 )  b
    on a.Fornecedor = b.Fornecedor
   and a.documentDate = b.documentDate
   and a.Total = b.Total"  
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection  
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; user id = $uid; password = $pwd;"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand  
$SqlCmd.CommandText = $SqlQuery  
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter  
$SqlAdapter.SelectCommand = $SqlCmd   
$DataSet = New-Object System.Data.DataSet  
$SqlAdapter.Fill($DataSet)  
$DataSet.Tables[0] | export-csv -Delimiter $delimiter -Path $FilePath -NoTypeInformation -Encoding UTF8

Clear-Host

if((Test-Path $FilePath) -eq $true){Write-Host "O ficheiro csv com linhas repetidas foi exportado para '$($FilePath)'."}

