Clear-Host
#Ficheiro com os dados Sam, Password, FirstName, LastName e Email
$Emails = Import-Csv "C:\scripts\conversor_xml_csv\teste2.csv" -DeLimiter ";" -Encoding UTF8

#dados da conta de email que vai enviar os emails
$Email = "teste12345670542@hotmail.com"
$Password = "Teste123456789"
$secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential($Email, $secpasswd)

foreach($item in $Emails){
#Assunto do Email
$Subject = "Credenciais de acesso ao portal do colaborador (número: $($item.Nr))"

$Body = "Colaborador(a), $($item.FirstName) $($item.LastName)
#Body text here!!!!
#Nome de utilizador: $($item.Sam)
#Palavra-Passe: $($item.Password)  
#Com os melhores cumprimentos, 
#A Equipa da Informática do Município de Pombal 
#[Assinatura e-mail]"

#Enviar Email
    Send-MailMessage -from "<$($Email)>" `
                     -to "<$($item.Email)>" `
                     -subject $Subject `
                     -body $Body `
                     -Attachment "C:xxx\xxxx\xxxx" -Encoding utf8 -smtpServer smtp.outlook.com -Port xxx -UseSsl -Credential $creds

}
