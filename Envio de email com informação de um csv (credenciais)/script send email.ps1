$OUpath = 'OU=Utilizadores (Portal Colaborador),DC=CMPOMBAL,DC=PT'
$Userinf = Get-ADUser -Filter * -SearchBase $OUpath -Properties EmailAddress,Name,samaccountname | select UserPrincipalName,EmailAddress, Name

$i = 0
$infile = 'C:\temp\sendmail\teste2.csv'

$P = Import-Csv $infile -DeLimiter ";" |
Select-Object 'SAM', 'OU', 'Password', 'SMTP', 'First Name', 'Last Name', 'Nr', 'EMAIL';  
$P.GetType() | Format-Table -AutoSize

foreach($User in $P){

$EmailFrom = “bernaslopes44@outlook.com”

$EmailTo = $P.SMTP[$i] , $P.EMAIL[$i]

$Subject = "Credenciais de acesso ao portal do colaborador (número: " +$P.Nr[$i] + ")"

$Body =  "Colaborador(a), " +$P.'First Name'[$i]+ " " +$P.'Last Name'[$i] + "!`n`n Enviamos o seu dados de acesso ao portal do colaborador para que possa gerir a sua a assiduidade (ferias e faltas) `n Endereço: http://portaldocolaborador.cm-pombal.pt`n Nome de utilizador: " + $P.SAM[$i] + "`n Palavra-Passe: " + $P.Password[$i] + "`n`n Para qualquer questão, contacte a Natlhalie Fajardo pelos contactos normais: `n Email: nathalie.fajarddo@cm-pombal.pt `n Ext: 1416 `n`n Com os melhores cumprimentos,`n`n A Equipa da Informática do Município de Pombal `n`n[Assinatura e-mail]"

$SMTPServer = “smtp.outlook.com”

$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)

$SMTPClient.EnableSsl = $true

$SMTPClient.Credentials = New-Object System.Net.NetworkCredential(“bernaslopes44@outlook.com”, "T3st_123");
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)
$i++
}