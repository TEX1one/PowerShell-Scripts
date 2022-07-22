Import-Module ActiveDirectory
  
$filepath = Read-Host -Prompt "C:\temp\users.csv" 
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force

$users = Import-Csv -Delimiter ";" $filepath

ForEach($user in $users){

    $dname = $user.'Display Name'
    $fname = $user.'First Name'
    $lname = $user.'Last Name'
    $mailDomain = $user.'Maildomain'
    $SAM = $user.'SAM'
    $OUpath = $user.'OU'
    $Password = $user.'Password'
    $Description = $user.'Description'
    $SMTP = $user.'SMTP'

    New-ADUser -Name $dname -GivenName $fname -Surname $lname -UserPrincipalName $SAM -Path $OUpath -AccountPassword $securePassword -PasswordNeverExpires $true -Enabled $true -EmailAddress $SMTP -Description $Description -SamAccountName $SAM
    echo "A criar a conta do/a $fname $lname da OU $OUpath"
}