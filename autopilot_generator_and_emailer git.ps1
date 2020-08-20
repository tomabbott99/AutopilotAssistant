#############################################
# AutoPilot Hash Generator and Sender       #
# Created by: Tom Abbott                    #
# Date created: 18/08/2020                  #
#############################################


# Tests if the path exists, and creates it if it doesn't. 

$path = "C:\Autopilotscript"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
     }



# Sets active directory 
Set-Location C:\Autopilotscript

# Sets script execution policy for the scope of this script to unrestricted and and forces it without user input 
Set-ExecutionPolicy -scope Process unrestricted -force

# Installs autopilot generator script in ARCIT directory
Install-Script -Name Get-WindowsAutoPilotInfo -force

# Generates the hashfile with the name hash.csv
Get-WindowsAutoPilotInfo.ps1 -outputfile hash.csv

# Imports the contents of the hash file into a useable variable 
$hash = Import-Csv C:\autopilotscript\hash.csv 

# Gets the service tag of the laptop and imports it into a varialbe for use in the email subject line 
$serial = (Get-WMIObject -Class WIN32_SystemEnclosure -ComputerName $env:ComputerName).SerialNumber


# Defines the mailfrom address
$MailFrom = "TYPE MAIL FROM ADDRESS HERE"

# Defines the MailTo Address
$MailTo = "TYPE MAIL TO ADDRESS HERE"

# Defines the authentication address 
$Username = "TYPE MAIL AUTHENTICATION ADDRESS HERE - USUALLY THE SAME AS MAIL FROM"

# Defines the authentication password (CANNOT BE FOR AN MFA ENFORCED OR ENABLED ACCOUNT)
$Password = "PASSWORD FOR AUTH"

# Defines the smtp server address, we are using office 365
$SmtpServer = "smtp.office365.com"

# Defines the server port to send on, 587 is default but some use cases will use 25
$SmtpPort = "587"

# Defines the message subject line into a usable varialbe, calling the service tag variable and also getting the username 
$MessageSubject = "Test Email from $env:username on $serial" 

# Defines the messsage from,to subjet and body
$Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo

# Calls the previously filled subject variable 
$Message.Subject = $MessageSubject

# Calls the previously created hash variable imported from the file to generate the email body
$Message.Body = $hash

# Defines the attachment location 


# Attaches the file 

$currentAttachment = 'C:\Autopilotscript\hash.csv'

# Adds the attachment

$Message.Attachments.Add($currentAttachment)

# Fills the SMtP details in from the variables
$Smtp = New-Object Net.Mail.SmtpClient ($SmtpServer, $SmtpPort)

# Forces the use of SSL transmission (Required for Office 365 at time of writing) 
$Smtp.EnableSsl = $true

# Authenticates with the SMTP server 
$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)

# Sends the email 
$Smtp.Send($Message)


