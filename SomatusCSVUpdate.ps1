[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$csv = ''


Function doesUserExist($EmployeeEmail){
    if($EmployeeEmail){
                    $EmployeeEmail = $EmployeeEmail.ToString()
                    $azureaduser = Get-AzureADUser -All $true | Where-Object {$_.Userprincipalname -eq "$EmployeeEmail"}
                       #check if something found    
                       if($azureaduser){
                              Write-Host "User: $EmployeeEmail was found in $displayname AzureAD." -ForegroundColor Green
                             return $true
                             }
                             else{
                              Write-Host "User $EmployeeEmail was not found in $displayname Azure AD " -ForegroundColor Red
                             #$BadAddress.add($EmployeeEmail)
                             return $false
                             }
                        }
}

Function isModuleInstalled($module){
    try{
        Get-InstalledModule -Name $module -ErrorAction Stop
        return $true
    }
    catch [System.Exception]{
          #Write-host "Install Azure AD Module?"
          Install-module AzureAD
          try{
              Get-InstalledModule -Name $module -ErrorAction Stop
              return $true
          }
          catch [System.Exception]{
              return $false
          }

    }
}

Function getUserDisplayName($EmployeeEmail){
    $UserDisplayName = (Get-AzureAdUser -ObjectID $EmployeeEmail).displayname
    return $UserDisplayName
}

Function checkJobTitle($EmployeeEmail){
    $ADJobTitle = (Get-AzureAdUser -ObjectID $EmployeeEmail).jobtitle
    return $ADJobTitle
}

Function checkDepartment($EmployeeEmail){
     $ADDepartment = (Get-AzureAdUser -ObjectID $EmployeeEmail).department
     return $ADDepartment
}

Function checkManager($EmployeeEmail){
    $ADManager = (Get-AzureADUserManager -ObjectID $EmployeeEmail).UserPrincipalName
    return $ADManager
}

Function setJobTitle($JobTitle, $EmployeeEmail){
    Set-AzureADUser -ObjectID $EmployeeEmail -JobTitle $JobTitle

}

Function setDepartment($Department, $EmployeeEmail){
    Set-AzureADUser -ObjectID $EmployeeEmail -Department $Department
}

Function getManagerDisplayname($EmployeeEmail){
    $DisplayName = (Get-AzureADUserManager -ObjectId $EmployeeEmail).displayname
    return $DisplayName
}

Function setManager($Manager, $EmployeeEmail){
    Set-AzureADUserManager -ObjectId (Get-AzureADUser -ObjectID $EmployeeEmail).ObjectID  -RefObjectId (Get-AzureADUser -ObjectID $Manager).ObjectID
}

Function ReferenceJobTitle($ADJobTitle, $CSVJobTitle, $EmployeeEmail){
    
    if (!($ADJobTitle -eq $CSVJobTitle ))
    {
        setJobTitle $CSVJobTitle $EmployeeEmail
       return $true
    }
    return $false
}

Function ReferenceDepartment($ADDepartment, $CSVDepartment, $EmployeeEmail){
    if (!($ADDepartment -eq $CSVDepartment ))
    {
       setDepartment $CSVDepartment $EmployeeEmail
       return $true
    }
    return $false
}

Function ReferenceManager($ADManager, $CSVManager, $EmployeeEmail){
    if (!($ADManager -eq $CSVManager ))
    {
       setManager $CSVManager $EmployeeEmail
       return $true
    }
    return $false
}

Function CSVProcess(){
    $ExportCSV= [Environment]::GetFolderPath('MyDocuments') + "\Updated Employee Rsoter for ATS_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    $BadAddress = @()
    $Result=""
    $Output=@()
    
    $csv | ForEach-Object {

    $lastName = $_.'Last Name'
    $firstName = $_.'First Name'
    $EmployeeEmail = $_.'Employee Email'
    $JobTitle = $_.'Job Title'
    $Department = $_.'Department'
    $SupervisorName = $_.'Supervisor Name'
    $Manager = $_.'Supervisor Email Address'

    
        
        if(doesUserExist($EmployeeEmail)){
              

              #$UserDisplayName = getUserDisplayName($EmployeeEmail)
              $ADJobTitle = checkJobTitle($EmployeeEmail)
              $ADDepartment = checkDepartment($EmployeeEmail)
              $ADManager = checkManager($EmployeeEmail)
              

              $JobTitleChange = ReferenceJobTitle -ADJobTitle $ADJobTitle -CSVJobTitle $JobTitle -EmployeeEmail $EmployeeEmail
              $DepartmentChange = ReferenceDepartment -ADDepartment $ADDepartment -CSVDepartment $Department -EmployeeEmail $EmployeeEmail
              $ManagerChange = ReferenceManager -ADManager $ADManager -CSVManager $Manager -Employee $EmployeeEmail



              if($JobTitleChange -or $DepartmentChange -or $ManagerChange)
              {
                  $BadAddress += ExportToForm $EmployeeEmail $JobTitleChange $DepartmentChange $ManagerChange $ADJobTitle $ADDepartment $ADManager
              }

              $OutDepartment = checkDepartment($EmployeeEmail)
              $OutJobTitle = checkJobTitle($EmployeeEmail)
              $OutManagerDisplayName = getManagerDisplayname($EmployeeEmail)
              $OutManagerEmail = checkManager($EmployeeEmail)

              $Result=@{'Last name'=$lastName;'First Name'=$firstName;'Employee Email'=$EmployeeEmail;'Department'=$OutDepartment;'Job Title'=$OutJobTitle;'Supervisor Name'=$OutManagerDisplayName;'Supervisor Email Address'=$OutManagerEmail}              
              $Output= New-Object PSObject -Property $Result
              $Output | Select-Object 'Last Name','First Name','Employee Email',Department,'Supervisor Name', 'Supervisor Email Address' | Export-Csv -Path $ExportCSV -Notype -Append
        }
    }  
      return $BadAddress 
}

Function ExportToForm($EmployeeEmail, $BoolJob, $BoolDepartment, $BoolManager, $OldTitle, $OldDepartment, $OldManager){
    $JobTitleToText = checkJobTitle($EmployeeEmail)
    $DepartmentToText = checkDepartment($EmployeeEmail)
    $ManagerToText = checkManager($EmployeeEmail)

    if(!$BoolJob)        {$OldTitle = "----------"}

    if(!$BoolDepartment) {$OldDepartment = "----------"}

    if(!$BoolManager)    {$OldManager = "----------"}
    
    $item = New-Object PSObject
    $item | Add-Member -type NoteProperty -Name 'Email' -Value $EmployeeEmail
    $item | Add-Member -type NoteProperty -Name 'Old Job Title' -Value $OldTitle
    $item | Add-Member -type NoteProperty -Name 'Current Job Title' -Value $JobTitleToText
    $item | Add-Member -type NoteProperty -Name 'Old Department' -Value $OldDepartment
    $item | Add-Member -type NoteProperty -Name 'Current Department' -Value $DepartmentToText
    $item | Add-Member -type NoteProperty -Name 'Old Manager' -Value $OldManager
    $item | Add-Member -type NoteProperty -Name 'Current Manager' -Value $ManagerToText

    return $item
}

 function Read-MultiLineInputBoxDialog($BadAddress){
    $Output = $BadAddress | Out-String
    $WindowTitle = "List of users with field changes" 
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

    # Create the Label.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = "Final Result"

    # Create the TextBox used to capture the user's text.
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(10,40)
    $textBox.Size = New-Object System.Drawing.Size(575,200)
    #$textBox.AcceptsReturn = $true
    #$textBox.AcceptsTab = $false
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $Output
    $textBox.ReadOnly = $true
    
    # Create ticketbox Label.
    $ticketlabel = New-Object System.Windows.Forms.Label
    $ticketlabel.Location = New-Object System.Drawing.Size(185,252)
    $ticketlabel.Size = New-Object System.Drawing.Size(280,20)
    $ticketlabel.AutoSize = $true
    $ticketlabel.Text = "Ticket Number?"

    # Create the TicketBox
    $ticketbox = New-Object System.Windows.Forms.TextBox
    $ticketbox.Location = New-Object System.Drawing.Point(300,252)
    $ticketbox.Size = New-Object System.Drawing.Size(100,50)

    # Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(405,250)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "Send Email"
    $okButton.Add_Click({ $form.Tag = $textBox.Text; sendEmail $Output $ticketbox.Text; $form.Close() })
    #$okButton.Add_Click({ $form.Tag = $textBox.Text; Write-Host $ticketbox.text; $form.Close() })

    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(510,250)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Close"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })

    # Create the form.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(610,320)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true

    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($ticketlabel)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.Controls.Add(($ticketbox))

    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null  # Trash the text of the button that was clicked.
    #$form.ShowDialog()
    # Return the text that the user entered.
    return $form.Tag
}

Function SendEmail($Output, $TicketNumber){
    $smtpserver = "somatus-com.mail.protection.outlook.com"
    $from = "ats.jason@somatus.com"
    #from = $email = (Get-AzureADCurrentSessionInfo).account.id
    $emailaddress = "jason@myaligned.com"
    #emailaddress = helpdesk@myalignedit.com
    $subject = ""

    if($ticketnumber -eq [String]::Empty){
        $subject= "Office 365 Contact Info Update"
        }
        else{
        $subject= "#" + $ticketnumber + ": Office 365 Contact Info Update"
        }

    Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $Output

}

Function Main(){
    $message=[System.Windows.Forms.MessageBox]::Show("Would you like to process the .csv for updates to Azure AD?","Azure AD CSV Updater",[System.Windows.Forms.MessageBoxButtons]::OKCancel)
     switch ($message){
         "OK" {
             write-host "You pressed OK"
             StartProgram
         }
         "Cancel" {
             write-host "You pressed Cancel"
             # Enter some code
         }
     }
 }


Function getFileName(){
      $File = New-Object System.Windows.Forms.OpenFileDialog
      $File.initialDirectory = [Environment]::GetFolderPath('Desktop')
      $File.filter = "CSV (*.csv)| *.csv"
      $result = $File.ShowDialog()

      if ($result -eq "OK") {
         return  $File.FileName

      }
      return $null
 }

 Function StartProgram(){
     try{
     if(isModuleInstalled('AzureAD')){
         Connect-AzureAD -ErrorAction Stop
         $file = getFileName
         Write-Host "Producing File"
         $file
         if($file)
         {
             Write-Host "In Loop"
             $csv = Import-CSV $file
             $GUIOutput = CSVProcess
             #$GUIOutput
             Read-MultiLineInputBoxDialog($GUIOutput)
         }
         #$BadAddress
         Disconnect-AzureAD
     }
     }
     catch [Microsoft.Open.Azure.AD.CommonLibrary.AadAuthenticationFailedException]{
         Write-Host "Unsuccesful sign in or cancelled sign in, quitting program"
     }
     catch [System.Runtime.InteropServices.COMException]
     {
         Write-Host "Issue with file"
     }

 }
Main
