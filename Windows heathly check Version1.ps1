
#requires -version 2
<#
.SYNOPSIS
A script to restart automatic services on remote computers and send an email report.

.DESCRIPTION
This script imports a CSV file that contains a list of computer names and another CSV file that contains AD credentials.
It loops through each computer name and checks if it has any stopped automatic services.
If yes, it restarts the services and collects the status results.
If no, it collects the status results.
It outputs the service status for each computer to an HTML file with a beach background and a dragon, sword and phoenix frame.
It sends an email with the HTML file and the log file as attachments to a specified address.

.PARAMETER ComputerName
The path to the CSV file that contains the computer names.

.PARAMETER Credential
The path to the CSV file that contains the AD credentials.

.PARAMETER EmailAddress
The path to the CSV file that contains the email addresses.

.INPUTS
None

.OUTPUTS
An HTML file stored in C:\Windows\Temp\ComputerReport.html
A log file stored in C:\Windows\Temp\RestartServices.log

.NOTES
Version: 1.0
Author: Bing
Creation Date: 15/04/2023
Purpose/Change: Initial script development

.EXAMPLE
.\Restart-AutoService.ps1 -ComputerName C:\Computers.csv -Credential C:\Credentials.csv -EmailAddress C:\Emails.csv
#>

#--------------------------------------------------------- [Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Dot Source required Function Libraries
. "C:\Scripts\Functions\Logging_Functions.ps1"

#---------------------------------------------------------- [Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
$sLogPath = "C:\Windows\Temp"
$sLogName = "RestartServices.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#HTML File Info
$sHTMLPath = "C:\Windows\Temp"
$sHTMLName = "ComputerReport.html"
$sHTMLFile = Join-Path -Path $sHTMLPath -ChildPath $sHTMLName

#Email Info
$sEmailFrom = "admin@contoso.com"
$sEmailSubject = "Computer Service Status Report"

#----------------------------------------------------------- [Functions]------------------------------------------------------------

Function Get-ServiceStatus {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credential
    )

    Begin {
        Write-Verbose "Getting service status for $ComputerName"
    }

    Process {
        Try {
            #Get all automatic services that are not running on the remote computer using CIM session for better performance
            $CIMSession = New-CimSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
            $StoppedServices = Get-CimInstance -ClassName Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -CimSession $CIMSession -ErrorAction Stop

            #If there are any stopped services, restart them and collect the results using Invoke-CimMethod for better performance
            If ($StoppedServices) {
                Write-Verbose "Restarting stopped services on $ComputerName"
                $RestartResults = @()
                Foreach ($Service in $StoppedServices) {
                    #Create a custom object to store the service name and restart result
                    $RestartResult = New-Object -TypeName PSObject -Property @{
                        ServiceName = $Service.Name
                        RestartResult = $null
                    }

                    #Try to restart the service and catch any errors using Invoke-CimMethod for better performance
                    Try {
                        Invoke-CimMethod -InputObject $Service -MethodName StartService -CimSession $CIMSession -ErrorAction Stop | Out-Null
                        $RestartResult.RestartResult = "Success"
                    }
                    Catch {
                        $RestartResult.RestartResult = "Failed: $($_.Exception.Message)"
                    }

                    #Add the custom object to the array of results
                    $RestartResults += $RestartResult
                }

                #Return the array of results
                Return $RestartResults
            }
            #If there are no stopped services, return a message
            Else {
                Write-Verbose "No stopped services on $ComputerName"
                Return "No action required"
            }
        }
        Catch {
            #Return the error message if any exception occurs
            Write-Verbose "Error getting service status for $ComputerName: $($_.Exception.Message)"
            Return "Error: $($_.Exception.Message)"
        }
        Finally {
            #Remove the CIM session
            Remove-CimSession -CimSession $CIMSession
        }
    }

    End {
        Write-Verbose "Finished getting service status for $ComputerName"
    }
}

Function ConvertTo-HTMLReport {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [object[]]$ServiceStatus
    )

    Begin {
        Write-Verbose "Converting service status to HTML report"

        #Define the HTML header with a beach background and a dragon, sword and phoenix frame
        $HTMLHeader = @"
<html>
<head>
<style>
body {
    background-image: url('https://i.imgur.com/8sNz3wE.jpg');
    background-repeat: no-repeat;
    background-attachment: fixed;
    background-size: cover;
}
table, th, td {
    border: 1px solid black;
    border-collapse: collapse;
}
th, td {
    padding: 5px;
}
th {
    text-align: left;
}
</style>
</head>
<body>
<div style="position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);">
<img src="https://i.imgur.com/1uqTfXs.png" style="position:absolute; width:800px; height:600px; left:0px; top:0px;">
<table style="position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);">
"@

        #Define the HTML footer
        $HTMLFooter = @"
</table>
</div>
</body>
</html>
"@

        #Define the HTML content with a table of service status using StringBuilder for better performance
        $HTMLContent = New-Object -TypeName System.Text.StringBuilder
        [void]$HTMLContent.Append("<tr><th>Computer Name</th><th>Service Name</th><th>Restart Result</th></tr>")
    }

    Process {
        Foreach ($Status in $ServiceStatus) {
            #If the status is an array of custom objects, loop through each object and add a table row
            If ($Status -is [object[]]) {
                Foreach ($Result in $Status) {
                    [void]$HTMLContent.Append("<tr><td>$($Result.PSComputerName)</td><td>$($Result.ServiceName)</td><td>$($Result.RestartResult)</td></tr>")
                }
            }
            #If the status is a string, add a table row with the string as the last column
            ElseIf ($Status -is [string]) {
                [void]$HTMLContent.Append("<tr><td>$($Status.PSComputerName)</td><td></td><td>$Status</td></tr>")
            }
        }
    }

    End {
        Write-Verbose "Finished converting service status to HTML report"

        #Return the HTML report by concatenating the header, content and footer using StringBuilder for better performance
        Return ([System.Text.StringBuilder]::new($HTMLHeader) + $HTMLContent + $HTMLFooter).ToString()
    }
}

Function Send-EmailReport {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$HTMLFile,

        [Parameter(Mandatory=$true)]
        [string]$LogFile,

        [Parameter(Mandatory=$true)]
        [string[]]$EmailAddress,

        [Parameter(Mandatory=$true)]
        [string]$EmailFrom,

        [Parameter(Mandatory=$true)]
        [string]$EmailSubject
    )

    Begin {
        Write-Verbose "Sending email report to $EmailAddress"
    }

    Process {
        Try {
            #Create a mail message object with the HTML file and the log file as attachments using Send-MailMessage cmdlet for better performance
            Send-MailMessage -From $EmailFrom -To $EmailAddress -Subject $EmailSubject -Body (Get-Content -Path $HTMLFile -Raw) -BodyAsHtml -Attachments $HTMLFile,$LogFile

            #Return a success message
            Return "Email report sent successfully“

        }
        Catch {
            #Return an error message if any exception occurs
            Return "Error sending email report: $($_.Exception.Message)"
        }
    }

    End {
        Write-Verbose "Sending email report to $EmailAddress"
    }
}

#----------------------------------------------------------- [Main]------------------------------------------------------------

#Get the current date and time
$CurrentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

#Write the script start message to the log file
Write-Log -Message "Script $PSCommandPath started at $CurrentDateTime" -Path $sLogFile

#Import the CSV files that contain the computer names, the AD credentials and the email addresses
$Computers = Import-Csv -Path $ComputerName
$Credentials = Import-Csv -Path $Credential
$Emails = Import-Csv -Path $EmailAddress

#Convert the AD credentials to a PSCredential object
$PSCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Credentials.UserName, ($Credentials.Password | ConvertTo-SecureString -AsPlainText -Force)

#Create an empty array to store the service status for each computer
$ServiceStatus = @()

#Loop through each computer name in the CSV file using runspaces for better performance
$RunspacePool = [runspacefactory]::CreateRunspacePool(1,[int]$env:NUMBER_OF_PROCESSORS + 1)
$RunspacePool.Open()
$Jobs = @()
Foreach ($Computer in $Computers) {
    #Create a scriptblock with the Get-ServiceStatus function and pass the computer name and credential as parameters
    $ScriptBlock = {
        Param ($ComputerName,$Credential)
        Function Get-ServiceStatus {
            [CmdletBinding()]
            Param (
                [Parameter(Mandatory=$true)]
                [string]$ComputerName,

                [Parameter(Mandatory=$true)]
                [pscredential]$Credential
            )

            Begin {
                Write-Verbose "Getting service status for $ComputerName"
            }

            Process {
                Try {
                    #Get all automatic services that are not running on the remote computer using CIM session for better performance
                    $CIMSession = New-CimSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                    $StoppedServices = Get-CimInstance -ClassName Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -CimSession $CIMSession -ErrorAction Stop

                    #If there are any stopped services, restart them and collect the results using Invoke-CimMethod for better performance
                    If ($StoppedServices) {
                        Write-Verbose "Restarting stopped services on $ComputerName"
                        $RestartResults = @()
                        Foreach ($Service in $StoppedServices) {
                            #Create a custom object to store the service name and restart result
                            $RestartResult = New-Object -TypeName PSObject -Property @{
                                ServiceName = $Service.Name
                                RestartResult = $null
                            }

                            #Try to restart the service and catch any errors using Invoke-CimMethod for better performance
                            Try {
                                Invoke-CimMethod -InputObject $Service -MethodName StartService -CimSession $CIMSession -ErrorAction Stop | Out-Null
                                $RestartResult.RestartResult = "Success"
                            }
                            Catch {
                                $RestartResult.RestartResult = "Failed: $($_.Exception.Message)"
                            }

                            #Add the custom object to the array of results
                            $RestartResults += $RestartResult
                        }

                        #Return the array of results
                        Return $RestartResults
                    }
                    #If there are no stopped services, return a message
                    Else {
                        Write-Verbose "No stopped services on $ComputerName"
                        Return "No action required"
                    }
                }
                Catch {
                    #Return the error message if any exception occurs
                    Write-Verbose "Error getting service status for $ComputerName: $($_.Exception.Message)"
                    Return "Error: $($_.Exception.Message)"
                }
                Finally {
                    #Remove the CIM session
                    Remove-CimSession -CimSession $CIMSession
                }
            }

            End {
                Write-Verbose "Finished getting service status for $ComputerName"
            }
        }

        #Call the Get-ServiceStatus function and return the result with the computer name as a property
        Get-ServiceStatus -ComputerName $ComputerName -Credential $Credential | Add-Member -MemberType NoteProperty -Name PSComputerName -Value $ComputerName -PassThru 
    }

    #Create a PowerShell object


    $PowerShell = [powershell]::Create().AddScript($ScriptBlock).AddParameters(@{ComputerName=$Computer.Name;Credential=$PSCredential})

    #Assign the PowerShell object to the runspace pool
    $PowerShell.RunspacePool = $RunspacePool

    #Create a custom object to store the PowerShell object and the async result
    $Job = New-Object -TypeName PSObject -Property @{
        PowerShell = $PowerShell
        AsyncResult = $PowerShell.BeginInvoke()
    }

    #Add the custom object to the array of jobs
    $Jobs += $Job
}

#Loop through each job and get the end result
Foreach ($Job in $Jobs) {
    #Wait for the job to complete and get the output
    $JobOutput = $Job.PowerShell.EndInvoke($Job.AsyncResult)

    #Dispose the PowerShell object
    $Job.PowerShell.Dispose()

    #Add the job output to the array of service status
    $ServiceStatus += $JobOutput
}

#Remove the runspace pool
$RunspacePool.Close()
$RunspacePool.Dispose()

#Call the ConvertTo-HTMLReport function and save the result to the HTML file
ConvertTo-HTMLReport -ServiceStatus $ServiceStatus | Out-File -FilePath $sHTMLFile

#Loop through each email address in the CSV file using runspaces for better performance
$RunspacePool = [runspacefactory]::CreateRunspacePool(1,[int]$env:NUMBER_OF_PROCESSORS + 1)
$RunspacePool.Open()
$Jobs = @()
Foreach ($Email in $Emails) {
    #Create a scriptblock with the Send-EmailReport function and pass the email address as a parameter
    $ScriptBlock = {
        Param ($EmailAddress)
        Function Send-EmailReport {
            [CmdletBinding()]
            Param (
                [Parameter(Mandatory=$true)]
                [string]$HTMLFile,

                [Parameter(Mandatory=$true)]
                [string]$LogFile,

                [Parameter(Mandatory=$true)]
                [string]$EmailAddress,

                [Parameter(Mandatory=$true)]
                [string]$EmailFrom,

                [Parameter(Mandatory=$true)]
                [string]$EmailSubject
            )

            Begin {
                Write-Verbose "Sending email report to $EmailAddress"
            }

            Process {
                Try {
                    #Create a mail message object with the HTML file and the log file as attachments using Send-MailMessage cmdlet for better performance
                    Send-MailMessage -From $EmailFrom -To $EmailAddress -Subject $EmailSubject -Body (Get-Content -Path $HTMLFile -Raw) -BodyAsHtml -Attachments $HTMLFile,$LogFile

                    #Return a success message
                    Return "Email report sent successfully"
                }
                Catch {
                    #Return an error message if any exception occurs
                    Return "Error sending email report: $($_.Exception.Message)"
                }
            }

            End {
                Write-Verbose "Sending email report to $EmailAddress"
            }
        }

        #Call the Send-EmailReport function and return the result with the email address as a property
        Send-EmailReport -HTMLFile $using:sHTMLFile -LogFile $using:sLogFile -EmailAddress $EmailAddress -EmailFrom $using:sEmailFrom -EmailSubject $using:sEmailSubject | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $EmailAddress -PassThru 
    }

    #Create a PowerShell object

    $PowerShell = [powershell]::Create().AddScript($ScriptBlock).AddParameters(@{EmailAddress=$Email.Address})

    #Assign the PowerShell object to the runspace pool
    $PowerShell.RunspacePool = $RunspacePool

    #Create a custom object to store the PowerShell object and the async result
    $Job = New-Object -TypeName PSObject -Property @{
        PowerShell = $PowerShell
        AsyncResult = $PowerShell.BeginInvoke()
    }

    #Add the custom object to the array of jobs
    $Jobs += $Job
}

#Loop through each job and get the end result
Foreach ($Job in $Jobs) {
    #Wait for the job to complete and get the output
    $JobOutput = $Job.PowerShell.EndInvoke($Job.AsyncResult)

    #Dispose the PowerShell object
    $Job.PowerShell.Dispose()

    #Write the job output to the log file
    Write-Log -Message $JobOutput -Path $sLogFile
}

#Remove the runspace pool
$RunspacePool.Close()
$RunspacePool.Dispose()

#Get the current date and time
$CurrentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

#Write the script end message to the log file
Write-Log -Message "Script $PSCommandPath ended at $CurrentDateTime" -Path $sLogFile
