<# Author : Shane Wallace 
Student Id: 010895068
#>


#Region Function 
function Show-Menu

    param (
        [string]$Title = "Main Menu",
        [string[]]$Options
    )

    Write-Host $Title
    Write-Host "================="

    # loop fort the size of the array options provided
    for ($i = 0; $i -lt $Options.Length; $i++ )
    {
        # Iterates over the array and write the number and text of selection
        Write-Host "$($i+1). $($Options[$i])"
    }
    # Prompts user fot input
    $inputSeletion = Read-Host "Select an option"

    # Validates the input
    if ($inputSeletion -match '^\d+$' -and $inputSeletion -le $Options.Length)
    {
        # Returns the selection
        return $inputSeletion
    }
    else 
    {
        # if the input is invalid, inform the user and re-prompt
        Write-Host -ForegroundColor DarkMagenta "'n*Invalid Selection. Please try again"
        Show-Menu -Title $Title -Options $Options
    }
#EndRegion



# Main Fuction

$menuOptions = @("Name of Log files","Alphabetically tabular ","CPU and Memory Usage ","Running Processes " ,"Exit")

$userInput = 1
Try{
    while($userInput -ne 5)
    {
        $userInput = Show-Menu -Title "Task1 - Please Select an option" -Options $menuOptions   
    
        switch($userInput)
        { 
            1
            {
                # Variable to hold the log file names
                $outputFiles =  "$PSScriptRoot\DailyLogs.txt" 

                # Timepstamp for the log file
                "TIMESTAMP: "+(Get-Date) | Out-File -FilePath "DailyLogs.txt" -Append

                # Writes the log file names to the output file
                Get-ChildItem -Path "$PSScriptRoot" -Filter "*.log" | Out-file $outputFiles -Append
            }
            2
            {
                # Variable to hold the log file names
                $outputFiles =  "$PSScriptRoot\AlphabeticalLogs.txt" 

                # Timepstamp for the log file
                "TIMESTAMP: "+(Get-Date) | Out-File -FilePath "AlphabeticalLogs.txt" -Append

                # Writes the log file names to the output file in alphabetical order
                Get-ChildItem -Path "$PSScriptRoot" -Filter "*.log" | Sort-Object Name | Out-file $outputFiles -Append
            }
            3
            {
                # Variable to hold the CPU and Memory usage
                $outputFiles =  "$PSScriptRoot\CPU_Memory_Usage.txt" 

                # Writes the CPU and Memory usage to the output file
                Get-Process | Select-Object Name, CPU, WS | Out-file $outputFiles -Append
            }
            4
            {
                # Variable to hold the running processes
                $outputFiles =  "$PSScriptRoot\Running_Processes.txt" 

                # Timepstamp for the log file
                "TIMESTAMP: "+(Get-Date) | Out-File -FilePath "Running_Processes.txt" -Append

                # Writes the running processes to the output file
                Get-Process | Select-Object Name, Id, CPU | Out-file $outputFiles -Append
            }
            5
            {
                # Exit 
                Write-Host "Thanks for running the script. Goodbye!" -ForegroundColor cyan
                break
            }


            
                
        }
    }
        
}

Catch [System.OutOfMemoryException] 
{
    Write-Host "Out of memory error" -ForegroundColor Red
}  
Catch
{
    $_
}                
finally 
{
    Write-Host "Thanks for running the script. Goodbye!" -ForegroundColor cyan
}
