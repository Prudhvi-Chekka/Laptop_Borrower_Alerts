$folderPath = "C:\SLU_Pius"
New-Item -Path $folderPath -ItemType Directory -Force
attrib +s +h $folderPath

# nagScript content
$nagScriptContent = @'
param (
    [string]$phase
)


# Function to check if a Task Scheduler exist
function TaskschedulerExist {
    param (
                [string] $folderName,
                [string] $taskName
                )

    # Get the list of all tasks on the system
    $allTasks = Get-ScheduledTask | Get-ScheduledTaskInfo

    $taskExists = $allTasks | Where-Object { $_.TaskName -eq $taskName -and $_.TaskPath -eq "\$folderName\"}
    LogMessage "#################### Inside TaskschedulerExist function ####################"
    if ($taskExists) {
        LogMessage "$taskName Task already exist in the $folderName folder of Task scheduler"
        return $true
    } else {
        LogMessage "$taskName Task doesn't exist in the $folderName folder of Task scheduler"
        return $false
    }
    LogMessage "#################### End TaskschedulerExist function ####################"
}

# Function to check if a folder exist in Task Scheduler
function TaskSchedulerFolderExists {
    param (
        [string]$folderName
    )

    # Create a new Task Scheduler COM object
    $taskScheduler = New-Object -ComObject "Schedule.Service"
    $taskScheduler.Connect()

    # Get the root folder of the Task Scheduler
    $rootFolder = $taskScheduler.GetFolder("\")
    
    # Check if the specified folder exists
    try {
        $folder = $rootFolder.GetFolder($folderName)
        return $true
    } catch {
        return $false
    }
}

#Function to create a new folder in Task Scheduler
function CreateFolder {
    param (
            [string] $folderName
            )
    
    # Call the function to check if the folder exists
    if (TaskSchedulerFolderExists -folderName $folderName) {
        LogMessage "$folderName folder already exist in the Task scheduler"
    } else {
        LogMessage "$folderName folder doesn't exist in the Task scheduler"
        # Define the path and name of the new folder
        $folderPath = "\" + $folderName
        LogMessage "#################### Inside CreateFolder function ####################"
        # Create a new task folder
        try {
            $taskService = New-Object -ComObject "Schedule.Service"
            $taskService.Connect()

            $rootFolder = $taskService.GetFolder("\")
            $newFolder = $rootFolder.CreateFolder($folderPath)

            LogMessage "Task folder '$folderPath' created successfully."
        } catch {
            LogMessage "Error creating task folder: $_"
        }
        finally {
        LogMessage "#################### End CreateFolder function ####################"
    }
    }
}

#Function to create phase one task scheduler
function CreatePhaseOneTaskScheduler {
    param (
            [string]$folderName,
            [string]$taskName,
            [string]$currentFilepath
          )
    LogMessage "#################### Inside CreatePhaseOneTaskScheduler function ####################"
    try{
        $folderPath = "\" + $folderName
        # Define the action for the scheduled task
        $scriptArguments = "-ExecutionPolicy Bypass -WindowStyle hidden -File `"$currentFilepath`" `"`"$taskName`"`""
        $action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $scriptArguments

        # Specify the trigger to run at logon for any user
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        $description = "This Task Scheduler is to run the popup immediately after the user logon"

        $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 4) -AllowStartIfOnBatteries -WakeToRun -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    
        $groupName = "SLU\Domain Users"
        $principal =New-ScheduledTaskPrincipal -GroupId $groupName -RunLevel Highest

        Register-ScheduledTask -TaskName $taskName -TaskPath $folderPath -Principal $principal -Action $action -Description $description -Settings $settings  -Trigger $trigger
        LogMessage "Scheduled task '$taskName' created successfully inside folder '$folderPath'."
    } catch {
            LogMessage "Error creating task: $_"
        }
    finally {
        LogMessage "#################### End CreatePhaseOneTaskScheduler function ####################"
    }
}


#Function to create phase two task scheduler
function CreatePhaseTwoTaskScheduler {
    param (
            [string] $folderName,
            [string] $taskName,
            [string] $currentFilepath
            )
    LogMessage "#################### Inside CreatePhaseTwoTaskScheduler function ####################"
    try{
        $folderPath = "\" + $folderName
        $description = "This Task Scheduler is to generate the pop-up for every 1 hour"
    
        $scriptArguments = "-ExecutionPolicy Bypass -WindowStyle hidden -File `"$currentFilepath`" `"`"$taskName`"`""
        $action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $scriptArguments

        $frequencyInMin = 60
        # Define the trigger for the scheduled task to repeat every 1 hour
        $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date) `
        -RepetitionInterval (New-TimeSpan -Minutes $frequencyInMin) `
        -RepetitionDuration (New-TimeSpan -Days (365 * $frequencyInMin))
    
        $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries:$true

        $groupName = "SLU\Domain Users"
        $principal =New-ScheduledTaskPrincipal -GroupId $groupName -RunLevel Highest

        Register-ScheduledTask -TaskName $taskName -TaskPath $folderPath -Principal $principal -Action $action -Description $description -Settings $TaskSettings  -Trigger $trigger
    
        LogMessage "Scheduled task '$taskName' created successfully inside folder '$folderPath'."
    } catch {
            LogMessage "Error creating task: $_"
        }
    finally {
        LogMessage "#################### End CreatePhaseTwoTaskScheduler function ####################"
    }
}

#Function to create phase three task scheduler
function CreatePhaseThreeTaskScheduler {
    param (
                [string] $folderName,
                [string] $taskName,
                [string] $currentFilepath
                )
    LogMessage "#################### Inside CreatePhaseThreeTaskScheduler function ####################"
    try{
        $folderPath = "\" + $folderName
        $description = "This Task Scheduler is to generate the pop-up for every 5 mins"
        $scriptArguments = "-ExecutionPolicy Bypass -WindowStyle hidden -File `"$currentFilepath`" `"`"$taskName`"`""
        # Define the action for the scheduled task
        $action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -Argument $scriptArguments
        $frequencyInMin = 5
        # Define the trigger for the scheduled task to repeat every 5 mins
        $trigger = New-ScheduledTaskTrigger `
        -Once `
        -At (Get-Date) `
        -RepetitionInterval (New-TimeSpan -Minutes $frequencyInMin) `
        -RepetitionDuration (New-TimeSpan -Days (365 * $frequencyInMin))
    
        $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries:$true

        $groupName = "SLU\Domain Users"
        $principal = New-ScheduledTaskPrincipal -GroupId $groupName -RunLevel Highest

        Register-ScheduledTask -TaskName $taskName -TaskPath $folderPath -Principal $principal -Action $action -Description $description -Settings $TaskSettings  -Trigger $trigger

        Write-Host "Scheduled task '$taskName' created successfully inside folder '$folderPath'."
    } catch {
            LogMessage "Error creating task: $_"
        }
    finally {
        LogMessage "#################### End CreatePhaseThreeTaskScheduler function ####################"
    }
}

#Function to setup schedulers
function SetupSchedulers {
    param (
    [string] $folderName,
    [string] $firstTaskName
    )
    LogMessage "#################### Inside SetupSchedulers function ####################"
    # Call the function to create a folder in Task Scheduler
    CreateFolder -folderName $folderName
    
    # Call the function to create the scheduled task for Phase One
    CreatePhaseOneTaskScheduler -folderName "$folderName" -taskName $firstTaskName -currentFilepath $currentFilepath
    LogMessage "#################### End SetupSchedulers function ####################"
}

# This function creates a pop-up and can be used for phase 1 and phase 2
function PhaseOneTwo {
    param (
            [string] $asset_name,
            [string] $username
            )
    LogMessage "#################### Inside PhaseOneTwo function ####################"
    try{
        Add-Type -AssemblyName System.Windows.Forms

        # Create a new form
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "URGENT NOTICE - Reminder"
        $form.BackColor = [System.Drawing.Color]::Blue
        $form.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
        $form.Size = New-Object System.Drawing.Size(1000, 500)
        $form.FormBorderStyle = "FixedToolWindow"
        $form.StartPosition = "CenterScreen"
        $form.TopMost = $true
        $form.Visible = $false

        # Create a label for Asset return message
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(100, 70)
        $label.Size = New-Object System.Drawing.Size(800, 220)
 
        $label.Text = "Hello, ${username}, the return date for this laptop has now passed.`nPlease return the laptop (${asset_name}) and its power adapter to Pius Library as soon as possible. If you still need a laptop, you may check out another one at that time, depending on availability. `n`n`n`n If you have any questions, please call us at 314-977-3087 or email us at piuscirc@slu.edu.`nThank you!"
        $label.ForeColor = "White"
        $label.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter # Center text vertically and horizontally
        $form.Controls.Add($label)

        # Create an OK button
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(450, 380)
        $button.Size = New-Object System.Drawing.Size(100, 30)
        $button.Text = "OK"
        $button.BackColor= "White"
        $button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
        $button.Add_Click({ $form.Close() })
        $form.Controls.Add($button)

        # Display the form as a dialog
        $form.ShowDialog()
        LogMessage "Successfully completed function 'PhaseOneTwo'"
    } catch {
            LogMessage "Error occured while running the function 'PhaseOneTwo': $_"
        }
    finally {
        LogMessage "#################### End PhaseOneTwo function ####################"
    }

}

# This function creates a fullscreen display message in phase 3
function phaseThree {
    param (
        [string] $asset_name,
        [string] $username
    )
    LogMessage "#################### Inside PhaseThree function ####################"
    try{
        Add-Type -AssemblyName PresentationFramework

        # Create a new WPF window
        $window = New-Object System.Windows.Window
        $window.WindowStyle = 'None'  # Remove window borders
        $window.WindowState = 'Maximized'  # Maximize the window to full screen
        $window.Background = [System.Windows.Media.Brushes]::Blue
        $window.Topmost = $true
        $window.Opacity = 0.9

        # Create a TextBlock for the message
        $message = New-Object System.Windows.Controls.TextBlock
        $message.Text = "Hello, ${username}, the return date for this laptop has now passed.`nPlease return the laptop (${asset_name}) and its power adapter to Pius Library as soon as possible.`n If you still need a laptop, you may check out another one at that time, depending on availability. `n`n If you have any questions, please call us at 314-977-3087 or email us at piuscirc@slu.edu.`nThank you!"
        $message.FontSize = 30
        $message.TextAlignment = 'Center'

        # Set the initial font size
        $initialFontSize = $message.FontSize

        # Calculate the maximum font size based on screen height
        $maxFontSize = [Math]::Floor($window.ActualHeight / 10)
    
        # Set a minimum font size
        $minFontSize = 25

        # Loop to adjust font size based on content size
        while ($message.ActualWidth -gt $window.ActualWidth - 100 -or $message.ActualHeight -gt $window.ActualHeight - 100) {
            # Reduce font size, but don't go below the minimum font size
            $message.FontSize = [Math]::Max($message.FontSize - 1, $minFontSize)
        
            # Break out of the loop if the font size becomes too small
            if ($message.FontSize -eq $minFontSize) {
                break
            }
        }

        $message.Foreground = 'White'

        # Set the TextBlock's alignment to Center
        $message.HorizontalAlignment = 'Center'
        $message.VerticalAlignment = 'Center'

        # Set the window's content to the message TextBlock
        $window.Content = $message

        # Set the window's alignment to Center
        $window.WindowStartupLocation = 'CenterScreen'

        # Show the window
        $window.ShowDialog()
        LogMessage "Successfully completed function 'PhaseThree'"
    } catch {
            LogMessage "Error occured while running the function 'PhaseThree': $_"
        }
    finally {
        LogMessage "#################### End PhaseThree function ####################"
    }
}

function Main {
    param (
        [string] $phase
    )
    # Get the logged-in user's username
    $username = $env:USERNAME
    $asset_name = $env:COMPUTERNAME
    
    # Variable declaration
    $logDirectory = "C:\SLU_Pius\Logs"
    $fileName = "checkoutDate.txt"
    $checkoutFilepath = "C:\SLU_Pius\$fileName"
    $currentFilename ="NagScript.ps1"
    $currentFilepath = "C:\SLU_Pius\$currentFilename"
    $firstTaskName = "PhaseOne"
    $secondTaskName = "PhaseTwo"
    $thirdTaskName = "PhaseThree"
    $folderName = "PiusLibrary"
    $groupName = "LateNotebookNag"

    # Create Logs directory if it doesn't exist
    if (-not (Test-Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force
    }

    $logFilename = Get-Date -Format "yyyyMMddHHmmss"
    $filetype = ".log"
    $logFile = $logFilename + $filetype
    $logFilePath = "$logDirectory\$logFile"
    # Function to log messages to a file
    function LogMessage {
        param (
            [string] $message
        )
        $logTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFilePath -Value "[$logTime] $message"
    }

    # Log the start of the Main function
    LogMessage "########## Inside Main function ##########"
    LogMessage "Main function started. Phase: $phase"

    # Check whether the "checkoutDate" file exists
    if (-not (Test-Path $checkoutFilepath)) {
        # File does not exist, create it
        New-Item -ItemType File -Path $checkoutFilepath -Force
        $currentDate = Get-Date -Format "MM-dd-yyyy"
        Set-Content -Path $checkoutFilepath -Value $currentDate
        LogMessage "Created checkoutDate file with date: $currentDate"
    }
    else {
        LogMessage "The checkoutDate file already exists."
    }

    # Call the function to check if the folder exists
    if (TaskSchedulerFolderExists -folderName $folderName)
    {
        if (-not (TaskschedulerExist -folderName $folderName -taskName $firstTaskName))
        {
            CreatePhaseOneTaskScheduler -folderName "$folderName" -taskName "$firstTaskName" -currentFilepath $currentFilepath
        }
        LogMessage "The folder '$folderName' exists in Task Scheduler."
    }
    else {
        SetupSchedulers -folderName $folderName -firstTaskName $firstTaskName
        LogMessage "Created the Task Scheduler '$folderName' folder and PhaseOne task."
    }
    
    # Read the date from the file
    $dateString = Get-Content -Path $checkoutFilepath
    LogMessage "Date string from the file: $dateString"

    # Convert the date string to a DateTime object
    $date = [DateTime]::ParseExact($dateString, "MM-dd-yyyy", $null)
    LogMessage "Parsed date: $date"

    # Get the current date
    $currentDate = Get-Date
    LogMessage "Current date: $currentDate"
    
    # Calculate the difference in days
    $daysDifference = ($currentDate - $date).Days
    LogMessage "Days difference: $daysDifference and phase: $phase"

    # Choose the phase of scheduler to run based on the Overdue date
    if ($phase -eq $firstTaskName) {
        # Calling the pop-up function
        LogMessage "Running Phase One."
        PhaseOneTwo -asset_name "$asset_name" -username "$username"
    }
    if ($daysDifference -gt 7 -and $daysDifference -lt 15) {
        if ((TaskschedulerExist -folderName $folderName -taskName $secondTaskName) -and ($phase -eq $secondTaskName)){
            # Calling the pop-up function
            LogMessage "Running Phase Two."
            PhaseOneTwo -asset_name "$asset_name" -username "$username"
        }
        else{
            try{
                # Call the function to create the scheduled task for Phase Two
                CreatePhaseTwoTaskScheduler -folderName "$folderName" -taskName $secondTaskName -currentFilepath $currentFilepath
                LogMessage "Phase Two Task Scheduler created."
            }
            catch{
                LogMessage "Phase Two Task Scheduler already exists."
            }
        }
    }
    elseif ($daysDifference -gt 14) {
        if ((TaskschedulerExist -folderName $folderName -taskName $thirdTaskName) -and ($phase -eq $thirdTaskName)){
            # Calling the pop-up function
            LogMessage "Running Phase Three."
            PhaseThree -asset_name "$asset_name" -username "$username" 
        }
        else{
            try{
                # Call the function to create the scheduled task  for Phase Three
                CreatePhaseThreeTaskScheduler -folderName "$folderName" -taskName $thirdTaskName -currentFilepath $currentFilepath
                LogMessage "Phase Three Task Scheduler created."
            }
            catch{
                LogMessage "Phase Three Task Scheduler already exists."
            }
        }        
    }     
    LogMessage "########## End Main function ##########"
}

# Call the Main function
$phase = "PhaseOne" # Change the phase as needed
Main -phase $phase



'@


# DeleteTaskScheduler Content
$DeleteTaskSchedulerContent =@'
# Function to check if a folder exist in Task Scheduler
function TaskSchedulerFolderExists {
    param (
        [string]$folderName
    )

    # Create a new Task Scheduler COM object
    $taskScheduler = New-Object -ComObject "Schedule.Service"
    $taskScheduler.Connect()

    # Get the root folder of the Task Scheduler
    $rootFolder = $taskScheduler.GetFolder("\")
    
    # Check if the specified folder exists
    try {
        $folder = $rootFolder.GetFolder($folderName)
        return $true
    } catch {
        return $false
    }
}

# Function to delete all the tasks in the tasck scheduler folder
function DeleteTasks {
    param (
        [string] $folderName
    )

    # Create a new instance of the Task Scheduler COM object
    $service = New-Object -ComObject Schedule.Service

    # Connect to the local machine
    $service.Connect()

    # Get the folder from the Task Scheduler root folder
    $rootFolder = $service.GetFolder("\")
    $folder = $rootFolder.GetFolder($folderName)

    # Get all tasks within the specified folder
    $tasks = $folder.GetTasks(0)

    # Remove each task within the folder
    foreach ($task in $tasks) {
        $folder.DeleteTask($task.Name, 0)
    }
    Write-Host "All tasks in folder '$folderName' have been deleted."
}

# Function to delete the Checkout.txt file
function DeleteDateFile {
    param (
            [string] $filePath
        )

    if (Test-Path $filePath) {
        Remove-Item -Path $filePath -Force
        Write-Host "File '$filePath' deleted succesfully."
    }
    else {
        Write-Host "File '$filePath' not found."
    }
}
function DeleteDirectory {
    param (
            [string] $Directory
        )

    # Check if the folder exists
    if (Test-Path $Directory -PathType Container) {
        # Delete the folder
        Remove-Item -Path $Directory -Recurse -Force
        Write-Host "Directory folder deleted successfully."
    } else {
        Write-Host "Directory folder does not exist."
    }
}

function Main {
    # Variable declaration
    $folderName = "PiusLibrary"
    $Directory = "C:\SLU_Pius\"
    $fileName = "checkoutDate.txt"
    $checkoutFilepath = $filepath + $fileName

    # Delete the Checkout date file
    #DeleteDateFile -filePath $checkoutFilepath
    DeleteDirectory -Directory $Directory
    # Call the function to check if the folder exists
    if (TaskSchedulerFolderExists -folderName $folderName) {
        # Delete all the tasks in the folder
        DeleteTasks -folderName $folderName

        # Create a new instance of the Task Scheduler COM object
        $scheduleObject = New-Object -ComObject Schedule.Service

        # Connect to the local machine
        $scheduleObject.Connect()

        # Get the root folder of the Task Scheduler
        $rootFolder = $scheduleObject.GetFolder("\")

        # Try to get the specified folder
        $folder = $rootFolder.GetFolder($folderName)

        if ($folder -ne $null) {
            # Check if the folder is empty
            $tasks = $folder.GetTasks(0)
            $rootFolder.DeleteFolder($folderName, 0)
            try {
                if ($tasks.Count -eq 0) {
                    # Remove the folder if it is empty
                    $rootFolder.DeleteFolder($folderName, 0)
                    Write-Host "The Total folder '$folderName' has been deleted."
                } else {
                    Write-Host "The folder '$folderName' is not empty. Please delete all tasks within the folder before removing it."
                }
            }
            catch {
                Write-Host "The folder '$folderName' has been deleted."
            }
            
        }
    }
    else {
        Write-Host "The folder '$folderName' does not exist in Task Scheduler."
       
    }
}

# Call the Main function
Main

'@

$nagScriptPath = Join-Path $folderPath "nagScript.ps1"
$DeleteTaskSchedulerPath = join-path $folderPath "DeleteTaskScheduler.ps1"


Write-Host "Folder '$folderPath' created and hidden."

Set-Content -Path $nagScriptPath -Value $nagScriptContent
Write-Host "File 'nagScript.ps1' created in '$folderPath'."

Set-Content -Path $DeleteTaskSchedulerPath -Value $DeleteTaskSchedulerContent
Write-Host "File 'DeleteTaskScheduler.ps1' created in '$folderPath'."


# Run the script using the call operator
& $nagScriptPath
