# Laptop_Borrower_Alerts

Create Hidden Folder:

Creates a directory C:\SLU_Pius and hides it using the attrib command.
Alert Script (Alert.ps1):

## Functions:
TaskschedulerExist: Checks if a specific task exists in the Task Scheduler.
TaskSchedulerFolderExists: Checks if a specified folder exists in the Task Scheduler.
CreateFolder: Creates a new folder in the Task Scheduler if it doesn't exist.
CreatePhaseOneTaskScheduler, CreatePhaseTwoTaskScheduler, CreatePhaseThreeTaskScheduler: These functions create tasks for different phases with specific triggers and actions.
PhaseOneTwo: Displays a message box (for Phase One and Phase Two).
PhaseThree: Displays a full-screen message (for Phase Three).
SetupSchedulers: Sets up the initial folder and Phase One task scheduler.
Main Function:
Retrieves the username and asset name (computer name).
Sets up logging.
Ensures the existence of a checkoutDate.txt file to track the laptop checkout date.
Checks and sets up the necessary task scheduler folders and tasks based on the current date and the phase.
Runs the appropriate task (Phase One, Two, or Three) based on the number of days since the laptop was checked out.
Delete Task Scheduler Script (DeleteTaskScheduler.ps1):

## Functions:
TaskSchedulerFolderExists: Checks if a specified folder exists in the Task Scheduler.
DeleteTasks: Deletes all tasks within a specified folder in the Task Scheduler.
DeleteDateFile: Deletes the checkoutDate.txt file.
DeleteDirectory: Deletes the C:\SLU_Pius directory.
Main Function:
Deletes the checkoutDate.txt file and the hidden directory C:\SLU_Pius.
Deletes all tasks in the PiusLibrary folder in the Task Scheduler.
Deletes the PiusLibrary folder in the Task Scheduler if it is empty.
Use Case
The script is used by an institution (likely a library) to remind users to return borrowed laptops. It sets up three phases of reminders:

Phase One: Displays a reminder message upon user logon.
Phase Two: Displays a reminder message every hour if the laptop is overdue by 7-14 days.
Phase Three: Displays a full-screen reminder message every 5 minutes if the laptop is overdue by more than 14 days.
These phases help ensure that users are regularly reminded to return the laptop, and the intensity of the reminders increases as the overdue period extends. The script also includes a cleanup script to remove all task schedulers and related files when they are no longer needed.
