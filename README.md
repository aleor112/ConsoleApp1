# ConsoleApp1
1.If you want to stop the script, just close the console and the app should save how many tickets have been processed in state file
and generate an excel of the remaining ticket records that have been collected before closing.
2.Parameters regarding the number of tickets that should be saved in each excel file is in the Const file.
3.Url from which the app will start gathering tickets can be changed in appjson - urlWithPlaceholder (placeHolder uses the number in state
folder to tell servicenow from which ticket to start the processing. If there is no state folder it will start from the 1 ticket)
4.Configuration about number of tabs to load can be changed in Const file.
5.Errors are written to myappError.log, info messages in myapp.log - you can check if there have been problems with extracting information for some tickets in there.
