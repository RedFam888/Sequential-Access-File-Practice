Option Strict On
Option Explicit On
Option Infer Off
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports System.DirectoryServices.ActiveDirectory
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Threading
Imports System.Windows.Forms.ComponentModel.Com2Interop
Imports System.Windows.Forms.Design

'SQL Practice
'Mason

'CREATING AND USING AN OUTPUT FILE----
'---
'---

'1. Use the StreamWriter class to declare a StreamWriter variable

'           Syntax
'{Dim|Private} streamWriterVariable As IO.StreamWriter

'           Examples
'Dim outFile As IO.StreamWriter
'Private outFile As IO.StreamWriter

'2. Use Either the CreateText method or the AppendText method to open the file. Assign the StreamWriter object created by
'   either method to the StreamWriter variable from Step 1

'           Syntax
'IO.File.CreateText(fileName)
'IO.File.AppendText(fileName)

'           Examples
'outFile = IO.File.CreateText("employee.txt")
'outFile = IO.File.AppendText("F:\Chap09\report.txt")

'3. Use either the Write method or the WriteLine method to write data to the file. Unlike the Write method, the WriteLine method
'   writes a newline character after the data

'           Syntax
'streamWriterVariable.Write(data)
'streamWriterVariable.WriteLine(data)

'           Examples                                        Result
'outFile.Write("Hello")                                     Hello
'outFile.WriteLine("Hello")                                 Hello
'                           In the WriteLine Example the Cursor would have moved to this line to Write data

'4. Close the output file

'           Syntax
'streamWriterVariable.Close()

'           Example
'outFile.Close()

'CREATING AND USING AN INPUT FILE----
'---
'---

'1. Use the StreamReader class to delcare a StreamReader variable

'           Syntax
'{Dim|Private} streamReaderVariable As IO.StremReader

'           Examples
'Dim inFile as IO.StreamReader
'Private infile As IO.StreamReader

'2. Use the Exists method to determine wether the input file exists

'           Syntax
'IO.File.Exists(fileName)

'           Examples
'If IO.File>exists("contestants.txt") Then

'3. If the Exists methos determines that the file exists, use the OpenText method to open the file for Input. Assign the
'   StreamReader object created by the method to the SteamReader variable from Step 1.

'           Syntax
'IO.File.OpenText(fileName)

'           Examples
'inFile = Io.File.OpenText("contestants.txt")

'4. To read the file, line by line, use the Peek and ReadLine methods along with a loop procedure. To read the entire
'   File all at once, use the ReadToEnd method

'           Syntax
'streamReaderVariable.Peek
'streamReaderVariable.ReadLine
'streamReaderVariable.ReadToEnd

'           Example Reading the File line byb line
'Dim lineOfText as String
'Do Until inFile.Peek = -1
'   lineOfText = inFile.ReadLine
'   [Instructions]
'Loop

'           Example Reading the File Completely to End
'txtContestantsResults.Text = inFile.ReadToEnd

'5. Close the Input File

'           Syntax
'streamReaderVAriable.Close()

'           Examples
'inFile.Close()
Public Class Form1

    'Class Level Array For Continents Application
    Private continents(6) As String
    Private lastSub As Integer = continents.GetUpperBound(0)

    'Class Level Variable for States Application
    Private states(49) As String

    'class level candy array
    Private candy(4) As Integer

    'class level boolean for Excercise 7 buttons
    Private vacationSave As Boolean = False


    Private Sub btnGameShowRead_Click(sender As Object, e As EventArgs) Handles btnGameShowRead.Click

        'Create a new Streamreader Variable
        Dim inFile As IO.StreamReader

        'Find if File "Exists" if so we will "OpenText" for Input else we will show error message
        If IO.File.Exists("contestants.txt") Then
            'Opening File
            inFile = IO.File.OpenText("contestants.txt")
            'ReadToEnd and Write to the ContestantsResults textbox control
            txtGameShowResults.Text = inFile.ReadToEnd
            'Close the File
            inFile.Close()
        Else
            MessageBox.Show("Cannot find the File.", "Game Show", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End If
    End Sub

    Private Sub btnGameShowWrite_Click(sender As Object, e As EventArgs) Handles btnGameShowWrite.Click

        'Create new StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Open file to input data
        outFile = IO.File.AppendText("contestants.txt")

        'Write each Contestants Name to a New Line
        outFile.WriteLine(txtGameShowName.Text.Trim)

        'Lastly we will close the file
        outFile.Close()

        'Clear Text from Input TextBox and Set focus
        txtGameShowName.Text = ""
        txtGameShowName.Focus()
    End Sub

    Private Sub btnGameShowClear_Click(sender As Object, e As EventArgs) Handles btnGameShowClear.Click

        'Create new StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Open file to input data
        outFile = IO.File.CreateText("contestants.txt")

        'Lastly we will close the file
        outFile.Close()

        'Clear Text from Input TextBox and Set focus
        txtGameShowName.Text = ""
        txtGameShowName.Focus()
    End Sub

    Private Sub btnGame2Read_Click(sender As Object, e As EventArgs) Handles btnGame2Read.Click

        'Create a new StreamReader Variable
        Dim inFile As IO.StreamReader

        'See if File Exists if so the we will open the file to assign to variable and then read line by line else error message

        If IO.File.Exists("GameShowNames.txt") Then
            'Open file to read
            inFile = IO.File.OpenText("GameShowNames.txt")
            'Process Loop Instructions until End of file
            Do Until inFile.Peek = -1
                lstContestants.Items.Add(inFile.ReadLine)
            Loop
            'Close the File
            inFile.Close()
        Else
            MessageBox.Show("Can not find the file")
        End If
    End Sub

    Private Sub btnGame2Write_Click(sender As Object, e As EventArgs) Handles btnGame2Write.Click

        'Create New StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Open File so Application can Write to file
        outFile = IO.File.AppendText("GameShowNames.txt")

        'Write Each Name to a new line in the text file
        outFile.WriteLine(txtGame2Input.Text.Trim)

        'Close the File
        outFile.Close()

        'Clear text and Focus
        txtGame2Input.Text = ""
        txtGame2Input.Focus()
    End Sub

    Private Sub btnGame2Clear_Click(sender As Object, e As EventArgs) Handles btnGame2Clear.Click

        'Create New StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Open File so Application can Write to file
        outFile = IO.File.CreateText("GameShowNames.txt")

        'Clear ListBox
        lstContestants.Items.Clear()

        'Close the File
        outFile.Close()

        'Clear text and Focus
        txtGame2Input.Text = ""
        txtGame2Input.Focus()
    End Sub

    Private Sub btnDoIt1Write_Click(sender As Object, e As EventArgs) Handles btnDoIt1Write.Click

        'Create new StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Open File to Write Data from Application to File
        outFile = IO.File.AppendText("doitone.txt")

        'Write userInput to File
        outFile.WriteLine(txtDoItOneInput.Text)

        'Close the File and Focus text control
        outFile.Close()
        txtDoItOneInput.Text = ""
        txtDoItOneInput.Focus()
    End Sub

    Private Sub btnDoIt1ReadToEnd_Click(sender As Object, e As EventArgs) Handles btnDoIt1ReadToEnd.Click

        'Create New StreamReader Variable
        Dim inFile As IO.StreamReader

        'Find if file exists if so open for data reading, and read to the end
        If IO.File.Exists("doitone.txt") Then
            inFile = IO.File.OpenText("doitone.txt")
            'Write data to text control
            txtDoItOneResults.Text = inFile.ReadToEnd
            'Close File
            inFile.Close()
        Else
            MessageBox.Show("Can Not Find File")
        End If
    End Sub

    Private Sub btnDoIt1ReadLine_Click(sender As Object, e As EventArgs) Handles btnDoIt1ReadLine.Click

        'Create a new StreamReader variable
        Dim infile As IO.StreamReader

        'Find if File Exists, if so Open the File and the Peek to see if there is data if so readLine and add to listBox
        If IO.File.Exists("doitone.txt") Then
            infile = IO.File.OpenText("doitone.txt")
            'Peek at data using a loop
            Do Until infile.Peek = -1
                'Instuctions to ReadLine and add each line to a listbox
                lstDoItOne.Items.Add(infile.ReadLine)
            Loop
            'close file
            infile.Close()
        Else
            MessageBox.Show("Can not find the file")
        End If

    End Sub

    Private Sub btnDoItOneClear_Click(sender As Object, e As EventArgs) Handles btnDoItOneClear.Click

        'Create new StreamWriter variable
        Dim outfile As IO.StreamWriter

        'Create empty file
        outfile = IO.File.CreateText("doitone.txt")

        'Close File
        outfile.Close()

        'Clear textcontrols and ListBoxes
        txtDoItOneInput.Text = ""
        txtDoItOneResults.Text = ""
        lstDoItOne.Items.Clear()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Call to Procedure to Fill araay with data from text file and add to list box
        FillArrayandListBox()

        'Load states.txt file into list box
        'Decalre StreamReader Object
        Dim inFile As IO.StreamReader

        'See if file Exists
        If IO.File.Exists("states.txt") Then
            'Open file if exists
            inFile = IO.File.OpenText("states.txt")
            'Peek till end of File
            Do Until inFile.Peek = -1
                'ReadLine and Add to ListBox
                For intSub As Integer = 0 To states.GetUpperBound(0)
                    states(intSub) = inFile.ReadLine
                Next intSub
            Loop
            inFile.Close()
        Else
            MessageBox.Show("Can not Find the File")
        End If

        'Sort array before adding to list box
        Array.Sort(states)

        'Add array to states listbox
        For intSub As Integer = 0 To states.GetUpperBound(0)
            lstStates.Items.Add(states(intSub))
        Next intSub

        'Excercise 2 Load worker.txt File into list Box

        'new StreamReader
        Dim workerInFile As IO.StreamReader

        'File Exist
        If IO.File.Exists("workers.txt") Then
            'open file
            workerInFile = IO.File.OpenText("workers.txt")
            'peek at file
            Do Until workerInFile.Peek = -1
                'read file
                lstWorkers.Items.Add(workerInFile.ReadLine)
            Loop
            'close file
            workerInFile.Close()
        Else
            MessageBox.Show("Can not find file")
        End If

        'Excercise 6
        'Add selected Index to SchoolCandy List Box
        lstSchoolCandy.SelectedIndex = 0
        txtSchoolCandyInput.Focus()

        'Read schoolcandy.txt file and output information to Labels

        'Declare new StreamReader
        Dim inSchoolFile As IO.StreamReader

        'Find File
        If IO.File.Exists("schoolcandy.txt") Then
            'Open file
            inSchoolFile = IO.File.OpenText("schoolcandy.txt")
            'Peek at File
            Do Until inSchoolFile.Peek = -1
                For intSub As Integer = 0 To candy.GetUpperBound(0)
                    Integer.TryParse(inSchoolFile.ReadLine, candy(intSub))
                Next intSub
            Loop
            'close file
            inSchoolFile.Close()
        End If

        txtCandy1.Text = candy(0).ToString
        txtCandy2.Text = candy(1).ToString
        txtCandy3.Text = candy(2).ToString
        txtCandy4.Text = candy(3).ToString
        txtCandy5.Text = candy(4).ToString

        'EXCERCISE 7

        'Read Data from "needtovisit.txt" file and "visited.txt" file and add to listbox

        'Declare new streamreaders
        Dim inNeedVisitFile As IO.StreamReader
        Dim inVisitedFile As IO.StreamReader

        'find if "needtovisit.txt" file exists
        If IO.File.Exists("needtovisit.txt") Then
            'open file
            inNeedVisitFile = IO.File.OpenText("needtovisit.txt")
            'peek at file
            Do Until inNeedVisitFile.Peek = -1
                lstNeedVisit.Items.Add(inNeedVisitFile.ReadLine)
            Loop
            'close file
            inNeedVisitFile.Close()
        Else
            MessageBox.Show("Can not find File")
        End If

        'find if "visited.txt" file exists
        If IO.File.Exists("visited.txt") Then
            'open file
            inVisitedFile = IO.File.OpenText("visited.txt")
            'peek at file
            Do Until inVisitedFile.Peek = -1
                lstVisited.Items.Add(inVisitedFile.ReadLine)
            Loop
            'close file
            inVisitedFile.Close()
        Else
            MessageBox.Show("Can not find File")
        End If

        'Vacation listboxes selected idexes
        If lstNeedVisit.Items.Count > 0 Then
            lstNeedVisit.SelectedIndex = 0
        End If
        If lstVisited.Items.Count > 0 Then
            lstVisited.SelectedIndex = 0
        End If

        'EXCERCISE 9

        'Read Data from "needtovisit.txt" file and "visited.txt" file and add to listbox

        'Declare new streamreaders
        Dim inInvitedFile As IO.StreamReader
        Dim inAcceptedFile As IO.StreamReader
        Dim inRejectedFile As IO.StreamReader

        'find if "invited.txt" file exists
        If IO.File.Exists("invited.txt") Then
            'open file
            inInvitedFile = IO.File.OpenText("invited.txt")
            'peek at file
            Do Until inInvitedFile.Peek = -1
                lstInvited.Items.Add(inInvitedFile.ReadLine)
            Loop
            'close file
            inInvitedFile.Close()
        Else
            MessageBox.Show("Can not find File")
        End If

        'find if "accepted.txt" file exists
        If IO.File.Exists("accepted.txt") Then
            'open file
            inAcceptedFile = IO.File.OpenText("accepted.txt")
            'peek at file
            Do Until inAcceptedFile.Peek = -1
                lstAccepted.Items.Add(inAcceptedFile.ReadLine)
            Loop
            'close file
            inAcceptedFile.Close()
        Else
            MessageBox.Show("Can not find File")
        End If

        'find if "rejected.txt" file exists
        If IO.File.Exists("rejected.txt") Then
            'open file
            inRejectedFile = IO.File.OpenText("rejected.txt")
            'peek at file
            Do Until inRejectedFile.Peek = -1
                lstRejected.Items.Add(inRejectedFile.ReadLine)
            Loop
            'close file
            inRejectedFile.Close()
        Else
            MessageBox.Show("Can not find File")
        End If

        'selected index for excercise 9
        If lstInvited.Items.Count > 0 Then
            lstInvited.SelectedIndex = 0
        End If
    End Sub
    Private Sub FillArrayandListBox()
        'Create an InFile and Read Data from Continents.txt file which we manually created by clicking File and then New, then
        'click File. Slect the text file option and save to the Projects bin/Debug folder

        'Create new StreamReader Variable
        Dim inFile As IO.StreamReader

        'See if File Exists and then open the file and peek until end of file and read line by line adding to array,
        'Use the addtoListbox procedure that we created to save code time
        If IO.File.Exists("continents.txt") Then
            'Open File
            inFile = IO.File.OpenText("continents.txt")
            'Assign each line from text file to each memberVariable of the array continents
            For intSub As Integer = 0 To lastSub
                continents(intSub) = inFile.ReadLine
            Next intSub
            'Close File
            inFile.Close()
            'Add to Listbox
            AddtoListBox()
        Else
            MessageBox.Show("Can not find the file")
        End If
    End Sub
    Private Sub AddtoListBox()

        'Procedure to Add Vaues from Array to the List Box
        lstContinents.Items.Clear()
        For intSub As Integer = 0 To lastSub
            lstContinents.Items.Add(continents(intSub))
        Next intSub
        lstContinents.SelectedIndex = 0
    End Sub
    Private Sub AscendingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AscendingToolStripMenuItem.Click

        'Sort Items in ListBox by Aplhabetical Order
        Array.Sort(continents)
        AddtoListBox()
    End Sub

    Private Sub DesecendingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DesecendingToolStripMenuItem.Click

        'Sort Items in ListBox by Desecending Aplhabetical Order 
        Array.Sort(continents)
        Array.Reverse(continents)
        AddtoListBox()
    End Sub

    Private Sub OriginalOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OriginalOrderToolStripMenuItem.Click

        'Call to Procedure to Fill araay with data from text file and add to list box
        FillArrayandListBox()
    End Sub

    Private Sub btnHarkins_Click(sender As Object, e As EventArgs) Handles btnHarkins.Click

        'In this app we will create a text file manually using data from the book page 412. We will then accumulate the values 
        'stored in the text file to display the total sales for the year. The text file name will be sales.txt

        'Variables
        Dim inFile As IO.StreamReader
        Dim sales As Integer
        Dim annualSales As Integer

        'Procedure
        If IO.File.Exists("sales.txt") Then
            'Open File
            inFile = IO.File.OpenText("sales.txt")
            'peek Until End of Document
            Do Until inFile.Peek = -1
                'assign number to sales variable
                Integer.TryParse(inFile.ReadLine, sales)
                'accumulate sales
                annualSales += sales
            Loop
            inFile.Close()
            txtHarkinsResults.Text = annualSales.ToString("C2")
        Else
            MessageBox.Show("Can not find the File")
            txtHarkinsResults.Text = "N/A"
        End If
    End Sub
    Private Sub btnStateWrite_Click(sender As Object, e As EventArgs) Handles btnStateWrite.Click

        'Declare new StreamWriter Variable
        Dim outFile As IO.StreamWriter

        'Create a New Text named Sales.txt
        outFile = IO.File.CreateText("states.txt")

        'Write data from Sorted ListBox to New Text File
        For intIndex As Integer = 0 To lstStates.Items.Count - 1
            outFile.WriteLine(lstStates.Items(intIndex))
        Next intIndex

        'Close File
        outFile.Close()
    End Sub

    Private Sub btnStateRead_Click(sender As Object, e As EventArgs) Handles btnStateRead.Click

        lstStates.Items.Clear()
        'Decalre StreamReader Object
        Dim inFile As IO.StreamReader

        'See if file Exists
        If IO.File.Exists("states.txt") Then
            'Open file if exists
            inFile = IO.File.OpenText("states.txt")
            'Peek till end of File
            Do Until inFile.Peek = -1
                'ReadLine and Add to ListBox
                For intSub As Integer = 0 To states.GetUpperBound(0)
                    states(intSub) = inFile.ReadLine
                Next intSub
            Loop
            inFile.Close()
        Else
            MessageBox.Show("Can not Find the File")
        End If

        'Reverse Array then add to listbox
        Array.Reverse(states)

        'Add array to listbox
        For intSub As Integer = 0 To states.GetUpperBound(0)
            lstStates.Items.Add(states(intSub))
        Next intSub
    End Sub

    Private Sub btnExcercise1ElectricBillApp_Click(sender As Object, e As EventArgs) Handles btnExcercise1ElectricBillApp.Click

        'Create an app that will read the manually created text file "electricbill.txt". Display the average electric bill by
        'totaling the years electric bills and dividing by 12. Display results to user

        'Declare a new StreamReader Object
        Dim inFile As IO.StreamReader

        'Variables
        Dim monthlyBill As Double
        Dim yearlyBill As Double
        Dim averageBill As Double

        'Find if file exists
        If IO.File.Exists("electricbill.txt") Then
            'Open file
            inFile = IO.File.OpenText("electricbill.txt")
            'Peek until end of File
            Do Until inFile.Peek = -1
                'Tryparse data
                Double.TryParse(inFile.ReadLine, monthlyBill)
                'accumulate the yearly total
                yearlyBill += monthlyBill
            Loop
            'Close file
            inFile.Close()
        Else
            'error message if file can not be found
            MessageBox.Show("Can not find the file")
        End If

        'find average
        averageBill = yearlyBill / 12

        'display results
        txtElectricBillResults.Text = averageBill.ToString("C2")
    End Sub

    Private Sub btnExcercise2Workers_Click(sender As Object, e As EventArgs) Handles btnExcercise2Workers.Click

        'Create an App that will take in a User Name and add it to the Listbox. When the user closes the Application
        'The contents of the ListBox should be written to a SQL file named "workers.txt" When the Application loads
        'then names should be loaded into the listbox

        'Add userinput to Listbox
        lstWorkers.Items.Add(txtWorkersInput.Text)
        txtWorkersInput.Text = ""

    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        'When the Application Close have the Contents from the lsitbox in Excercise 2 ( lstWorkers ) added to a SQL file named
        '"workers.txt".

        'Declare new StreamWriter Object
        Dim workerFile As IO.StreamWriter
        Dim counter As Integer
        Dim listcounter As Integer = lstWorkers.Items.Count

        Dim counterReader As IO.StreamReader
        If IO.File.Exists("counter.txt") Then
            counterReader = IO.File.OpenText("counter.txt")
            Integer.TryParse(counterReader.ReadToEnd, counter)
            counterReader.Close()
        Else
            MessageBox.Show("file does not exist")
        End If

        'set variable to loop through the listbox count will be 1 more than index so ( -1 ) then write each item to SQL file

        If listcounter = 0 Then
            workerFile = IO.File.CreateText("workers.txt")
        Else
            workerFile = IO.File.AppendText("workers.txt")
            For intIndex As Integer = counter To lstWorkers.Items.Count - 1
                workerFile.WriteLine(lstWorkers.Items(intIndex))
            Next
        End If

        'Close file
        workerFile.Close()

        Dim counterWriter As IO.StreamWriter
        counterWriter = IO.File.CreateText("counter.txt")
        counterWriter.Write(lstWorkers.Items.Count)
        counterWriter.Close()

        'Excercise 6 Application. Will Write to the schoolcandy.txt file the amounts from the textbox controls in Excercise 6 App.

        'Declare new StreamWriter object
        Dim outSchoolFile As IO.StreamWriter

        'Create new Textfile
        outSchoolFile = IO.File.CreateText("schoolcandy.txt")

        'Write Data to Textfile
        outSchoolFile.WriteLine(txtCandy1.Text)
        outSchoolFile.WriteLine(txtCandy2.Text)
        outSchoolFile.WriteLine(txtCandy3.Text)
        outSchoolFile.WriteLine(txtCandy4.Text)
        outSchoolFile.WriteLine(txtCandy5.Text)

        'close file
        outSchoolFile.Close()

        'EXCERCISE 7
        'When application closes if changed were made to the listboxes write information from needvisit listbox to text file named "needtovisit.txt" and write the information
        'from visited listbox to a text file name "visited.txt"

        'Check if user made changes to data in app
        If vacationSave = True Then
            'Ask user if they want to save information
            Dim response As Integer
            response = MessageBox.Show("Do you wish to Save?", "Save Information?",
                   MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If response = vbYes Then
                'declare new stream writer one for needtovisit text file and one for visited text file
                Dim outNeedVisitFile As IO.StreamWriter
                Dim outVisitedFile As IO.StreamWriter
                'create new files
                outNeedVisitFile = IO.File.CreateText("needtovisit.txt")
                outVisitedFile = IO.File.CreateText("visited.txt")
                'write information to files
                For intSub As Integer = 0 To lstNeedVisit.Items.Count - 1
                    outNeedVisitFile.WriteLine(lstNeedVisit.Items(intSub).ToString)
                Next intSub
                For intSub As Integer = 0 To lstVisited.Items.Count - 1
                    outVisitedFile.WriteLine(lstVisited.Items(intSub).ToString)
                Next intSub
                outNeedVisitFile.Close()
                outVisitedFile.Close()
            End If
        Else
            Application.ExitThread()
        End If

        'EXCERCISE 9
        'Save information from wedding application listboxes to the appropriate text files "invited.txt", "accepted.txt" and "rejected.txt"

        'declare new stream writers for wedding text files
        Dim outInvitedFile As IO.StreamWriter
        Dim outAcceptedFile As IO.StreamWriter
        Dim outRejectedFile As IO.StreamWriter

        'create new files
        outInvitedFile = IO.File.CreateText("invited.txt")
        outAcceptedFile = IO.File.CreateText("accepted.txt")
        outRejectedFile = IO.File.CreateText("rejected.txt")
        'write information to files
        For intSub As Integer = 0 To lstInvited.Items.Count - 1
            outInvitedFile.WriteLine(lstInvited.Items(intSub).ToString)
        Next intSub
        For intSub As Integer = 0 To lstAccepted.Items.Count - 1
            outAcceptedFile.WriteLine(lstAccepted.Items(intSub).ToString)
        Next intSub
        For intSub As Integer = 0 To lstRejected.Items.Count - 1
            outRejectedFile.WriteLine(lstRejected.Items(intSub).ToString)
        Next intSub

        'close files
        outInvitedFile.Close()
        outAcceptedFile.Close()
        outRejectedFile.Close()
    End Sub
    Private Sub btnExcercise2WorkersClear_Click(sender As Object, e As EventArgs) Handles btnExcercise2WorkersClear.Click

        'Clear list box
        lstWorkers.Items.Clear()
    End Sub

    Private Sub btnCustomerInformationExcercise3_Click(sender As Object, e As EventArgs) Handles btnCustomerInformationExcercise3.Click

        'Create an Application that will take in User INput about a customer and save it to a text file. The save button should save the customer information to a Sequential Access File named "customers.txt"
        'Make sure all textboxes have data before saving display an error message if and textbox is empty. Use the following format when saving the data.
        'First and Last name should be on the same line seperated by a space. On the next line save the customer's ADDRESS followed by a comma space CITY NAME comma space then lastly the ZIP a Blank line should
        'seperate each customer's information

        'Declare new StreamWriter variable
        Dim outFile As IO.StreamWriter

        'Append Text
        outFile = IO.File.AppendText("customers.txt")

        'Write input to file
        outFile.WriteLine(txtCust1stNameExcer3.Text.Trim & " " & txtCustLastNameExcer3.Text & vbCrLf & txtCustAddressExcer3.Text & ", " & txtCustCityExcer3.Text & ", " & txtCustZipExcer3.Text & vbCrLf)

        'Close File
        outFile.Close()

        lblSavedCustomer.Text = "Customer Information Saved"
    End Sub

    Private Sub btnStudentApplicationExcercise4_Click(sender As Object, e As EventArgs) Handles btnStudentApplicationExcercise4.Click

        'Create an App that will read the studentgrades.txt file and if a student has the grade selected  from the listbox "grades" then display the names of the students in the listbox "students"
        'and how many students got that grade in the number of students textbox.

        'Declare new StreamReader and Variables to store data
        Dim inFile As IO.StreamReader
        Dim studentName As String
        Dim studentGrade As String
        Dim studentCounter As Integer = 0

        'See if Student file exists
        If IO.File.Exists("studentgrades.txt") Then
            'Open file
            inFile = IO.File.OpenText("studentgrades.txt")
            'peek at text
            Do Until inFile.Peek = -1
                'Read lines and store variables
                studentName = inFile.ReadLine
                studentGrade = inFile.ReadLine
                If studentGrade = lstGrades.SelectedItem.ToString Then
                    lstStudentNames.Items.Add(studentName)
                    studentCounter += 1
                End If
            Loop
            'close file
            inFile.Close()
        Else
            MessageBox.Show("Can not find file")
        End If
        txtNumberofStudents.Text = studentCounter.ToString
    End Sub

    Private Sub lstGrades_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstGrades.SelectedIndexChanged
        lstStudentNames.Items.Clear()
    End Sub

    Private Sub btnCookieSalesExcercise5_Click(sender As Object, e As EventArgs) Handles btnCookieSalesExcercise5.Click

        'Create an App that will read the cookiesales.txt file and accumlate the amount of cookies that were sold by each type. Display the totals for each cookie in the appropriate textbox control

        'Declare new StreamReader and Variables to Store Information Read
        Dim inFile As IO.StreamReader
        Dim cookieType As String
        Dim cookieSold As Integer
        Dim totalCookies(3) As Integer

        'See if File Exists
        If IO.File.Exists("cookiesales.txt") Then
            'Open File
            inFile = IO.File.OpenText("cookiesales.txt")
            'peek at File
            Do Until inFile.Peek = -1
                'read file
                cookieType = inFile.ReadLine.Trim
                Integer.TryParse(inFile.ReadLine, cookieSold)
                If cookieType = "Chocolate Chip" Then
                    totalCookies(0) += cookieSold
                    txtChoco.Text = totalCookies(0).ToString
                ElseIf cookieType = "Peanut Butter" Then
                    totalCookies(1) += cookieSold
                    txtPB.Text = totalCookies(1).ToString
                ElseIf cookieType = "Oatmeal" Then
                    totalCookies(2) += cookieSold
                    txtOat.Text = totalCookies(2).ToString
                ElseIf cookieType = "Sugar" Then
                    totalCookies(3) += cookieSold
                    txtSugar.Text = totalCookies(3).ToString
                End If
            Loop
            'close file
            inFile.Close()
        Else
            MessageBox.Show("Can not find file")
        End If
    End Sub

    Private Sub btnAddSchoolCandyExcercise6_Click(sender As Object, e As EventArgs) Handles btnAddSchoolCandyExcercise6.Click

        'Create an Application that will allow user to select a candy type from the listbox. The user will then add a sold amount to the textbox and and click add to total based on the selected index 
        'add to the correlating position in the candy array and add the total

        'Variables and Tryparse

        Dim userNumber As Integer
        Dim index As Integer = lstSchoolCandy.SelectedIndex
        Integer.TryParse(txtSchoolCandyInput.Text, userNumber)

        'Add user input to candy array based on selected index from listbox schoolcandy
        candy(index) += userNumber

        'Display totals to User
        txtCandy1.Text = candy(0).ToString
        txtCandy2.Text = candy(1).ToString
        txtCandy3.Text = candy(2).ToString
        txtCandy4.Text = candy(3).ToString
        txtCandy5.Text = candy(4).ToString
    End Sub

    Private Sub btnVacationAppExcercise7_Click(sender As Object, e As EventArgs) Handles btnVacationAppExcercise7.Click

        'Create an app that will display Vacation Destenations to a user. Allow the user to input a new destenation to need to visit listbox by clickig add location 
        'the move to visit button will move the slected item in need to visit listbox and move it to the visited listbox. The undo buttun will take the selected item from
        'visited listbox and move it back to the need to visit list box. When the app closes have a messagebox appear and ask the user if the want to save the data from the application
        'the information from the needtovisit listbox wll write to a text file name needtovisit.txt and the visited listbox will write to visited.txt.

        'when user clicks add information from textbox to the Listbox NeedVisit
        lstNeedVisit.Items.Add(txtVacationInput.Text)
        vacationSave = True

    End Sub

    Private Sub btnVacationMoveLocation_Click(sender As Object, e As EventArgs) Handles btnVacationMoveLocation.Click

        'when the user clicks move the selectedItem from needvisit listbox to the visited listbox
        If lstNeedVisit.Items.Count = 0 Then
            lstNeedVisit.Items.Clear()
        Else
            Dim location As String = lstNeedVisit.SelectedItem.ToString
            lstNeedVisit.Items.RemoveAt(lstNeedVisit.SelectedIndex)
            lstVisited.Items.Add(location)
            If lstNeedVisit.Items.Count > 0 Then
                lstNeedVisit.SelectedIndex = 0
            End If
            If lstVisited.Items.Count > 0 Then
                lstVisited.SelectedIndex = 0
            End If
        End If
        vacationSave = True
    End Sub

    Private Sub btnVacationUndo_Click(sender As Object, e As EventArgs) Handles btnVacationUndo.Click

        'when the user clicks move the selectedItem from visited listbox to the needvisit listbox
        If lstVisited.Items.Count = 0 Then
            lstNeedVisit.Items.Clear()
        ElseIf lstVisited.Items.Count > 0 Then
            Dim location As String = lstVisited.SelectedItem.ToString
            lstVisited.Items.RemoveAt(lstVisited.SelectedIndex)
            lstNeedVisit.Items.Add(location)
            If lstVisited.Items.Count > 0 Then
                lstVisited.SelectedIndex = 0
            End If
            If lstNeedVisit.Items.Count > 0 Then
                lstNeedVisit.SelectedIndex = 0
            End If
        End If
        vacationSave = True
    End Sub

    Private Sub btnWeddingInputAdd_Click(sender As Object, e As EventArgs) Handles btnWeddingInputAdd.Click

        'Create an app that will display the wedding list text files in their respective listboxes. the text files are "invited.txt", "accepted.txt" and "rejected.txt" when the user clicks add person
        'take user input and add it to the invited list box.

        'adding user input to list box
        lstInvited.Items.Add(txtWeddingInput.Text)
    End Sub

    Private Sub btnMovetoAccepted_Click(sender As Object, e As EventArgs) Handles btnMovetoAccepted.Click

        'when the user clicks movetoaccepted the selectedItem from either invited or rejected listbox will be removed and then added to the accepted listbox 
        If lstInvited.Items.Count = 0 Then
            lstInvited.Items.Clear()
        ElseIf lstInvited.SelectedIndex >= 0 Then
            Dim person As String = lstInvited.SelectedItem.ToString
            lstInvited.Items.RemoveAt(lstInvited.SelectedIndex)
            lstAccepted.Items.Add(person)
            If lstInvited.Items.Count > 0 Then
                lstInvited.SelectedIndex = 0
            End If
        End If
        If lstRejected.Items.Count = 0 Then
            lstRejected.Items.Clear()
        ElseIf lstRejected.SelectedIndex >= 0 Then
            Dim person As String = lstRejected.SelectedItem.ToString
            lstRejected.Items.RemoveAt(lstRejected.SelectedIndex)
            lstAccepted.Items.Add(person)
            If lstRejected.Items.Count > 0 Then
                lstRejected.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btnMovetoInvited_Click(sender As Object, e As EventArgs) Handles btnMovetoInvited.Click

        'when the user clicks movetoinvited the selectedItem from either accepted or rejected listbox will be removed and then added to the Invited listbox 
        If lstAccepted.Items.Count = 0 Then
            lstAccepted.Items.Clear()
        ElseIf lstAccepted.SelectedIndex >= 0 Then
            Dim person As String = lstAccepted.SelectedItem.ToString
            lstAccepted.Items.RemoveAt(lstAccepted.SelectedIndex)
            lstInvited.Items.Add(person)
            If lstAccepted.Items.Count > 0 Then
                lstAccepted.SelectedIndex = 0
            End If
        End If
        If lstRejected.Items.Count = 0 Then
            lstRejected.Items.Clear()
        ElseIf lstRejected.SelectedIndex >= 0 Then
            Dim person As String = lstRejected.SelectedItem.ToString
            lstRejected.Items.RemoveAt(lstRejected.SelectedIndex)
            lstInvited.Items.Add(person)
            If lstRejected.Items.Count > 0 Then
                lstRejected.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btnMoveToRejected_Click(sender As Object, e As EventArgs) Handles btnMoveToRejected.Click

        'when the user clicks movetorejected the selectedItem from either accepted or invited listbox will be removed and then added to the rejected listbox 
        If lstInvited.Items.Count = 0 Then
            lstInvited.Items.Clear()
        ElseIf lstInvited.SelectedIndex >= 0 Then
            Dim person As String = lstInvited.SelectedItem.ToString
            lstInvited.Items.RemoveAt(lstInvited.SelectedIndex)
            lstRejected.Items.Add(person)
            If lstInvited.Items.Count > 0 Then
                lstInvited.SelectedIndex = 0
            End If
        End If
        If lstAccepted.Items.Count = 0 Then
            lstAccepted.Items.Clear()
        ElseIf lstAccepted.SelectedIndex >= 0 Then
            Dim person As String = lstAccepted.SelectedItem.ToString
            lstAccepted.Items.RemoveAt(lstAccepted.SelectedIndex)
            lstRejected.Items.Add(person)
            If lstAccepted.Items.Count > 0 Then
                lstAccepted.SelectedIndex = 0
            End If
        End If
    End Sub
    Private Sub lstInvitedSelectedIndexChanged(sender As Object, e As EventArgs) Handles lstInvited.SelectedIndexChanged
        'clear selected indexes if greater than 1 gets selected
        If lstInvited.SelectedIndex > -1 Then
            lstAccepted.SelectedIndex = -1
            lstRejected.SelectedIndex = -1
        End If
    End Sub
    Private Sub lstAccepted_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstAccepted.SelectedIndexChanged
        'clear selected indexes if greater than 1 gets selected
        If lstAccepted.SelectedIndex > -1 Then
            lstInvited.SelectedIndex = -1
            lstRejected.SelectedIndex = -1
        End If
    End Sub
    Private Sub lstRejected_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRejected.SelectedIndexChanged
        'clear selected indexes if greater than 1 gets selected
        If lstRejected.SelectedIndex > -1 Then
            lstInvited.SelectedIndex = -1
            lstAccepted.SelectedIndex = -1
        End If
    End Sub
End Class
