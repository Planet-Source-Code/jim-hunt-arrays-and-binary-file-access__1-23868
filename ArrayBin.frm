VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArrayBin 
   Caption         =   "Using Arrays to Access Binary Files"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   Picture         =   "ArrayBin.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddSomeRecords 
      Caption         =   "Add Records"
      Height          =   315
      Left            =   6300
      TabIndex        =   29
      ToolTipText     =   "Add 100 Records to the database"
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   -300
      TabIndex        =   28
      Top             =   0
      Width           =   8490
   End
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   -750
      TabIndex        =   27
      Top             =   5625
      Width           =   8490
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   240
      Left            =   4500
      TabIndex        =   26
      Top             =   5700
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer tmrInfo 
      Interval        =   5000
      Left            =   3825
      Top             =   5625
   End
   Begin VB.CommandButton cmdRecordCount 
      Caption         =   "Record &Count"
      Height          =   315
      Left            =   6300
      TabIndex        =   13
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Items"
      Height          =   3015
      Left            =   525
      TabIndex        =   24
      Top             =   2475
      Width           =   2640
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "&Remove Item"
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   300
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Item"
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   1140
      End
      Begin VB.ListBox lstPut 
         Height          =   2205
         ItemData        =   "ArrayBin.frx":171C
         Left            =   150
         List            =   "ArrayBin.frx":171E
         TabIndex        =   6
         Top             =   675
         Width           =   2340
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings (Stored before records - can be used for program settings, etc.)"
      Height          =   2265
      Left            =   525
      TabIndex        =   19
      Top             =   150
      Width           =   5565
      Begin VB.TextBox txtPut 
         Height          =   315
         Index           =   3
         Left            =   1125
         TabIndex        =   3
         Text            =   "txtPut"
         Top             =   1725
         Width           =   4290
      End
      Begin VB.TextBox txtPut 
         Height          =   315
         Index           =   2
         Left            =   1125
         TabIndex        =   2
         Text            =   "txtPut"
         Top             =   1275
         Width           =   4290
      End
      Begin VB.TextBox txtPut 
         Height          =   315
         Index           =   1
         Left            =   1125
         TabIndex        =   1
         Text            =   "txtPut"
         Top             =   825
         Width           =   4290
      End
      Begin VB.TextBox txtPut 
         Height          =   315
         Index           =   0
         Left            =   1125
         TabIndex        =   0
         Text            =   "txtPut"
         Top             =   375
         Width           =   4290
      End
      Begin VB.Label Label4 
         Caption         =   "Setting 4"
         Height          =   240
         Left            =   225
         TabIndex        =   23
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Setting 3"
         Height          =   240
         Left            =   225
         TabIndex        =   22
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Setting 2"
         Height          =   240
         Left            =   225
         TabIndex        =   21
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Setting 1"
         Height          =   240
         Left            =   225
         TabIndex        =   20
         Top             =   450
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Details (make changes to record below)"
      Height          =   3015
      Left            =   3300
      TabIndex        =   15
      Top             =   2475
      Width           =   4215
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   390
         Index           =   0
         Left            =   975
         TabIndex        =   7
         Top             =   300
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   688
         _Version        =   393217
         TextRTF         =   $"ArrayBin.frx":1720
      End
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   390
         Index           =   1
         Left            =   975
         TabIndex        =   8
         Top             =   750
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   688
         _Version        =   393217
         TextRTF         =   $"ArrayBin.frx":17EF
      End
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   1665
         Index           =   2
         Left            =   975
         TabIndex        =   9
         Top             =   1200
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   2937
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"ArrayBin.frx":18BE
      End
      Begin VB.Label Label8 
         Caption         =   "Data"
         Height          =   240
         Left            =   150
         TabIndex        =   18
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Filename"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   825
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Title"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   375
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   6300
      TabIndex        =   14
      Top             =   2100
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   315
      Left            =   6300
      TabIndex        =   12
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6300
      TabIndex        =   10
      Top             =   225
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   6300
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "txtInfo"
      Height          =   240
      Left            =   75
      TabIndex        =   25
      Top             =   5715
      Width           =   4380
   End
End
Attribute VB_Name = "frmArrayBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  --------------------------------------  '
'  Code by:        Jim Hunt                '
'  E-mail:         jim@huntcs.com          '
'                                          '
'  Enjoy!  If you make any modifications   '
'  or improvements to this code, I would   '
'  appreciate an e-mail with the changes   '
'  - or at least honourable mention in     '
'    your software release.                '
'  - or you can you can send me money :)   '
'  --------------------------------------  '

Option Explicit

'the following type declarations define the records

'Record to store program settings or whatever
Private Type Settings 'Would be declared as Public in module
    Setting1 As String
    Setting2 As String
    Setting3 As String
    Setting4 As String
End Type

'The record used to store TotalRecords in the binary file
Private Type RecordCount
    NumberOfRecords As Long
End Type

'The actual database record
Private Type Records 'Would be declared as Public in module
    Field1 As String
    Field2 As String
    Field3 As String
    Field4 As String
End Type

Dim RecordArray() As Records 'Stores the database records
Dim TotalRecords As Long 'Keeps track of how many records are in the DB
Dim TempArray() As Records 'Used for record deletion

Private Sub cmdAdd_Click()
    Dim Response As String
    Dim Record As Records
    
    Response = InputBox("Enter a title for this record")
    If Response = "" Then Exit Sub
    
    'Add some data to the fields
    Record.Field1 = Response
    Record.Field2 = "Title: " & Response
    Record.Field3 = "Filename: " & Response
    Record.Field4 = "Data: " & Response
        
    If TotalRecords = 0 Then
        ReDim RecordArray(0) As Records 'Simply reinitialize the array to prepare for new data
    Else
        ReDim Preserve RecordArray(TotalRecords) As Records 'Remember, TotalRecords is always 1 more than UBound(RecordArray, 1)
    End If

    TotalRecords = TotalRecords + 1 'Update number of records
    
    'Add the data from above to the array
    RecordArray(UBound(RecordArray)) = Record
    
    RefillListBox
    
    'Select the added item in the listbox
    lstPut.Selected(lstPut.NewIndex) = True
    
End Sub

Private Sub cmdAddSomeRecords_Click()
    Dim Record As Records
    Dim Counter As Long
    Dim Response As String
    
    Dim RecordsToAdd As Long
    
    Response = InputBox("This will add records to the database" & vbCrLf & vbCrLf & "Enter a number")
    
    If Response = "" Then Exit Sub
    If CLng(Response) = 0 Then Exit Sub
    RecordsToAdd = CLng(Response)
    
    Progress.Max = RecordsToAdd
    
    If TotalRecords = 0 Then
        ReDim RecordArray(0) As Records 'Simply reinitialize the array to prepare for new data
    End If

    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    lblInfo.Caption = "Working..."
    lblInfo.Refresh 'Force label to show new caption
    
    For Counter = 1 To RecordsToAdd
    
        'Add some data to the fields
        Record.Field1 = "Item " & TotalRecords + Counter
        Record.Field2 = "Title of " & Record.Field1
        Record.Field3 = "Filename of " & Record.Field1
        Record.Field4 = "Data of " & Record.Field1
        
        ReDim Preserve RecordArray((TotalRecords - 1) + Counter) As Records 'Remember, TotalRecords is always 1 more than UBound(RecordArray, 1)
    
        'Add the data from above to the array
        RecordArray(UBound(RecordArray)) = Record
    
        Progress.Value = Counter
    
    Next
    
    TotalRecords = TotalRecords + RecordsToAdd 'Update number of records
    
    If TotalRecords < 32767 Then
        RefillListBox
        'Select the added item in the listbox
        lstPut.Selected(lstPut.NewIndex) = True
    Else
        MsgBox "There are too many records to display in the listbox!"
    End If
    
    'Display message for 5 seconds
    tmrInfo.Enabled = False
    lblInfo.Caption = RecordsToAdd & " records added successfully... Ready"
    tmrInfo.Enabled = True
    Progress.Value = 0

End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdNew_Click()
    'Clear all text controls and reset array
    Dim Counter As Integer
    For Counter = 0 To 3
        txtPut(Counter).Text = ""
    Next
    lstPut.Clear
    For Counter = 0 To 2
        rtfPut(Counter).Text = ""
    Next
    ReDim RecordArray(0) As Records
    TotalRecords = 0
End Sub

Private Sub cmdOpen_Click()
On Error GoTo ErrHandler
    'Open the included test file (if you mess this one up, rename "binarytest.bak" to "binarytest.dat"
    Dim fname As String
    Dim Counter As Long
    Dim Record As Records
    Dim Setting As Settings
    Dim RecCount As RecordCount
    
    fname = App.Path & "\binarytest.dat"
    
    Open fname For Binary As #1
    Get #1, , Setting 'Retrieve the settings first
    Get #1, , RecCount 'Retrieve number of records
    
    With Progress
        .Min = 0
        .Max = RecCount.NumberOfRecords
    End With
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    lblInfo.Caption = "Retrieving Records..."
    lblInfo.Refresh 'Force label to show new caption
    
    'Loop through the data and add it to the array
    For Counter = 0 To RecCount.NumberOfRecords - 1
        ReDim Preserve RecordArray(Counter) As Records 'Increase the size of the array
        Get #1, , RecordArray(Counter) 'Add record to array
        Progress.Value = Counter
    Next
    Close #1
    
    TotalRecords = RecCount.NumberOfRecords
    
    If Counter < 1 Then
        Exit Sub
    End If
    
    'Populate controls with loaded data
    txtPut(0).Text = Setting.Setting1
    txtPut(1).Text = Setting.Setting2
    txtPut(2).Text = Setting.Setting3
    txtPut(3).Text = Setting.Setting4
    
    RefillListBox
    
    'Display message for 5 seconds
    tmrInfo.Enabled = False
    lblInfo.Caption = TotalRecords & " records retrieved successfully... Ready"
    tmrInfo.Enabled = True
    
    'Select the first item in the listbox
    lstPut.Selected(0) = True
    
    'Reset Progress bar
    Progress.Value = 0
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Open failed!"
End Sub

Private Sub cmdRecordCount_Click()
    'I just used this for testing
    If TotalRecords = 1 Then
        MsgBox "There is only " & TotalRecords & " record."
    Else
        MsgBox "There are " & TotalRecords & " records."
    End If
End Sub

Private Sub cmdRemoveItem_Click()
    'Check if anything selected in listbox
    If lstPut.ListIndex = -1 Then Exit Sub
    
    'Ask if it's okay
    Dim Response As Integer
    Response = MsgBox("You are about to delete the record: " & Chr(34) & RecordArray(lstPut.ListIndex).Field1 & Chr(34) & "." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbExclamation, "WARNING")
    If Response = 7 Then Exit Sub
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    lblInfo.Caption = "Deleting Record..."
    lblInfo.Refresh 'Force label to show new caption
    
    'With verification out of the way let's continue
    Dim Counter As Long
    Dim DeletedFlag As Boolean

    'Prepare TempArray to receive data
    ReDim TempArray(TotalRecords - 1) As Records
        
    'Copy the contents of RecordArray, minus the deleted record
    For Counter = 0 To TotalRecords - 1
        If Counter = lstPut.ListIndex Then
            DeletedFlag = True 'Raise flag
        Else
            If DeletedFlag Then 'Move remaining records down by one to fill the gap
                TempArray(Counter - 1) = RecordArray(Counter)
            Else
                TempArray(Counter) = RecordArray(Counter)
            End If
        End If
    Next

    'Now initialize RecordArray and fill with TempArray values
    TotalRecords = TotalRecords - 1 'Update total records to show deletion
    If TotalRecords > 0 Then
        ReDim RecordArray(TotalRecords - 1)
    Else
        ReDim RecordArray(0)
        TotalRecords = 0
    End If
    
    'Start filling RecordArray
    For Counter = 0 To TotalRecords - 1
        RecordArray(Counter) = TempArray(Counter)
    Next Counter
    
    'Clear Field Boxes and refresh listbox
    For Counter = 0 To 2
        rtfPut(Counter).Text = ""
    Next
    
    'Remove item from listbox
    lstPut.RemoveItem (lstPut.ListIndex)
    
    ' There's no need to call RefillListBox since the listbox keeps track of what we need

    'Display message for 5 seconds
    tmrInfo.Enabled = False
    lblInfo.Caption = "Record deleted successfully... Ready"
    tmrInfo.Enabled = True
    Progress.Value = 0

End Sub

Private Sub cmdSave_Click()
    'Save all your hard work!
    Dim fname As String
    Dim Record As Records
    Dim Setting As Settings
    Dim RecCount As RecordCount
    Dim Counter As Long
    
    fname = App.Path & "\binarytest.dat"
    
    'if the file is there, get rid of it
    'this is a good spot to add some code to backup the original file
    If Dir(fname) <> "" Then
        Kill fname
    End If
    
    'add settings
    Setting.Setting1 = txtPut(0).Text
    Setting.Setting2 = txtPut(1).Text
    Setting.Setting3 = txtPut(2).Text
    Setting.Setting4 = txtPut(3).Text
    
    'Add RecordCount
    RecCount.NumberOfRecords = TotalRecords
    
    'Prepare progress bar
    Progress.Max = TotalRecords - 1
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    lblInfo.Caption = "Saving Records..."
    lblInfo.Refresh 'Force label to show new caption
    
    'Output all data to the file
    Open fname For Binary As #1
    Put #1, , Setting 'Write the program settings first
    Put #1, , RecCount 'Write the record count
    For Counter = 0 To TotalRecords - 1
        Put #1, , RecordArray(Counter)
        Progress.Value = Counter
    Next
    Close #1
    
    Progress.Value = 0
    
    'Display message for 5 seconds
    lblInfo.Caption = TotalRecords & " records saved successfully... Ready"
    tmrInfo.Enabled = True
    
End Sub

Private Sub Form_Load()
    Me.Show
    
    cmdNew_Click 'Clear the form
    
    Dim MSG As String
    MSG = MSG & "This example demonstrates how to use a" & vbCrLf
    MSG = MSG & "dynamic array to load the contents of a" & vbCrLf
    MSG = MSG & "binary file used as a simple flat-file database." & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "This type of data access will also reduce" & vbCrLf
    MSG = MSG & "your project distribution size by several" & vbCrLf
    MSG = MSG & "megabytes, since you won't need to include" & vbCrLf
    MSG = MSG & "DAO/ADO/RDO support files with your app!" & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "If you find this project useful in any way," & vbCrLf
    MSG = MSG & "post your comment good or bad!" & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "You may use this however you see fit." & vbCrLf
    MSG = MSG & "The file " & Chr(34) & "binarytest.dat" & Chr(34) & " will now be opened" & vbCrLf
    MSG = MSG & "Created by Jim Hunt" & vbCrLf
    
    MsgBox MSG, vbInformation, "Arrays and Binary Files"
    
    ReDim RecordArray(0) As Records 'Initialize array
    
    'Open the file "binarytest.dat" and if it's not there, just continue normally
    If Dir(App.Path & "\binarytest.dat") <> "" Then
        cmdOpen_Click
    End If
    
End Sub

Private Function RefillListBox()
    'Clears, then refills using the array data
    Dim RecordNo As Integer
    lstPut.Clear
    For RecordNo = 0 To TotalRecords - 1
        lstPut.AddItem RecordArray(RecordNo).Field1
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Free up the memory used for records
    'Gotta find a better way!
    ReDim RecordArray(0) As Records
End Sub

Private Sub lstPut_Click()
    'Show field contents for selected record
    rtfPut(0).Text = RecordArray(lstPut.ListIndex).Field2
    rtfPut(1).Text = RecordArray(lstPut.ListIndex).Field3
    rtfPut(2).Text = RecordArray(lstPut.ListIndex).Field4
End Sub

Private Sub rtfPut_LostFocus(Index As Integer)
    'Whenever there is a change, update the array when the user moves off the control
     If lstPut.ListCount = 0 Then Exit Sub
    
    'update the array when focus changes
    Select Case Index
        Case 0
            RecordArray(lstPut.ListIndex).Field2 = rtfPut(0).Text
        Case 1
            RecordArray(lstPut.ListIndex).Field3 = rtfPut(1).Text
        Case 2
            RecordArray(lstPut.ListIndex).Field4 = rtfPut(2).Text
    End Select
End Sub

Private Sub tmrInfo_Timer()
    ' This timer resets lblInfo to Ready
    lblInfo.Caption = "Ready"
    tmrInfo.Enabled = False
End Sub
