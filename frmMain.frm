VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Process Prioritiser"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin ComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4605
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "processCount"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "threadCount"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5477
            Text            =   "Process Prioritiser v1.0"
            TextSave        =   "Process Prioritiser v1.0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Operation"
      TabPicture(0)   =   "frmMain.frx":61CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrMain"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdHide"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Real-Time"
      TabPicture(1)   =   "frmMain.frx":61E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lstRealTime"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdRealTimeAdd"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRealTimeDelete"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRealTimeClear"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "High"
      TabPicture(2)   =   "frmMain.frx":6202
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdHighClear"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdHighDelete"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdHighAdd"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lstHigh"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Normal"
      TabPicture(3)   =   "frmMain.frx":621E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdNormalClear"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdNormalDelete"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdNormalAdd"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lstNormal"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Image2"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Idle"
      TabPicture(4)   =   "frmMain.frx":623A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdIdleClear"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdIdleDelete"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdIdleAdd"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lstIdle"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Current Processes"
      TabPicture(5)   =   "frmMain.frx":6256
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdRefreshTasks"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lstTasks"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Image4"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblRefreshInfo"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label4"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "About"
      TabPicture(6)   =   "frmMain.frx":6272
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label5"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "picLogo"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   1560
         Picture         =   "frmMain.frx":628E
         ScaleHeight     =   795
         ScaleWidth      =   2265
         TabIndex        =   32
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Minimise to Tray"
         Height          =   375
         Left            =   -74880
         TabIndex        =   31
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefreshTasks 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -70320
         TabIndex        =   29
         Top             =   3840
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Operating Mode"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   5535
         Begin VB.OptionButton optScan 
            Caption         =   "System Process Scan"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optHook 
            Caption         =   "Shell Hook"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin ComctlLib.Slider sldInterval 
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   327682
            Min             =   5
            Max             =   60
            SelStart        =   5
            TickFrequency   =   5
            Value           =   5
         End
         Begin VB.Label lblInterval 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   3360
            TabIndex        =   26
            Top             =   960
            Width           =   45
         End
         Begin VB.Label lblRefresh 
            Caption         =   "Refresh interval:"
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":852E
            Top             =   1920
            Width           =   480
         End
         Begin VB.Label lblInfo 
            Height          =   855
            Left            =   720
            TabIndex        =   24
            Top             =   1920
            Width           =   4695
         End
      End
      Begin VB.CommandButton cmdRealTimeClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   -70320
         TabIndex        =   19
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdRealTimeDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -71280
         TabIndex        =   18
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdRealTimeAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdHighClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   -70320
         TabIndex        =   15
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdHighDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -71280
         TabIndex        =   14
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdHighAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdNormalClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   -70320
         TabIndex        =   12
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdNormalDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -71280
         TabIndex        =   11
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdNormalAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdIdleClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   -70320
         TabIndex        =   9
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdIdleDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -71280
         TabIndex        =   8
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdIdleAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -74880
         TabIndex        =   7
         Top             =   3960
         Width           =   975
      End
      Begin VB.ListBox lstRealTime 
         Height          =   2400
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   5535
      End
      Begin VB.ListBox lstHigh 
         Height          =   3375
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   5535
      End
      Begin VB.ListBox lstNormal 
         Height          =   2790
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   5535
      End
      Begin VB.ListBox lstIdle 
         Height          =   3375
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   5535
      End
      Begin VB.Timer tmrMain 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -69720
         Top             =   3960
      End
      Begin ComctlLib.ListView lstTasks 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Process"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Priority"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   $"frmMain.frx":8E4E
         Height          =   855
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   $"frmMain.frx":8F45
         Height          =   975
         Left            =   240
         TabIndex        =   33
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMain.frx":9099
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblRefreshInfo 
         Height          =   495
         Left            =   -74280
         TabIndex        =   30
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Current Process List:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Most processes are launched at normal priority by default. There is no need to manually set the priority of such processes."
         Height          =   495
         Left            =   -74280
         TabIndex        =   6
         Top             =   3360
         Width           =   4935
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMain.frx":99B9
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":A2D9
         Height          =   795
         Left            =   -74280
         TabIndex        =   1
         Top             =   3000
         Width           =   4905
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMain.frx":A3BB
         Top             =   3000
         Width           =   480
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IconObject As Object

Private Sub AddItem(list As ListBox)

Dim processName As String

frmAdd.Show vbModal
processName = frmAdd.txtName.Text
frmAdd.txtName.Text = ""
If Len(processName) Then
    list.AddItem processName
End If

End Sub

Private Sub cmdHide_Click()

Me.Visible = False

End Sub

Private Sub cmdHighAdd_Click()

AddItem lstHigh

End Sub

Private Sub cmdHighClear_Click()

ClearList lstHigh

End Sub

Private Sub cmdHighDelete_Click()

DeleteItem lstHigh

End Sub

Private Sub cmdIdleAdd_Click()

AddItem lstIdle

End Sub

Private Sub cmdIdleClear_Click()

ClearList lstIdle

End Sub

Private Sub ClearList(list As ListBox)

Dim retCode As Integer

retCode = MsgBox("Are you sure you wish to clear this process list?", vbQuestion + vbYesNo, "Process Prioritiser")
If retCode = 6 Then list.Clear

End Sub

Private Sub DeleteItem(list As ListBox)

If list.ListIndex <> -1 Then list.RemoveItem list.ListIndex

End Sub

Private Sub cmdIdleDelete_Click()

DeleteItem lstIdle

End Sub

Private Sub cmdNormalAdd_Click()

AddItem lstNormal

End Sub

Private Sub cmdNormalClear_Click()

ClearList lstNormal

End Sub

Private Sub cmdNormalDelete_Click()

DeleteItem lstNormal

End Sub

Private Sub cmdRealTimeAdd_Click()

AddItem lstRealTime

End Sub

Private Sub cmdRealTimeClear_Click()

ClearList lstRealTime

End Sub

Private Sub cmdRealTimeDelete_Click()

DeleteItem lstRealTime

End Sub

Private Sub cmdRefreshTasks_Click()

RefreshTasks

End Sub

Private Sub Initialize()

Set IconObject = Me.Icon
AddIcon Me, IconObject.Handle, IconObject, "Process Prioritiser"

End Sub

Private Sub Form_Load()
    
Initialize
RefreshTasks
tmrMain.Enabled = True
sldInterval.Value = 5
lblInterval.Caption = Str$(sldInterval.Value) + " seconds"
lblInfo.Caption = "This mode of operation scans processes every x seconds and changes any " & _
            "instance of any process listed for priority change. This mode is not " & _
            "limited to 'windowed' applications but requires more CPU time."
lblRefreshInfo.Caption = "Process Scan mode. Process list refreshes every" & Str$(sldInterval.Value) & " seconds."

End Sub

Private Sub Form_Unload(Cancel As Integer)

If optHook.Value = True Then EndHook
delIcon IconObject.Handle
delIcon Me.Icon.Handle

End Sub

Private Sub mnuPopupExit_Click()

Unload Me

End Sub

Private Sub mnuPopupOpen_Click()

Me.Visible = True

End Sub

Private Sub optHook_Click()

lblInfo.Caption = "This mode of operation uses a hook into the Windows Shell Notification " & _
            "calls and matches the process behind any new window opened to those listed " & _
            "for priority change. This method is limited to 'windowed' applications but " & _
            "requires less CPU time."
lblRefreshInfo.Caption = "System Hook mode. Process list does not refresh automatically."
lblRefresh.Enabled = False
lblInterval.Enabled = False
sldInterval.Enabled = False
tmrMain.Enabled = False

InitiateHook

End Sub

Private Sub optScan_Click()

EndHook

lblInfo.Caption = "This mode of operation scans processes every x seconds and changes any " & _
            "instance of any process listed for priority change. This mode is not " & _
            "limited to 'windowed' applications but requires more CPU time."
lblRefreshInfo.Caption = "Process Scan mode. Process list refreshes every " & Str$(sldInterval.Value) & "seconds."
lblRefresh.Enabled = True
lblInterval.Enabled = True
sldInterval.Enabled = True
tmrMain.Enabled = True

End Sub

Private Sub sldInterval_Scroll()

tmrMain.Interval = sldInterval.Value * 1000
lblInterval.Caption = Str$(sldInterval.Value) + " seconds"


End Sub

Private Sub tmrMain_Timer()

RefreshTasks

End Sub

Sub InitiateHook()

uRegMsg = RegisterWindowMessage(ByVal "SHELLHOOK")
Call RegisterShellHook(hwnd, RSH_REGISTER_TASKMAN)
OldProc = GetWindowLong(hwnd, GWL_WNDPROC)
SetWindowLong hwnd, GWL_WNDPROC, AddressOf WndProc

End Sub

Sub EndHook()

Call RegisterShellHook(hwnd, RSH_DEREGISTER)
SetWindowLong hwnd, GWL_WNDPROC, OldProc

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP:
        Me.PopupMenu mnuPopup
    End Select
End Sub
