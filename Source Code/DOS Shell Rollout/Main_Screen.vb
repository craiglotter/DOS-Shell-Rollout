Imports Microsoft.Win32
Imports System.IO
Imports System.Data.Sqlclient

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerErrorEncountered(ByVal ex As Exception, ByVal descriptor As String)
    Public Delegate Sub WorkerhStatusMessageUpdate(ByVal message As String)

    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private testingconnection As Boolean = False
    Private error_reporting_level

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen, Optional ByVal currentuser As String = "Unknown User")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
        AddHandler Worker1.WorkerStatusMessageUpdate, AddressOf WorkerStatusMessageUpdate
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents StatusBar As System.Windows.Forms.Label
    Friend WithEvents ServerStatus As System.Windows.Forms.Label
    Friend WithEvents ConnectionTesterCounter As System.Windows.Forms.Label
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ConnectionTester As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.ServerStatus = New System.Windows.Forms.Label
        Me.ConnectionTesterCounter = New System.Windows.Forms.Label
        Me.StatusBar = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ConnectionTester = New System.Windows.Forms.Timer(Me.components)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(240, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(368, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 16)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "BUILD 20060113.3"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label9, "Application Build Version")
        '
        'ServerStatus
        '
        Me.ServerStatus.BackColor = System.Drawing.Color.Transparent
        Me.ServerStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ServerStatus.ForeColor = System.Drawing.Color.Orange
        Me.ServerStatus.Location = New System.Drawing.Point(160, 160)
        Me.ServerStatus.Name = "ServerStatus"
        Me.ServerStatus.Size = New System.Drawing.Size(176, 16)
        Me.ServerStatus.TabIndex = 72
        Me.ServerStatus.Text = "SQL Server Status: Checking"
        Me.ServerStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.ServerStatus, "SQL Server Status")
        '
        'ConnectionTesterCounter
        '
        Me.ConnectionTesterCounter.BackColor = System.Drawing.Color.Transparent
        Me.ConnectionTesterCounter.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnectionTesterCounter.ForeColor = System.Drawing.Color.Black
        Me.ConnectionTesterCounter.Location = New System.Drawing.Point(240, 144)
        Me.ConnectionTesterCounter.Name = "ConnectionTesterCounter"
        Me.ConnectionTesterCounter.Size = New System.Drawing.Size(64, 16)
        Me.ConnectionTesterCounter.TabIndex = 73
        Me.ConnectionTesterCounter.Text = "60"
        Me.ConnectionTesterCounter.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.ConnectionTesterCounter, "Countdown to next Server Status Check")
        '
        'StatusBar
        '
        Me.StatusBar.BackColor = System.Drawing.Color.Transparent
        Me.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar.ForeColor = System.Drawing.Color.Black
        Me.StatusBar.Location = New System.Drawing.Point(32, 104)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(440, 40)
        Me.StatusBar.TabIndex = 71
        Me.StatusBar.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.StatusBar, "Status Bar")
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 77
        Me.Label2.Text = "FORCE"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label2, "Application Build Version")
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.BackColor = System.Drawing.Color.Orange
        Me.NumericUpDown1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.NumericUpDown1.ForeColor = System.Drawing.Color.White
        Me.NumericUpDown1.Location = New System.Drawing.Point(24, 8)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {43200, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(64, 13)
        Me.NumericUpDown1.TabIndex = 78
        Me.NumericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.NumericUpDown1, "Monitor Interval in Seconds")
        Me.NumericUpDown1.Value = New Decimal(New Integer() {60, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(88, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 79
        Me.Label3.Text = "seconds"
        Me.ToolTip1.SetToolTip(Me.Label3, "Monitor Interval in Seconds")
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(296, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 16)
        Me.Label4.TabIndex = 80
        Me.Label4.Text = "s to Go"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label4, "Countdown to next Server Status Check")
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(432, 160)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(408, 160)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 66
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(392, 160)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 65
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(376, 160)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 64
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(360, 160)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(344, 160)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.White
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.DetectUrls = False
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.Firebrick
        Me.RichTextBox1.Location = New System.Drawing.Point(32, 30)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(440, 64)
        Me.RichTextBox1.TabIndex = 74
        Me.RichTextBox1.Text = ""
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Firebrick
        Me.Panel1.Location = New System.Drawing.Point(23, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(458, 74)
        Me.Panel1.TabIndex = 75
        '
        'Panel2
        '
        Me.Panel2.Location = New System.Drawing.Point(24, 26)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(456, 72)
        Me.Panel2.TabIndex = 76
        '
        'ConnectionTester
        '
        Me.ConnectionTester.Interval = 60000
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(506, 184)
        Me.Controls.Add(Me.ConnectionTesterCounter)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.ServerStatus)
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.Firebrick
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(512, 216)
        Me.MinimumSize = New System.Drawing.Size(512, 216)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DOS Shell Rollout 1.0"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                'If error_reporting_level = "minimal" Then
                'Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & ex.Message.ToString)
                'Display_Message1.ShowDialog()
                'Else
                '   Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                '  Display_Message1.ShowDialog()
                'End If
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in DOS Shell Rollout's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub run_green_lights()
        If shutting_down = False Then

            Try
                Label1.ForeColor = Color.LimeGreen
                Label1.Text = "Waiting"


                current_light = current_light - 1
                If current_light < 1 Then
                    current_light = 5
                End If
                current_colour = 0



                PictureBox1.Image = ImageList1.Images(1)
                PictureBox2.Image = ImageList1.Images(1)
                PictureBox3.Image = ImageList1.Images(1)
                PictureBox4.Image = ImageList1.Images(1)
                PictureBox5.Image = ImageList1.Images(1)



                Select Case current_light
                    Case 0

                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1

                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2

                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3

                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4

                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5

                        PictureBox1.Image = ImageList1.Images(0)
                End Select

                current_light = current_light + 1
                If current_light > 5 Then
                    current_light = 1
                End If
            Catch err As System.InvalidOperationException
                current_light = current_light
            Catch ex As Exception
                Error_Handler(ex, "Running Lights")
            End Try
        End If
    End Sub

    Private Sub run_orange_lights()
        If shutting_down = False Then


            Try
                Label1.ForeColor = Color.DarkOrange
                Label1.Text = "Working"

                current_light = current_light - 1
                If current_light < 1 Then
                    current_light = 5
                End If
                current_colour = 1



                PictureBox1.Image = ImageList1.Images(3)
                PictureBox2.Image = ImageList1.Images(3)
                PictureBox3.Image = ImageList1.Images(3)
                PictureBox4.Image = ImageList1.Images(3)
                PictureBox5.Image = ImageList1.Images(3)

                Select Case current_light
                    Case 0
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox1.Image = ImageList1.Images(2)
                End Select

                current_light = current_light + 1
                If current_light > 5 Then
                    current_light = 1
                End If
            Catch err As System.InvalidOperationException
                current_light = current_light
            Catch ex As Exception
                Error_Handler(ex, "Running Lights")
            End Try
        End If
    End Sub

    Private Sub run_lights()
        If shutting_down = False Then


            Try
                If current_colour = 1 Then
                    Select Case current_light
                        Case 0
                            PictureBox5.Image = ImageList1.Images(3)
                            PictureBox1.Image = ImageList1.Images(2)
                        Case 1
                            PictureBox1.Image = ImageList1.Images(3)
                            PictureBox2.Image = ImageList1.Images(2)
                        Case 2
                            PictureBox2.Image = ImageList1.Images(3)
                            PictureBox3.Image = ImageList1.Images(2)
                        Case 3
                            PictureBox3.Image = ImageList1.Images(3)
                            PictureBox4.Image = ImageList1.Images(2)
                        Case 4
                            PictureBox4.Image = ImageList1.Images(3)
                            PictureBox5.Image = ImageList1.Images(2)
                        Case 5
                            PictureBox5.Image = ImageList1.Images(3)
                            PictureBox1.Image = ImageList1.Images(2)
                    End Select
                Else
                    Select Case current_light
                        Case 0
                            PictureBox5.Image = ImageList1.Images(1)
                            PictureBox1.Image = ImageList1.Images(0)
                        Case 1
                            PictureBox1.Image = ImageList1.Images(1)
                            PictureBox2.Image = ImageList1.Images(0)
                        Case 2
                            PictureBox2.Image = ImageList1.Images(1)
                            PictureBox3.Image = ImageList1.Images(0)
                        Case 3
                            PictureBox3.Image = ImageList1.Images(1)
                            PictureBox4.Image = ImageList1.Images(0)
                        Case 4
                            PictureBox4.Image = ImageList1.Images(1)
                            PictureBox5.Image = ImageList1.Images(0)
                        Case 5
                            PictureBox5.Image = ImageList1.Images(1)
                            PictureBox1.Image = ImageList1.Images(0)
                    End Select

                End If

                current_light = current_light + 1
                If current_light > 5 Then
                    current_light = 1
                End If
            Catch err As System.InvalidOperationException
                current_light = current_light
            Catch ex As Exception
                Error_Handler(ex, "Running Lights")
            End Try
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")

            Try
                Dim filereader As IO.StreamReader = New StreamReader((Application.StartupPath & "\").Replace("\\", "\") & "config.ini")

                While filereader.Peek > -1
                    Dim lineread As String
                    lineread = filereader.ReadLine
                    If lineread.StartsWith("Server:") = True Then
                        Worker1.dbserver = lineread.Replace("Server:", "").Trim
                    End If

                    If lineread.StartsWith("Table:") = True Then
                        Worker1.dbtable = lineread.Replace("Table:", "").Trim
                    End If
                    If lineread.StartsWith("User:") = True Then
                        Worker1.dbuser = lineread.Replace("User:", "").Trim
                    End If
                    If lineread.StartsWith("Errors:") = True Then
                        error_reporting_level = lineread.Replace("Errors:", "").Trim
                    End If
                End While

                filereader.Close()
            Catch ex As Exception
                Dim displ As Display_Message = New Display_Message("No valid config.ini file could be located in the application startup folder. This application will now be forced to shutdown.")
                displ.ShowDialog()
                application_exit = True

            End Try




            ServerStatus.Text = Worker1.serverstatus.Text
            If Not Worker1.dbserver = "" And Not Worker1.dbserver = Nothing Then
                ToolTip1.SetToolTip(ServerStatus, "Status of: " & Worker1.dbserver)
            End If
            ServerStatus.ForeColor = Worker1.serverstatus.ForeColor
            Label4.ForeColor = Worker1.serverstatus.ForeColor
            ConnectionTesterCounter.ForeColor = ServerStatus.ForeColor
            'If currently_working = False Then


            ConnectionTesterCounter.Text = (CInt(ConnectionTesterCounter.Text) - 1).ToString
            'If ConnectionTesterCounter.Text.Length < 2 Then
            'ConnectionTesterCounter.Text = "0" & ConnectionTesterCounter.Text
            'End If

            If CInt(ConnectionTesterCounter.Text) = -1 Then
                force_check(1)
                ConnectionTesterCounter.Text = (ConnectionTester.Interval / 1000)
            End If

            ' End If

            If application_exit = True Then
                Me.Close()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Load_Registry_Values()

            If File.Exists((Application.StartupPath & "\").Replace("\\", "\") & "config.ini") = False Then
                If File.Exists((Application.StartupPath & "\").Replace("\\", "\") & "default_config.ini") = True Then
                    File.Copy((Application.StartupPath & "\").Replace("\\", "\") & "default_config.ini", (Application.StartupPath & "\").Replace("\\", "\") & "config.ini")
                End If
            End If

            If File.Exists((Application.StartupPath & "\").Replace("\\", "\") & "config.ini") = True Then
                Dim server As String = ""
                Dim table As String = ""
                Dim user As String = ""

                Try
                    Dim filereader As IO.StreamReader = New StreamReader((Application.StartupPath & "\").Replace("\\", "\") & "config.ini")
                    While filereader.Peek > -1
                        Dim lineread As String
                        lineread = filereader.ReadLine
                        If lineread.StartsWith("Server:") = True Then
                            server = lineread.Replace("Server:", "").Trim
                        End If

                        If lineread.StartsWith("Table:") = True Then
                            table = lineread.Replace("Table:", "").Trim
                        End If
                        If lineread.StartsWith("User:") = True Then
                            user = lineread.Replace("User:", "").Trim
                        End If
                        If lineread.StartsWith("Errors:") = True Then
                            error_reporting_level = lineread.Replace("Errors:", "").Trim
                        End If
                    End While

                    Worker1.dbserver = server
                    Worker1.dbuser = user
                    Worker1.dbtable = table

                    filereader.Close()
                Catch ex As Exception
                    Dim displ As Display_Message = New Display_Message("No valid config.ini file could be located in the application startup folder. This application will now be forced to shutdown.")
                    displ.ShowDialog()
                    dataloaded = True
                    splash_loader.Visible = False
                    application_exit = True
                    Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
                    Timer2.Start()
                End Try
                If application_exit = False Then


                    Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
                    Timer2.Start()
                    'ConnectionTester.Start()


                    dataloaded = True
                    splash_loader.Visible = False

                    Try
                        Dim ApplicationName As String
                        ApplicationName = "DOS Shell Rollout"
                        Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
                        Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)
                        If Process.GetProcessesByName(aProcName).Length > 1 Then
                            Me.Hide()
                            Dim Display_Message1 As New Display_Message("Another Instance of " & ApplicationName & " is already running. Only one instance of " & ApplicationName & " is allowed to run at any time. This instance of the program will now commence shut down operations.")
                            Display_Message1.ShowDialog()
                            application_exit = True
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Checking Multiple Application Instances")
                    End Try
                    force_check(1)
                End If
            Else
                Dim displ As Display_Message = New Display_Message("No required config.ini file could be located in the application startup folder. This application will no be forced to shutdown.")
                displ.ShowDialog()
                dataloaded = True
                splash_loader.Visible = False
                application_exit = True
                Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
                Timer2.Start()
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            shutting_down = True
            Timer2.Stop()
            Save_Registry_Values()
            '            ConnectionTester.Stop()
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex, "Shutting Down Application")
        Finally
            Application.Exit()
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerHandler(ByVal Result As String)
        Try
            StatusBar.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & Result

        Catch ex As Exception
            Error_Handler(ex)
        Finally
            currently_working = False

            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        End Try
    End Sub



    Public Sub WorkerErrorEncounteredHandler(ByVal ex As Exception, ByVal descriptor As String)
        Try
            Error_Handler(ex, descriptor)
        Catch exc As Exception
            Error_Handler(exc)
        End Try
    End Sub

    Private Sub run_worker(Optional ByVal threadselect As Integer = 1)
        run_orange_lights()

        Select Case threadselect
            Case 1
                Worker1.ChooseThreads(1)
                'Case 3

                'If testingconnection = False Then
                '    ConnectionTesterCounter.Text = 0
                '    ServerStatus.ForeColor = Color.Orange
                '    ServerStatus.Text = "SQL Server Status: Checking"
                '    ToolTip1.SetToolTip(ServerStatus, "SQL Server connection status on " & Worker1.dbserver & " (user: " & Worker1.dbuser & ")")
                '    ServerStatus.Refresh()
                '    testingconnection = True
                '    Worker1.ChooseThreads(3)
                'End If
        End Select

        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Protected Friend Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check(Optional ByVal threadselect As Integer = 1)
        Try
            NotifyIcon1.Text = "Processing Request..."
            If currently_working = False Then

                Select Case threadselect
                    Case 1
                        ' ConnectionTesterCounter.Text = "00"
                        run_worker(threadselect)
                    Case 3
                        run_worker(threadselect)
                End Select
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub WorkerStatusMessageUpdate(ByVal message As String)
        Try
            If RichTextBox1.Text.Length < 1 Then
                RichTextBox1.Text = RichTextBox1.Text.Insert(0, message)
            Else
                RichTextBox1.Text = RichTextBox1.Text.Insert(0, message & vbCrLf)
            End If

            RichTextBox1.Select(0, 0)
            StatusBar.Text = message
        Catch ex As Exception
            Error_Handler(ex, "Worker Status Message Update")
        End Try
    End Sub


    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        force_check(1)
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged, NumericUpDown1.TextChanged
        Try
            ' ConnectionTester.Stop()
            ConnectionTester.Interval = NumericUpDown1.Value * 1000
            ConnectionTesterCounter.Text = NumericUpDown1.Value
            'ConnectionTester.Start()
        Catch ex As Exception
            Error_Handler(ex, "Updating Monitor Timer Interval")
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\DOS Shell Rollout") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\DOS Shell Rollout")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\DOS Shell Rollout", True)


            str = oKey.GetValue("interval")
            If Not IsNothing(str) And Not (str = "") Then
                NumericUpDown1.Value = CInt(str)
            Else
                oKey.SetValue("interval", 60)
                NumericUpDown1.Value = 60
            End If


            oKey.SetValue("executablePath", """" & Application.ExecutablePath & """")



            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\DOS Shell Rollout", True)




            oKey.SetValue("interval", NumericUpDown1.Value)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


End Class
