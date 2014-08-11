Imports Microsoft.Win32
Imports System.IO

Public Class Login_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerCompletehHandler(ByVal Result As String)
    Public Delegate Sub WorkerErrorEncounteredhHandler(ByVal ex As Exception)

    Private shutting_down As Boolean = False

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
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
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ContextSelector As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtContext As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblContext As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblError As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Login_Screen))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtContext = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.txtUsername = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ContextSelector = New System.Windows.Forms.ContextMenu
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblError = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.lblContext = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
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
        Me.Label8.ForeColor = System.Drawing.Color.Silver
        Me.Label8.Location = New System.Drawing.Point(104, 184)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 16)
        Me.Label8.TabIndex = 33
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Silver
        Me.Label9.Location = New System.Drawing.Point(8, 184)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(88, 16)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "BUILD 20051122.1"
        Me.ToolTip1.SetToolTip(Me.Label9, "Screen Build Number")
        '
        'txtContext
        '
        Me.txtContext.BackColor = System.Drawing.Color.White
        Me.txtContext.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtContext.ForeColor = System.Drawing.Color.Firebrick
        Me.txtContext.Location = New System.Drawing.Point(32, 104)
        Me.txtContext.Name = "txtContext"
        Me.txtContext.ReadOnly = True
        Me.txtContext.Size = New System.Drawing.Size(184, 20)
        Me.txtContext.TabIndex = 120
        Me.txtContext.Text = "Staff.com.main.uct"
        Me.ToolTip1.SetToolTip(Me.txtContext, "Context to login with.")
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.Color.White
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPassword.ForeColor = System.Drawing.Color.Firebrick
        Me.txtPassword.Location = New System.Drawing.Point(32, 64)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(184, 20)
        Me.txtPassword.TabIndex = 119
        Me.txtPassword.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtPassword, "Novell Network Password")
        '
        'txtUsername
        '
        Me.txtUsername.BackColor = System.Drawing.Color.White
        Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUsername.ForeColor = System.Drawing.Color.Firebrick
        Me.txtUsername.Location = New System.Drawing.Point(32, 24)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(184, 20)
        Me.txtUsername.TabIndex = 118
        Me.txtUsername.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtUsername, "Novell Network Username")
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Orange
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(288, 104)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 123
        Me.Button1.Text = "LOGIN"
        Me.ToolTip1.SetToolTip(Me.Button1, "Login to DOS Shell Rollout")
        '
        'ContextSelector
        '
        Me.ContextSelector.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MenuItem5})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.Text = "Helpdesk Tutor"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "Commerce I.T. Staff"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(392, 184)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(368, 184)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 66
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(352, 184)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 65
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(336, 184)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 64
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(320, 184)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(304, 184)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.AcceptsTab = True
        Me.RichTextBox1.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.DetectUrls = False
        Me.RichTextBox1.ForeColor = System.Drawing.Color.DimGray
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 232)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(416, 64)
        Me.RichTextBox1.TabIndex = 110
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Firebrick
        Me.Panel1.Controls.Add(Me.lblError)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.txtContext)
        Me.Panel1.Controls.Add(Me.txtPassword)
        Me.Panel1.Controls.Add(Me.txtUsername)
        Me.Panel1.Controls.Add(Me.lblContext)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(0, 32)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(456, 144)
        Me.Panel1.TabIndex = 115
        '
        'lblError
        '
        Me.lblError.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.lblError.Location = New System.Drawing.Point(240, 72)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(200, 32)
        Me.lblError.TabIndex = 124
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(240, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(200, 56)
        Me.Label6.TabIndex = 122
        Me.Label6.Text = "YOU NEED TO AUTHENTICATE YOURSELF TO THE SYSTEM BY LOGGING IN USING YOUR NOVELL N" & _
        "ETWORK CREDENTIALS. NOTE THAT YOU WILL NEED TO BE A COMMERCE I.T. STAFF MEMBER T" & _
        "O GAIN ACCESS TO THIS SYSTEM."
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(128, 124)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 8)
        Me.Label5.TabIndex = 121
        Me.Label5.Text = "RIGHT-CLICK TO CHANGE"
        '
        'lblContext
        '
        Me.lblContext.Location = New System.Drawing.Point(32, 88)
        Me.lblContext.Name = "lblContext"
        Me.lblContext.Size = New System.Drawing.Size(168, 16)
        Me.lblContext.TabIndex = 117
        Me.lblContext.Text = "Context: (Commerce I.T. Staff)"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 116
        Me.Label3.Text = "Password:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 115
        Me.Label2.Text = "Username:"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(416, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(16, 16)
        Me.Button2.TabIndex = 125
        Me.Button2.TabStop = False
        Me.Button2.Text = "Button2"
        Me.Button2.Visible = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Firebrick
        Me.Label4.Location = New System.Drawing.Point(8, 2)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(368, 23)
        Me.Label4.TabIndex = 116
        Me.Label4.Text = "DOS Shell Rollout"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Login_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(450, 208)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Button2)
        Me.ForeColor = System.Drawing.Color.Firebrick
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(456, 240)
        Me.MinimumSize = New System.Drawing.Size(456, 240)
        Me.Name = "Login_Screen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "System Login"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
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
        Catch ex As Exception
            Error_Handler(ex, "Running Lights")
            End Try
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex, "System Timer")
        End Try
    End Sub

    Private Sub Login_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try


            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            Timer2.Start()
            txtUsername.Select()
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex, "Screen Load")
        End Try
    End Sub

    Private Sub exit_application()
        Try
            shutting_down = True
            Timer2.Stop()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex, "Application Exit")
        End Try
    End Sub

    Private Sub Login_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex, "Application Exit")
        End Try
    End Sub


    Public Sub WorkerCompleteHandler(ByVal Result As String)
        Try
            currently_working = False
            If Result.StartsWith("Result") Then
                Result = Result.Substring(13)
            End If
            lblError.Text = Result

            If lblError.Text.ToLower.StartsWith("success") = True Then
                Dim use As String = txtUsername.Text
                run_green_lights()
                Me.Hide()
                Dim Splash As New Splash_Screen
                Splash.Show()
                While Splash.dataloaded = False
                End While
                Dim mainscr As New Main_Screen(Splash, use.Trim)
                Try
                    mainscr.ShowDialog()
                    mainscr.Visible = False
                Catch ex As Exception
                End Try
                While mainscr.dataloaded = False
                End While
                Try
                    mainscr.Visible = True
                    mainscr.show_application()

                Catch ex As System.ObjectDisposedException
                Catch exr As Exception
                    Error_Handler(exr, "Disposed Main Screen Check")
                End Try
                exit_application()
            End If
            run_orange_lights()
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex, "Thread Complete")
        End Try
    End Sub




    Public Sub WorkerErrorEncounteredHandler(ByVal ex As Exception)
        Try
            Error_Handler(ex, "Worker Error Encountered")
        Catch exc As Exception
            Error_Handler(exc, "Thread Error")
        End Try
    End Sub

    Private Sub run_worker(Optional ByVal threadselect As Integer = 1)
        Try


            run_orange_lights()

            Worker1.username = txtUsername.Text.Trim
            Worker1.password = txtPassword.Text.Trim
            Worker1.context = txtContext.Text.Trim
            If txtContext.Text.Trim = "Students.com.main.uct" Then
                Worker1.group = "tutors_g"
            Else
                Worker1.group = "cf_allstaff_g,"
            End If
            Worker1.error_level = "none"
            Select Case threadselect
                Case 1
                    Worker1.ChooseThreads(1)
            End Select

            currently_working = True
        Catch ex As Exception
            Error_Handler(ex, "Launching Thread")
        End Try
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub







    Private Sub force_check(Optional ByVal threadselect As Integer = 1)
        Try
            lblError.Text = ""
            If currently_working = False Then
                run_worker(threadselect)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            lblContext.Text = "Context: (Helpdesk Tutor)"
            txtContext.Text = "Students.com.main.uct"
        Catch ex As Exception
            Error_Handler(ex, "Login Procedure")
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            lblContext.Text = "Context: (Commerce I.T. Staff)"
            txtContext.Text = "Staff.com.main.uct"
        Catch ex As Exception
            Error_Handler(ex, "Login Procedure")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        force_check()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtUsername.Text = "clotter"
        txtPassword.Text = "finalfantasy"
        txtContext.Text = "Staff.com.main.uct"
        force_check()
    End Sub
End Class
