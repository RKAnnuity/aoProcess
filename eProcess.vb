
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
'Imports Microsoft.Office.Interop
'Imports SHDocVw

Public Class eProcess
    'Public sSITE As String = "iTeleCollect"
    Public sSITE As String = "AnnuityOne"

    Public msSQLAOconnection As String = "", msSQLAOuser As String = ""
    Public msSQLConnectionString As String = "", msSQLuser As String = ""
    Public msSQL2ConnectionString As String = "", msSQL2user As String = ""
    Public msSQL3ConnectionString As String = "", msSQL3user As String = ""
    Public DB2SQLConnectionString As String = "", DB2SQLuser As String = ""

    'Public dir_DATA As String = "\\TeleServer\TeleServer$\DATA\"
    'Public dir_EMAIL As String = "\\TeleServer\TeleServer$\SCHEDULE\email\"
    Public dir_REPORT As String = "\\production\Reports\"
    '**************************************************************************
    '* 2018-08-07 RFK: "\\TeleServer\TeleServer$\DATA\" is GONE
    Public dir_CHK As String = "\\production\automation$\CHK\"
    Public dir_EMAIL As String = "\\production\automation$\email\"
    Public dir_EXE As String = "\\production\Required_Files\EXE\"
    Public dir_INI As String = "\\production\Required_Files\INI\"
    'Public dir_LOG As String = "\\reporting\report$\LOG\"
    Public dir_LOG As String = "\\production\reports\LOG\"
    Public dir_SCHEDULE As String = "\\production\automation$\SCHEDULE\"

    Public tUNIQUE As String = "", tCLIENT As String = "", tTOB As String = "", tFACILITY As String = ""
    Public tLOCX As String = "", tACCOUNTNUMBER As String = "", tSUFFIX As String = ""
    Public tSTATUS As String = "", tCC As String = "", t1 As String = "", tMC As String = ""
    Public tNEXTDATE As String = "", tLETTER As String = "", tLETTERDATE As String = "", tPAYMENTTYPE As String = "", tPAYMENTDATE As String = ""
    Public tNOTE As String = "", tDATE As String = "", tBY As String = ""
    Public tEMAILfrom As String = "", tEMAILfromname As String = "", tEMAILto As String = "", tEMAILtoname As String = ""
    Public tEMAILcc As String = "", tEMAILbcc As String = "", tEMAILsubject As String = "", tEMAILmessage As String = "", tEMAILattach As String = ""
    Public iLastQryHour As Integer, iLastQryMinute As Integer
    Public sEncryptionPW As String = "Ryan.Kiechle_2194404440"
    Public swQRYtable As Boolean = False
    Public swTEST As Boolean = False
    Public DTqry As New DataTable

    Private Sub iScraper_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            '**************************************************************************************
            Me.Text = "aoProcessor v202206.29"
            '**************************************************************************************
            '* 2012-04-27 RFK:
            '* 2011-11-01 RFK: commands
            Select Case sSITE
                Case "AnnuityOne"
                    msSQLAOconnection = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "msSQLAOconnection.INI"), sEncryptionPW)
                    msSQLAOuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "sa_Dialer-PW.INI"), sEncryptionPW)
                    msSQLConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "msSQLConnectionString.INI"), sEncryptionPW)
                    msSQLuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "sa_Dialer-PW.INI"), sEncryptionPW)
                    msSQL2ConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "msSQL2ConnectionString.INI"), sEncryptionPW)
                    msSQL2user = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "sa_Dialer-PW.INI"), sEncryptionPW)
                    msSQL3ConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "msSQL3ConnectionString.INI"), sEncryptionPW)
                    msSQL3user = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "msSQL3user.INI"), sEncryptionPW)
                    DB2SQLConnectionString = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "DB2SQLConnectionString.INI"), sEncryptionPW)
                    DB2SQLuser = rkutils.Decrypt(".NET", rkutils.FILE_read(dir_INI + "DB2SQLuser-PW.INI"), sEncryptionPW)
                Case "iTeleCollect"
                    '******************************************************************************
            End Select
            '**************************************************************************************
            ComboBox1.Items.Clear()
            ComboBox1.Items.Add("")
            ComboBox1.Items.Add("CallList")
            ComboBox1.Items.Add("Report")
            ComboBox1.Items.Add("WorkQUE")
            ComboBox1.SelectedIndex = 0
            '**************************************************************************************
            DataGridView1.Visible = False
            DataGridView2.Visible = False
            DataGridView_COMMANDS.Visible = False
            DataGridView_QRYs.Visible = False
            DataGridView_QRYoutput.Visible = False
            '**************************************************************************************
            'swTEST = True
            If swTEST Then
                CheckBox_RUN.Checked = True
                CheckBox_Status.Checked = False
                CheckBox_EMailSecure.Checked = False
                CheckBox_Note.Checked = False
                CheckBox_Track.Checked = False
                CheckBox_Queries.Checked = False
                CheckBox_Queries_TCode.Checked = False
                CheckBox_Queries_TCodeMatched.Checked = False
                Button_TEST.Visible = True
            Else
                CheckBox_RUN.Checked = True
                CheckBox_Status.Checked = True
                CheckBox_EMailSecure.Checked = True
                CheckBox_Note.Checked = True
                CheckBox_Track.Checked = True
                CheckBox_Queries.Checked = True
                CheckBox_Queries_TCode.Checked = True
                CheckBox_Queries_TCodeMatched.Checked = False
                Panel_Commands.Visible = True
                Panel_Scrape.Visible = True
                Panel_Query.Visible = True
            End If
            '******************************************************************
            '* 2015-04-22 RFK: 
            Dim sCommand = Command()
            If swTEST Then sCommand = "/TEST"
            'sCommand = "/2"
            'sCommand = "/QRYRUN REPORT WorkFlowS Balance Zero"
            'sCommand = "/QRYRUNLIVE REPORT Ryan.Kiechle"
            'sCommand = "/QRYRUN CallList SelfPay TEST"
            'sCommand = "/STATUSNEXT"
            'sCommand = "/QRYRUN REPORT IT PRONTO UCM SELFPAY SCRUB"
            'sCommand = "/QRYRUN CallList Collections GEN1 - NW"
            '******************************************************************
            Select Case STR_BREAK(sCommand, 1)
                Case "/STATUSNEXT", "/CHECKSTATUSNEXT"
                    Me.Text += "[Check Status Next]"
                    CheckBox_RUN.Checked = True
                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = False
                    CheckBox_Track.Checked = False
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    Button_TEST.Visible = False
                    '**********************************************************
                    CheckStatusNext()
                    '**********************************************************
                    End
                    '**********************************************************
                Case "/TEST"
                    Me.Text += "[TEST]"
                    'Handled by TIMER
                Case "/1"
                    Text += " [1]"
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    CheckBox_CallList.Checked = False
                    Panel_Query.Visible = False
                Case "/2"
                    Text += " [2]"
                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = False
                    CheckBox_Track.Checked = False
                    Panel_Commands.Visible = False
                    Panel_Scrape.Visible = False
                    'CheckBox_RUN.Checked = False
                Case "/STATUS"
                    Text += " [STATUS " + STR_BREAK(sCommand, 2) + "]"
                    Panel_Commands.Visible = True

                    CheckBox_Status.Checked = True
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = False
                    CheckBox_Track.Checked = False

                    Panel_Query.Visible = False
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    CheckBox_CallList.Checked = False

                    Panel_Scrape.Visible = False
                Case "/NOTE"
                    Text += " [NOTE " + STR_BREAK(sCommand, 2) + "]"
                    Panel_Commands.Visible = True

                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = True
                    CheckBox_Track.Checked = False

                    Panel_Query.Visible = False
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    CheckBox_CallList.Checked = False

                    Panel_Scrape.Visible = False
                Case "/TRACK"
                    Text += " [TRACK " + STR_BREAK(sCommand, 2) + "]"
                    Panel_Commands.Visible = True

                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = True
                    CheckBox_Track.Checked = False

                    Panel_Query.Visible = False
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    CheckBox_CallList.Checked = False

                    Panel_Scrape.Visible = False
                Case "/QRY"
                    Me.Text += " [QRY " + STR_BREAK(sCommand, 2) + "]"
                    CheckBox_Queries.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False
                    CheckBox_CallList.Checked = False

                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = False
                    CheckBox_Track.Checked = False
                    CheckBox_Queries_TCode.Checked = False
                    CheckBox_Queries_TCodeMatched.Checked = False

                    Panel_Commands.Visible = False
                    Panel_Scrape.Visible = False
                    Panel_Query.Visible = False
                Case "/QRYRUN", "/QRYRUNLIVE"
                    MsgStatus(sCommand, True)
                    Me.Show()
                    Me.Refresh()
                    rkutils.DoEvents()
                    '**********************************************************
                    CheckBox_Status.Checked = False
                    CheckBox_EMailSecure.Checked = False
                    CheckBox_Note.Checked = False
                    CheckBox_Track.Checked = False
                    Panel_Commands.Visible = False
                    Panel_Scrape.Visible = False
                    '**********************************************************
                    Dim sQRYtype As String = STR_BREAK_PIECES(sCommand, 2, " ")
                    MsgStatus("Type=" + sQRYtype, True)
                    Dim sQRYwho As String = STR_BREAK_PIECES(sCommand, 3, " ")
                    MsgStatus("Who=" + sQRYwho, True)
                    Dim sQRYname As String = STR_BREAK(STR_BREAK(STR_BREAK(sCommand, 2), 2), 2) 'Need to get ALL of the parameters after that SPACE
                    MsgStatus("Name=" + sQRYname, True)
                    '**********************************************************
                    Me.Text += " [QRYRUN " + sQRYtype + " " + sQRYwho + " " + sQRYname + "]"
                    MsgStatus(Me.Text, True)
                    '**********************************************************
                    Dim sSQL As String = "SELECT * FROM RevMD.dbo.query WHERE TYPE='" + sQRYtype + "' AND WHO='" + sQRYwho + "' AND NAME='" + sQRYname + "'"
                    MsgStatus(sSQL, True)
                    DataGridView_QRYs.Visible = SQL_READ_DATAGRID(DataGridView_QRYs, "MSSQL", "*", msSQLConnectionString, msSQLuser, sSQL)
                    '**********************************************************
                    QRYread(0)
                    '**********************************************************
                    QRYselect()
                    rkutils.DoEvents()
                    Dim iCount As Integer = DataGridView_QRYoutput.Rows.Count
                    '**********************************************************
                    Dim swLive As Boolean = False
                    Select Case STR_BREAK(sCommand, 1)
                        Case "/QRYRUNLIVE"
                            swLive = True
                        Case Else
                            swLive = False
                    End Select
                    '**********************************************************
                    Select Case sQRYtype.ToUpper
                        Case "CallList".ToUpper
                            Dim sCallListName As String = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_NAME", 0).Trim
                            MsgStatus("Building CallList:" + sCallListName + " Selected:" + Str(iCount).Trim, True)
                            If iCount > 0 Then
                                Label_RUNNING.Text = "Running"
                                CallList_ADD(rkutils.STR_TRIM(rkutils.STR_BREAK(sCallListName, 1) + "-" + rkutils.STR_format("TODAY", "mmdd"), 10), rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_GROUP", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_DIALER", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_APPTYPE", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_RATIO", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_START", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_STOP", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_INSERT", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_ACCOUNTNUMBER", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_CLIENTOUTPUT", 0).Trim)
                            End If
                        Case "Report".ToUpper
                            If iCount > 0 Then
                                If swQRYtable Then
                                    QRYrunDT(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, swLive)
                                Else
                                    QRYrun(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, swLive)
                                End If
                            End If
                    End Select
                    '**********************************************************
                    If swLive Then
                        QRYcompleteLive()
                    Else
                        QRYcomplete()
                    End If
                    '**********************************************************
                    MsgStatus("Complete", True)
                    End
                    '**********************************************************
                Case Else
                    CheckBox_RUN.Checked = False
                    MsgStatus("DEFAULT:" + sCommand, True)
                    MsgBox("Default:" + sCommand + " [ENDING]")
                    End
            End Select
            '******************************************************************
            '* 2015-04-22 RFK: 
            If swTEST = False Then
                QRYblank()
            End If
            InitApp()
            ResizeApp()
            '******************************************************************
            Label_QRY_Running.Text = ""
            '******************************************************************
            Select Case STR_BREAK(sCommand, 1)
                Case "/TEST"
                    'Nothing Here
            End Select
            Timer1.Enabled = True
            Timer1.Interval = 1100
            '****************************************************
        Catch ex As Exception
            MsgError("eProcess_Load", ex.ToString)
        End Try
    End Sub

    Private Sub eProcess_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ExitApp()
    End Sub

    Private Sub eProcess_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        ResizeApp()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Try
            Label_TIME.Text = STR_TRIM(Now.TimeOfDay.ToString, 8)
            System.Windows.Forms.Application.DoEvents()
            If CheckBox_RUN.Checked Then
                '**************************************************************
                If Me.Text.Contains("[2]") = False Then
                    If Label_RUNNING.Text = "Ready" Then
                        Label_RUNNING.Text = "Running"
                        ProcessInit()
                        If Label_RUNNING.Text = "Running" Then Label_RUNNING.Text = "Ready"
                    End If
                End If
                '**************************************************************
                '* 2012-10-12 RFK: Queries
                If CheckBox_Queries.Checked Then
                    '**********************************************************
                    '* 2017-01-11 RFK: Only 1 time per hour
                    If iLastQryHour <> Now.TimeOfDay.Hours Then
                        iLastQryHour = Now.TimeOfDay.Hours  'Reset hour
                        If Label_RUNNING.Text = "Ready" Then
                            CheckStatusNext_shell()
                        End If
                    End If
                    '**********************************************************
                    '* 2012-10-12 RFK: Only 1 time per minute
                    If iLastQryMinute <> Now.TimeOfDay.Minutes Then
                        If Label_RUNNING.Text = "Ready" Then
                            QRYinit()
                            iLastQryMinute = Now.TimeOfDay.Minutes  'Reset minute
                        End If
                    End If
                End If
                '**************************************************************
                '* 2012-10-12 RFK: scrapes
                '* 2015-04-28 RFK: Do Not do for [2]
                If Me.Text.Contains("[2]") = False Then
                    If Label_RUNNING.Text = "Ready" Then
                        Label_RUNNING.Text = "Running"
                        ScrapeInit()
                        If Label_RUNNING.Text = "Running" Then Label_RUNNING.Text = "Ready"
                    End If
                End If
            End If
            If Me.Text.Contains("[2]") Then
                FILE_create(dir_CHK + "aoProcessor2.CHK", True, False, Date.Now.ToString)
            Else
                FILE_create(dir_CHK + "aoProcessor.CHK", True, False, Date.Now.ToString)
            End If
        Catch ex As Exception
            MsgError("Timer1", ex.ToString)
        End Try
    End Sub

    Private Sub ExitApp()
        Try
            FILE_create(dir_INI + "aoProcessor-" + rkutils.WhoAmI() + ".INI", True, False, Me.Top.ToString + " " + Me.Left.ToString + " " + Me.Height.ToString + " " + Me.Width.ToString + vbCrLf)
            End
        Catch ex As Exception
            MsgError("ExitApp", ex.ToString)
        End Try
    End Sub

    Private Sub MsgError(ByVal tMODULE As String, ByVal tMSG As String)
        Try
            '***************************************************************************
            FILE_create(dir_LOG + "aoProcessor-" + DateToday(8) + ".LOG", False, True, DateToday(18) + " ERROR:" + tMODULE + " " + tMSG + vbCrLf)
            '***************************************************************************
            MsgStatus("ERROR:" + tMODULE + "_" + tMSG, True)
        Catch ex As Exception
            'MsgError("ExitApp", ex.ToString)
        End Try
    End Sub

    Public Sub MsgStatus(ByVal tMSG As String, ByVal iScreen As Boolean)
        Try
            '***************************************************************************
            FILE_create(dir_LOG + "aoProcessor-" + DateToday(8) + ".LOG", False, True, DateToday(18) + " " + tMSG + vbCrLf)
            '***************************************************************************
            If iScreen Then
                Me.ListBox_LOG.Items.Add(DateToday(18) + " " + tMSG)
                If Me.ListBox_LOG.Items.Count > 500 Then
                    Me.ListBox_LOG.Items.RemoveAt(1)
                End If
                Me.ListBox_LOG.SelectedIndex = Me.ListBox_LOG.Items.Count - 1
            End If
        Catch ex As Exception
            MsgError("MsgStatus", ex.ToString)
        End Try
    End Sub

    Private Sub InitApp()
        Try
            '**************************************************************
            '* 2015-04-23 RFK:
            Dim tFileINI As String = dir_INI + "aoProcessor-" + rkutils.WhoAmI() + ".INI"
            If IS_File(tFileINI) Then
                Dim tREAD As String = FILE_read(tFileINI)
                Dim iTop As Integer = Val(STR_BREAK_PIECES(tREAD, 1, " "))
                Dim iLeft As Integer = Val(STR_BREAK_PIECES(tREAD, 2, " "))
                Dim iHeight As Integer = Val(STR_BREAK_PIECES(tREAD, 3, " "))
                Dim iWidth As Integer = Val(STR_BREAK_PIECES(tREAD, 4, " "))
                'If iTop <= My.Computer.Screen.Bounds.Height Then Me.Top = iTop
                'If iLeft <= My.Computer.Screen.Bounds.Left Then Me.Left = iLeft
                If iTop < 1 Then iTop = 1
                If iLeft < 1 Then iLeft = 1
                Me.Top = iTop
                Me.Left = iLeft
                If iHeight > 0 Then Me.Height = iHeight
                If iWidth > 0 Then Me.Width = iWidth
                '**************************************************************
                '* 2015-04-23 RFK:
                MsgStatus("Height=" + My.Computer.Screen.Bounds.Height.ToString + " Top=" + My.Computer.Screen.Bounds.Top.ToString + " Me.Height=" + Me.Height.ToString, True)
                If Text.Contains("[1]") Then
                    Left = 1
                    Top = 1
                    Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                Else
                    If Text.Contains("[2]") Then
                        Left = 1
                        Top = (My.Computer.Screen.Bounds.Height - 50) / 2
                        Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                    Else
                        If Text.Contains("[STATUS") Then
                            Width = 290
                            If Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) > 9 Then
                                Me.Top = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 20
                                Me.Left = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 20
                            Else
                                If Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) > 5 Then
                                    Me.Top = (My.Computer.Screen.Bounds.Height - 50) / 2
                                    Me.Left = (Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) - 4) * 290
                                Else
                                    Me.Top = 1
                                    Me.Left = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 290
                                End If
                            End If
                            Me.Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                        Else
                            If Me.Text.Contains("[NOTE") Then
                                Me.Width = 290
                                If Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) > 5 Then
                                    Me.Top = (My.Computer.Screen.Bounds.Height - 50) / 2
                                    Me.Left = (Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) - 4) * 290
                                Else
                                    Me.Top = 1
                                    Me.Left = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 290
                                End If
                                Me.Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                            Else
                                If Me.Text.Contains("[TRACK") Then
                                    Me.Width = 290
                                    If Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) > 5 Then
                                        Me.Top = (My.Computer.Screen.Bounds.Height - 50) / 2
                                        Me.Left = (Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) - 4) * 290
                                    Else
                                        Me.Top = 1
                                        Me.Left = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 290
                                    End If
                                    Me.Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                                Else
                                    If Me.Text.Contains("[QRY") Then
                                        Me.Width = 290
                                        If Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) > 5 Then
                                            Me.Top = (My.Computer.Screen.Bounds.Height - 50) / 2
                                            Me.Left = (Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) - 4) * 290
                                        Else
                                            Me.Top = 1
                                            Me.Left = Val(STR_BREAK(STR_BREAK_STR(Me.Text, "[", "]", 0), 2)) * 290
                                        End If

                                        Me.Height = (My.Computer.Screen.Bounds.Height - 50) / 2
                                    Else
                                        Me.Left = 1
                                        Me.Top = 1
                                        Me.Height = My.Computer.Screen.Bounds.Height - 50
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Me.Left > My.Computer.Screen.Bounds.Width Then Me.Left = My.Computer.Screen.Bounds.Width - 50
                '**************************************************************
            End If
        Catch ex As Exception
            MsgError("InitApp", ex.ToString)
        End Try
    End Sub

    Private Sub ListBox_LOG_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListBox_LOG.MouseDoubleClick
        Try
            Shell("NOTEPAD.EXE " + dir_LOG + "aoProcessor-" + DateToday(8) + ".LOG", AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgError("ListBox_LOG_MouseDoubleClick", ex.ToString)
        End Try
    End Sub

    Private Sub ComboBox1_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseDoubleClick
        QRYselect()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If swQRYtable Then
            QRYrunDT(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, False)
        Else
            QRYrun(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, False)
        End If
    End Sub

    Private Sub DataGridView_QRYs_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView_QRYs.CellMouseDoubleClick
        QRYselect()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Shell("EXPLORER.EXE " + dir_REPORT, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgError("ListBox_LOG_MouseDoubleClick", ex.ToString)
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Try
            Dim tReportName As String = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + DateToday(8) + "_" + rkutils.WhoAmI() + ".XLS"
            If File.Exists(tReportName) Then
                'Shell("NOTEPAD.EXE " + tReportName, AppWinStyle.NormalFocus)
                Shell("EXCEL.EXE " + tReportName, AppWinStyle.NormalFocus)
            Else
                MsgStatus(tReportName + " does NOT exist", True)
            End If
        Catch ex As Exception
            MsgError("ListBox_LOG_MouseDoubleClick", ex.ToString)
        End Try
    End Sub

    Private Sub CheckBox_RUN_Click(sender As Object, e As System.EventArgs) Handles CheckBox_RUN.Click
        If Label_RUNNING.Text = "Running" Then
            Label_RUNNING.Text = "STOPPED"
        Else
            If CheckBox_RUN.Checked Then
                Label_RUNNING.Text = "Ready"
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        QRYcomplete()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text.Trim.Length = 0 Then Exit Sub
            QRYlist(ComboBox1.Text, False)
        Catch ex As Exception
            MsgError("ComboBox1_SelectedIndexChanged", ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView_QRYs_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_QRYs.CellDoubleClick
        QRYread(DataGridView_QRYs.CurrentCellAddress.Y)
    End Sub

    Private Sub QRYcomplete()
        Try
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = ""
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL = "UPDATE RevMD.dbo.query"
                Case Else
                    tSEL = "UPDATE iTeleCollect.dbo.query"
            End Select
            tSEL += " SET scheduler_runtime='" + rkutils.STR_format("TODAY", "ccyy-mm-dd HH:MM:SS") + "'"
            tSEL += " WHERE TUNIQUE='" + Label_QRY_tUnique.Text + "'"
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, tSEL)
            MsgStatus(tSEL, True)
        Catch ex As Exception
            MsgError("QRYload", ex.ToString)
        End Try
    End Sub

    Private Sub QRYcompleteLive()
        Try
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = ""
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL = "UPDATE RevMD.dbo.query"
                Case Else
                    tSEL = "UPDATE iTeleCollect.dbo.query"
            End Select
            tSEL += " SET LIVE_RUN='C'"
            tSEL += " WHERE TUNIQUE='" + Label_QRY_tUnique.Text + "'"
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, tSEL)
            MsgStatus(tSEL, False)
        Catch ex As Exception
            MsgError("QRYload", ex.ToString)
        End Try
    End Sub

    Protected Function CallList_exists(ByVal tCallName As String) As Boolean
        '*************************************************************************************************************
        Label_Unique_Name.Text = rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "TUNIQUE", msSQLConnectionString, msSQLuser, "SELECT TUNIQUE FROM TeleServer.dbo.calllists WHERE CALLLIST='" + tCallName + "'")
        If Val(Label_Unique_Name.Text) > 0 Then Return True
        Return False
    End Function

    Private Sub QRYblank()
        Try
            Label_QRY_tUnique.Text = ""
            Label_QRY_Name.Text = ""
            Label_QRY_Status.Text = ""
            Label_SetValue.Text = ""
            Label_SetField.Text = ""
            Label_QRY_EmailType.Text = ""
            Label_QRY_EMail.Text = ""
            Label_QRY_EmailMessage.Text = ""
            Label_QRY.Text = ""
            DataGridView_QRYs.Visible = False
            DataGridView_QRYoutput.Visible = False
        Catch ex As Exception
            MsgError("QRYblank", ex.ToString)
        End Try
    End Sub

    Private Sub QRYlive()
        Try
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = ""
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL = "SELECT * FROM RevMD.dbo.query"
                Case Else
                    tSEL = "SELECT * FROM iTeleCollect.dbo.query"
            End Select
            tSEL += " WHERE LIVE_RUN='Y'"
            tSEL += " ORDER BY NAME"
            MsgStatus(tSEL, False)
            Me.DataGridView_QRYs.Visible = SQL_READ_DATAGRID(DataGridView_QRYs, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL)
            '*******************************************************************
            '* 2012-10-12 RFK: If only the header row then HIDE the DataGridView
            If DataGridView_QRYs.RowCount = 1 Then Me.DataGridView_QRYs.Visible = False
            Label_NumberQueries0.Text = Me.DataGridView_QRYs.RowCount - 1.ToString.Trim
            '*******************************************************************
            '* 2013-02-27 RFK:
            If Val(Label_NumberQueries0.Text) > 0 Then
                MsgStatus("QRYlive:" + Label_NumberQueries0.Text, True)
            End If
            ResizeApp()
        Catch ex As Exception
            MsgError("QRYlive", ex.ToString)
        End Try
    End Sub

    Private Sub ScrapeInit()
        Try
            Dim tCOMMAND As String = "", tCOUNT As String = "", tNAME As String = ""
            Dim iRow As Integer = 0, iCount As Integer = 0
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = "SELECT SCRAPETYPE,COUNT(*) AS COUNT"
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL += " FROM RevMD.dbo.commands_scrape"
                Case Else
                    '**********************************************************
                    '* 2015-07-09 RFK: 
                    Exit Sub
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    tSEL += " FROM iTeleCollect.commands_scrape"
            End Select
            tSEL += " WHERE TYPE='S' AND LEN(SCRAPETYPE)> 0 AND SCRAPETYPE IS NOT NULL GROUP BY SCRAPETYPE ORDER BY SCRAPETYPE"
            If SQL_READ_DATAGRID(DataGridView1, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL) Then
                '**************************************
                For i1 = 0 To Me.DataGridView1.RowCount - 1
                    tNAME = rkutils.DataGridView_ValueByColumnName(DataGridView1, "SCRAPETYPE", i1).Trim
                    tCOUNT = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1).Trim
                    If tNAME.Length > 0 And Val(tCOUNT) > 1 Then
                        '**************************************
                        iCount += 1
                        iRow = Listbox_Row(Me.ListBox_Scrape, tNAME, False)
                        If iRow >= 0 Then
                            Me.ListBox_Scrape.Items.Item(iRow) = tNAME + " " + tCOUNT
                        Else
                            Me.ListBox_Scrape.Items.Add(tNAME + " " + tCOUNT)
                        End If
                    End If
                Next
                Label_NumberScrape0.Text = iCount.ToString.Trim
                If iCount = 0 And Me.ListBox_Scrape.Items.Count > 0 Then
                    Me.ListBox_Scrape.Items.Clear()
                End If
            Else
                Label_NumberScrape0.Text = "0"
                Label_NumberScrape0.Text = "0"
                Me.ListBox_Scrape.Items.Clear()
            End If
            '**************************************
        Catch ex As Exception
            MsgError("ScrapeInit", ex.ToString)
        End Try
    End Sub

    Private Sub WorkQUE_ADD(ByVal sWorkQue As String, ByVal sDescription As String)
        '**********************************************************************
        '* 2011-11-01 RFK: 
        Try
            Dim SQLcommandstring As String = ""
            Dim sReportString As String = ""
            Dim sBalance As String = ""

            sWorkQue = sWorkQue.Replace(" ", "_")
            MsgStatus("WorkQUE_ADD/" + sWorkQue, True)
            '******************************************************************
            '* 2015-07-09 RFK: 
            Select Case sSITE
                Case "AnnuityOne"
                    SQLcommandstring = "DELETE FROM RevMD.dbo.que"
                Case Else
                    SQLcommandstring = "DELETE FROM iTeleCollect.dbo.que"
            End Select
            SQLcommandstring += " WHERE QUE='" + STR_LEFT(sWorkQue.Replace(" ", "_"), 20) + "'"
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
            '******************************************************************
            Label_QueriesAccounts.Text = "0"
            For i1 = 0 To Me.DataGridView_QRYoutput.Rows.Count - 1
                System.Windows.Forms.Application.DoEvents()
                '**************************************************************
                '* 2014-08-06 RFK: Only add if a valid LOCX
                If Val(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RALOCX", i1).Trim) > 0 Then
                    '**********************************************************
                    '* 2015-07-09 RFK: 
                    Select Case sSITE
                        Case "AnnuityOne"
                            SQLcommandstring += "INSERT INTO RevMD.dbo.que"
                        Case Else
                            SQLcommandstring += "INSERT INTO iTeleCollect.dbo.que"
                    End Select
                    SQLcommandstring += " (tunique"
                    SQLcommandstring += ", type"
                    SQLcommandstring += ", que"
                    SQLcommandstring += ", client"
                    SQLcommandstring += ", locx"
                    SQLcommandstring += ", balance"
                    SQLcommandstring += ", insert_date, insert_by"
                    SQLcommandstring += ", modified_date, modified_by)"
                    SQLcommandstring += " values('" + System.Guid.NewGuid.ToString() + "'"
                    SQLcommandstring += ", 'W'"
                    SQLcommandstring += ", '" + STR_LEFT(sWorkQue, 20) + "'"
                    SQLcommandstring += ", '" + STR_LEFT(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RACL#", i1).Trim, 20) + "'"
                    SQLcommandstring += ", '" + STR_LEFT(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RALOCX", i1).Trim, 20) + "'"
                    '******************************
                    sBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RABALD", i1).Trim
                    If Val(sBalance) = 0 Then
                        sBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", i1).Trim
                    End If
                    SQLcommandstring += ", " + STR_format(sBalance, "0") + ""
                    SQLcommandstring += ", '" + Date.Now + "'"
                    SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
                    SQLcommandstring += ", '" + Date.Now + "'"
                    SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
                    SQLcommandstring += ")" + vbCr
                    If i1 Mod 100 = 0 And SQLcommandstring.Length > 0 Then
                        rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
                        SQLcommandstring = ""
                    End If
                End If
                Label_QueriesAccounts.Text = Trim(Str(Val(Label_QueriesAccounts.Text) - 1))
            Next
            MsgStatus(SQLcommandstring, False)
            If SQLcommandstring.Length > 0 Then rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
            '********************************************************
            If Label_QRY_EMail.Text.Contains("@") Then
                sReportString = "<html><body>"
                sReportString += sWorkQue
                sReportString += "<br>"
                sReportString += sDescription
                sReportString += "<br><br>"
                sReportString += "Added " + Str(DataGridView_QRYoutput.Rows.Count - 1) + " to " + STR_LEFT(sWorkQue, 20)
                sReportString += "<br><br>"
                sReportString += "</body></html>"
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "aoProcessor", Label_QRY_EMail.Text, Label_QRY_EMail.Text, "", "WorkQUE " + sWorkQue + " created", "", sReportString, "")
                'MsgStatus(sReportString, True)
            Else
                MsgStatus("No Email", True)
            End If
            '********************************************************
            MsgStatus("Added " + Str(DataGridView_QRYoutput.Rows.Count - 1) + " to WorkQUE:" + STR_LEFT(sWorkQue, 20), True)
            Label_QueriesAccounts.Text = "0"
        Catch ex As Exception
            MsgStatus("WorkQUE_ADD:" + ex.ToString, True)
        End Try
    End Sub

    Private Sub ResizeApp()
        Try
            '******************************************************************
            Label_TIME.Top = 10
            Label_TIME.Left = Me.Width - 75     'Label_TIME.Width - 20
            '******************************************************************
            Label_QRY_tUnique.Left = Label_QRY_Name.Left + Label_QRY_Name.Width + 10
            Label_QRY_EmailType.Left = Label_QRY_Status.Left + Label_QRY_Status.Width + 10
            Label_QRY_EMail.Left = Label_QRY_EmailType.Left + Label_QRY_EmailType.Width + 10
            Label_QRY_EmailMessage.Left = Label_QRY_EMail.Left + Label_QRY_EMail.Width + 10
            '******************************************************************
            '* 2015-04-22 RFK:
            Me.ListBox_LOG.Height = 100
            Me.ListBox_LOG.Top = Me.Height - Me.ListBox_LOG.Height - 40
            Me.ListBox_LOG.Width = Me.Width - 25
            '******************************************************************
            Dim iHeight = Me.ListBox_LOG.Top - Me.TabControl1.Top
            'MsgStatus(iHeight.ToString, True)
            '******************************************************************
            Me.TabControl1.Width = Me.Width - 25
            Me.TabControl1.Height = iHeight / 2
            '******************************************************************
            Me.TextBox_testQRY.Width = Me.TabControl1.Width - 10
            Me.TextBox_testQRY.Height = Me.TabControl1.Height - 25
            Me.Button_QRYrun.Left = Me.TabControl1.Width - Me.Button_QRYrun.Width - 25
            Me.Button_Excel.Left = Me.TabControl1.Width - Me.Button_Excel.Width - 25
            '******************************************************************
            '* 2015-04-22 RFK:
            Me.DataGridView_QRYs.Width = Me.TabControl1.Width - 25
            Me.DataGridView_QRYs.Height = Me.TabControl1.Height - 50
            '******************************************************************
            '* 2015-04-22 RFK:
            If Me.DataGridView_QRYoutput.RowCount - 1 > 0 Then
                Me.DataGridView_QRYoutput.Height = iHeight / 2
                Me.DataGridView_QRYoutput.Width = Me.Width - 25
                Me.DataGridView_QRYoutput.Top = Me.ListBox_LOG.Top - Me.DataGridView_QRYoutput.Height
            Else
                Me.ListBox_LOG.Top = Me.TabControl1.Top + Me.TabControl1.Height
                Me.ListBox_LOG.Height = Me.Height - Me.ListBox_LOG.Top - 25
            End If
        Catch ex As Exception
            MsgError("ResizeApp", ex.ToString)
        End Try
    End Sub

    Private Sub Button_QRYrun_Click(sender As System.Object, e As System.EventArgs) Handles Button_QRYrun.Click
        Label_QRY.Text = Me.TextBox_testQRY.Text
        Label_QRY_Name.Text = "TEST"
        QRYselect()
    End Sub

    Private Sub UpdateCommand(ByVal tUnique As String)
        Try
            If tUnique.Length > 0 Then
                '**********************************************************
                '* 2015-07-09 RFK: 
                Dim SQLcommandstring As String = ""
                Select Case sSITE
                    Case "AnnuityOne"
                        SQLcommandstring = "UPDATE RevMD.dbo.commands"
                    Case Else
                        SQLcommandstring = "UPDATE iTeleCollect.dbo.commands"
                End Select
                SQLcommandstring += " SET COMPLETED_DATE='" + DateToday(20) + "' WHERE tUNIQUE='" + tUnique + "'"
                rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
            End If
        Catch ex As Exception
            MsgError("UpdateCommand", ex.ToString)
        End Try
    End Sub

    Private Sub UpdateCommands(ByVal tUniques As String)
        Try
            If tUniques.Length > 0 Then
                '**************************************************************
                '* 2015-07-09 RFK: 
                Dim SQLcommandstring As String = ""
                Select Case sSITE
                    Case "AnnuityOne"
                        SQLcommandstring = "UPDATE RevMD.dbo.commands"
                    Case Else
                        SQLcommandstring = "UPDATE iTeleCollect.dbo.commands"
                End Select
                SQLcommandstring += " SET COMPLETED_DATE='" + DateToday(20) + "' WHERE tUNIQUE IN (" + tUniques + ")"
                rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
                MsgStatus(tUniques, True)
            End If
        Catch ex As Exception
            MsgError("UpdateCommands", ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button_Excel.Click
        Try
            ExportToCSV("c:\temp\temp.csv", ",", DataGridView_QRYoutput)
        Catch ex As Exception
            MsgError("Button3_Click", ex.ToString)
        End Try
    End Sub

    Private Sub QRYlist(ByVal tTYPE As String, ByVal tLimit As Boolean)
        Try
            QRYblank()
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = ""
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL = "SELECT * FROM RevMD.dbo.query"
                Case Else
                    tSEL = "SELECT * FROM iTeleCollect.dbo.query"
            End Select
            '******************************************************************
            Select Case tTYPE
                Case "*"
                    tSEL += " WHERE TYPE IN('Command', 'Report', 'WorkQUE'"
                    If CheckBox_CallList.Checked Then
                        tSEL += ",'CallList'"
                    End If
                    tSEL += ")"
                Case Else
                    tSEL += " WHERE TYPE='" + tTYPE + "'"
            End Select
            '******************************************************************
            '* 2016-01-27 RFK: 
            'tSEL += " AND WHO NOT IN('Deleted', 'History')"
            tSEL += " AND WHO NOT IN('Deleted', 'History')"
            '******************************************************************
            '* 2017-10-16 RFK: 
            tSEL += " AND LEFT(NAME, 1) <> '_'"
            '******************************************************************
            tSEL += " AND ("
            tSEL += " (scheduler IN ('DAILY'"
            tSEL += ",'" + Now.Day.ToString + "'"
            Select Case Now.DayOfWeek
                Case 0
                    tSEL += ",'SUNDAY'"
                Case 1
                    tSEL += ",'MONDAY','MON-FRI','MON-SAT'"
                Case 2
                    tSEL += ",'TUESDAY','MON-FRI','MON-SAT','TUE-SAT'"
                Case 3
                    tSEL += ",'WEDNESDAY','MON-FRI','MON-SAT','TUE-SAT'"
                Case 4
                    tSEL += ",'THURSDAY','MON-FRI','MON-SAT','TUE-SAT'"
                Case 5
                    tSEL += ",'FRIDAY','MON-FRI','MON-SAT','TUE-SAT'"
                Case 6
                    tSEL += ",'SATURDAY','MON-SAT','TUE-SAT'"
                Case Else
                    MsgStatus("Day of week=" + Now.DayOfWeek, True)
            End Select
            tSEL += ")"
            If tLimit Then
                tSEL += " AND '" + rkutils.STR_format("TODAY", "HH:MM") + "' >= scheduler_time"
                tSEL += " AND (scheduler_runtime IS NULL"
                tSEL += " OR CONVERT(varchar(16), scheduler_runtime, 21)"
                tSEL += " < '" + rkutils.STR_format("TODAY", "ccyy-mm-dd") + " '+scheduler_time"
                tSEL += ")"
            End If
            tSEL += ")"
            '*************************************************
            '* 2015-01-05 RFK:
            '* 2015-03-30 RFK:
            tSEL += " OR ("
            tSEL += "(scheduler IN ('HOURLY'"
            Select Case Now.DayOfWeek
                Case DayOfWeek.Saturday
                    tSEL += ",'HOURLY_MS'"
                Case DayOfWeek.Sunday
                    'Nothing
                Case Else
                    tSEL += ",'HOURLY_MF'"
                    tSEL += ",'HOURLY_MS'"
            End Select
            tSEL += ") AND "
            tSEL += "('" + rkutils.STR_format("TODAY", "HH") + "' >= left(scheduler_time,2)) AND"
            tSEL += "('" + rkutils.STR_format("TODAY", "HH") + "' <= left(scheduler_timeuntil,2)) AND"
            tSEL += "("
            tSEL += "(scheduler_runtime IS NULL) OR  (left(right(CONVERT(varchar(16), scheduler_runtime, 21),5),2) <> '" + rkutils.STR_format("TODAY", "HH") + "')))"
            tSEL += " )"
            '*************************************************
            tSEL += " )"
            tSEL += " ORDER BY NAME"
            MsgStatus(tSEL, False)
            DataGridView_QRYs.Visible = SQL_READ_DATAGRID(DataGridView_QRYs, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL)
            '*******************************************************************
            '* 2012-10-12 RFK: If only the header row then HIDE the DataGridView
            If DataGridView_QRYs.RowCount = 1 Then Me.DataGridView_QRYs.Visible = False
            Label_NumberQueries0.Text = Me.DataGridView_QRYs.RowCount - 1.ToString.Trim
            ResizeApp()
        Catch ex As Exception
            MsgError("QRYlist", ex.ToString)
        End Try
    End Sub

    Protected Sub CallList_Error(ByVal tCallList As String, ByVal tDescription As String, ByVal tError As String)
        Try
            Dim SQLcommandstring As String = ""
            SQLcommandstring = "INSERT INTO TeleServer.dbo.CallListErrors"
            SQLcommandstring += "(callList, error, description"
            SQLcommandstring += ", timeinserted, modified_date, modified_by)"
            SQLcommandstring += "values('" + rkutils.STR_LEFT(tCallList, 20) + "'"
            SQLcommandstring += ", '" + tError + "'"
            SQLcommandstring += ", '" + tDescription + "'"
            SQLcommandstring += ", '" + Date.Now + "'"
            SQLcommandstring += ", '" + Date.Now + "'"
            SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
            SQLcommandstring += ")" + vbCr
            rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
        Catch ex As Exception
            MsgStatus("CallList_Error:" + ex.ToString, True)
        End Try
    End Sub

    Private Sub QRYinit()
        Try
            Dim swLIVE As Boolean = False
            If Label_RUNNING.Text = "Ready" Then
                If swTEST Then MsgStatus("QRYinit", True)
                '*****************************************
                '* 2012-01-01 RFK:
                QRYlist("*", True)
                '**************************************************************
                '* 2013-02-27 RFK:
                If DataGridView_QRYs.RowCount - 1 < 1 Then
                    QRYlive()
                    swLIVE = True
                End If
                '**************************************************************
                If Me.DataGridView_QRYs.RowCount - 1 > 0 Then
                    Dim iQRYLoop As Integer = 0
                    Dim sQRYtype As String = "", sQRYwho As String = ""
                    iLastQryMinute = Now.TimeOfDay.Minutes  'Reset to this minute
                    Do While (Label_RUNNING.Text = "Ready" And iLastQryMinute = Now.TimeOfDay.Minutes And iQRYLoop < Me.DataGridView_QRYs.RowCount - 1)
                        '******************************************************
                        '* 2017-01-10 RFK: REPORTS run in a seperate .EXE space
                        '* 2017-02-15 RFK: corrected WHO parameter
                        sQRYtype = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", iQRYLoop).Trim
                        sQRYwho = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "WHO", iQRYLoop).Trim
                        QRYread(iQRYLoop)
                        '******************************************************
                        MsgStatus("QRY:" + sQRYtype + " " + sQRYwho + " " + Label_QRY_Name.Text, True)
                        '******************************************************
                        Select Case sQRYtype
                            Case "CallList"
                                '**********************************************
                                '* 2014-07-21 RFK:
                                '* 2017-07-21 RFK: CallLists run in a seperate .EXE space
                                QRYrun_shell(sQRYtype, sQRYwho, Label_QRY_Name.Text, swLIVE)
                                '**********************************************
                                If swLIVE Then
                                    QRYcompleteLive()
                                Else
                                    QRYcomplete()
                                End If
                                '**********************************************
                            Case "Command"
                                COMMANDrun()
                                '**********************************************
                                If swLIVE Then
                                    QRYcompleteLive()
                                Else
                                    QRYcomplete()
                                End If
                                '**********************************************
                            Case "Report"
                                '**********************************************
                                '* 2017-01-10 RFK: REPORTS run in a seperate .EXE space
                                '* 2017-02-15 RFK: corrected WHO parameter
                                QRYrun_shell(sQRYtype, sQRYwho, Label_QRY_Name.Text, swLIVE)
                                '**********************************************
                                If swLIVE Then
                                    QRYcompleteLive()
                                Else
                                    QRYcomplete()
                                End If
                                '**********************************************
                            Case "WorkQUE"
                                QRYselect()
                                If rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", iQRYLoop).Trim.Length > 0 Then
                                    Label_RUNNING.Text = "Running"
                                    WorkQUE_ADD(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", iQRYLoop).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "DESCRIPTION", iQRYLoop).Trim)
                                End If
                                If Label_RUNNING.Text = "Running" Then Label_RUNNING.Text = "Ready"
                                '**********************************************
                                If swLIVE Then
                                    QRYcompleteLive()
                                Else
                                    QRYcomplete()
                                End If
                                '**********************************************
                        End Select
                        '******************************************************
                        iQRYLoop += 1
                    Loop
                End If
                '*****************************************
            End If
        Catch ex As Exception
            MsgError("QRYinit", ex.ToString)
        End Try
    End Sub

    Protected Sub CheckStatusNext_shell()
        '**********************************************************************
        '* 2017-01-11 RFK:
        Try
            Dim swFOUND As Boolean = False
            For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                '**************************************************************
                '* 2015-04-28 RFK:
                If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[Check Status Next]") Then
                    swFOUND = True
                    Exit For
                End If
            Next
            '******************************************************************
            '* 2017-01-11 RFK:
            If swFOUND = False Then rkutils.ShellIT(Application.ExecutablePath, "/CHECKSTATUSNEXT", 1)
        Catch ex As Exception
            MsgError("CheckStatusNext_shell", ex.ToString)
        End Try
    End Sub

    Protected Sub CheckStatusNext()
        Try
            Dim sSQL As String = ""
            sSQL += "SELECT RALOCX AS LOCX, RACL#"
            sSQL += ", digits(RARNMO)||'/'||digits(RARNDY)||'/'||digits(RARNYR) AS NextReviewDate"
            sSQL += ", A.RAPSTA, A.RARSTA, S.STNXTS, S.STDAYS"
            sSQL += " FROM ROIDATA.RACCTP A"
            sSQL += " JOIN ROIDATA.HCLNTP H ON A.RACL# = H.HCCL#"
            sSQL += " JOIN ROIDATA.STATP S ON A.RAMTTP = S.STMTTP AND A.RARSTA=S.STSTAT"
            sSQL += rkutils.WhereAnd(sSQL, "H.HCACTV <> 'N'")
            sSQL += rkutils.WhereAnd(sSQL, "A.RACLOS <> 'C'")
            sSQL += rkutils.WhereAnd(sSQL, "A.RATOB IN ('5')")  '* 2017-05-04 RFK: Only Type Of Business (5=Insurance, 3=SelfPay, 222=BadDebt)
            sSQL += rkutils.WhereAnd(sSQL, "DAYS(digits(A.RARNMO)||'/'||digits(A.RARNDY)||'/'||digits(A.RARNYR)) <= days(CURRENT DATE)")
            sSQL += rkutils.WhereAnd(sSQL, "S.STSTAT<>S.STNXTS")
            sSQL += rkutils.WhereAnd(sSQL, "TRIM(S.STNXTS) NOT IN ('','---')")   '* 2018-03-28 RFK: Query or Logic moves these, NOT BLANK
            MsgStatus(sSQL, True)
            If swQRYtable Then
                DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
                'Should convert to DTABLE
                'rkutils.SQL_READ_DATATABLE(DTqry, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, Label_QRY.Text)
            Else
                DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, sSQL)
            End If
            If DataGridView_QRYoutput.RowCount > 1 Then
                Dim sHTML As String = ""
                Dim sFileName As String = dir_REPORT + "StatusCodeNext\" + rkutils.STR_format("TODAY", "ccyymmddHHMMSS") + "_StatusToNext.XLS"
                rkutils.ExportToCSV(sFileName, vbTab, DataGridView_QRYoutput)
                '**************************************************************
                If DataGridView_QRYoutput.RowCount <= 1000 Then
                    sHTML += "<html><head></head>"
                    sHTML += rkutils.DataGridview_ToHTMLtable(DataGridView_QRYoutput, "")
                    sHTML += "</html>"
                End If
                '**************************************************************
                '* 2017-01-11 RFK: Email them
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "eProcess_DoNotReply@AnnuityHealth.com", "eProcess CheckStatusNext", "Ryan.Kiechle@AnnuityHealth.com", "IT", "", "Next Status Codes", "These accounts were (in process of) statused to the Next Status Code", sHTML, sFileName)
                '**************************************************************
                '* 2017-01-11 RFK: Status Them
                For intRow As Integer = 0 To DataGridView_QRYoutput.RowCount - 1
                    rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LOCX", intRow), rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "STNXTS", intRow), "", "", "")
                Next
            End If
            MsgStatus("CheckStatusNext checked [" + DataGridView_QRYoutput.RowCount.ToString + "]", True)
        Catch ex As Exception
            MsgError("CheckStatusNext", ex.ToString)
        End Try
    End Sub

    Protected Sub QRYrun_shell(ByVal sQRYtype As String, ByVal sQRYwho As String, ByVal sQRYname As String, ByVal bLive As Boolean)
        Try
            Dim swFOUND As Boolean = False
            For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                '**************************************************
                '* 2015-04-28 RFK:
                If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRYRUN " + sQRYtype + " " + sQRYwho + " " + sQRYname + "]") Then
                    swFOUND = True
                    Exit For
                End If
            Next
            '******************************************************
            '* 2015-04-28 RFK:
            If swFOUND Then
                MsgStatus(sQRYtype + " " + sQRYname + " already running", True)
            Else
                Dim sRunFile As String = Application.ExecutablePath
                If IS_File(sRunFile) Then
                    MsgStatus(sRunFile + " /QRYRUN " + sQRYtype + " " + sQRYwho + " " + sQRYname, True)
                    If bLive Then
                        rkutils.ShellIT(sRunFile, "/QRYRUNLIVE " + sQRYtype + " " + sQRYwho + " " + sQRYname, 1)
                    Else
                        rkutils.ShellIT(sRunFile, "/QRYRUN " + sQRYtype + " " + sQRYwho + " " + sQRYname, 1)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgError("QRYrun_shell", ex.ToString)
        End Try
    End Sub

    Private Sub Button_TEST_Click(sender As System.Object, e As System.EventArgs) Handles Button_TEST.Click
        'Dim tSEL As String = "SELECT * FROM RevMD.dbo.query WHERE NAME='TEST'"
        Dim tSEL As String = "SELECT * FROM RevMD.dbo.query WHERE NAME='INS Ingalls Automated' AND TYPE='CallList'"
        MsgStatus(tSEL, True)
        Me.DataGridView_QRYs.Visible = SQL_READ_DATAGRID(DataGridView_QRYs, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL)

        Label_QRY_Name.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", 0)
        Dim sQRYtype As String = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim
        Dim sQRYwho As String = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "WHO", 0).Trim
        '******************************************************
        MsgStatus("QRY:" + sQRYtype + " " + sQRYwho + " " + Label_QRY_Name.Text, True)
        Select Case sQRYtype
            Case "CallList"
                QRYread(0)
                QRYselect()
                CallList_ADD(rkutils.STR_TRIM(rkutils.STR_BREAK(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_NAME", 0).Trim, 1) + "-" + rkutils.STR_format("TODAY", "mmdd"), 10), rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_GROUP", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_DIALER", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_APPTYPE", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_RATIO", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_START", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_STOP", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_INSERT", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_ACCOUNTNUMBER", 0).Trim, rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "CLST_CLIENTOUTPUT", 0).Trim)
                '**********************************************
                '* 2014-07-21 RFK:
                '* 2017-07-21 RFK: CallLists run in a seperate .EXE space
                'QRYrun_shell(sQRYtype, sQRYwho, Label_QRY_Name.Text, False)
            Case "Report"
                'QRYread(0)
                'QRYselect()
                'Create Reports
                'If swQRYtable Then
                '    QRYrunDT(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, False)
                'Else
                '    QRYrun(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", 0).Trim, False)
                'End If
        End Select
        '**********************************************
        'If swLIVE Then
        '    QRYcompleteLive()
        'Else
        QRYcomplete()
        'End If
        '**********************************************
    End Sub

    Private Sub ProcessInit()
        Try
            Label_NumberCommands0.Text = "0".PadLeft(10)
            Label_NumberEMail0.Text = "0".PadLeft(10)
            Label_NumberEMailSecure0.Text = "0".PadLeft(10)
            Label_NumberNote0.Text = "0".PadLeft(10)
            Label_NumberStatus0.Text = "0".PadLeft(10)
            Label_NumberQueries0.Text = "0".PadLeft(10)
            '******************************************************************
            '* 2015-07-09 RFK: 
            Dim tSEL As String = "SELECT COMMAND,COUNT(*) AS COUNT"
            Select Case sSITE
                Case "AnnuityOne"
                    tSEL += " FROM RevMD.dbo.commands"
                Case Else
                    tSEL += " FROM iTeleCollect.dbo.commands"
            End Select
            tSEL += " WHERE COMPLETED_DATE IS NULL GROUP BY COMMAND"
            If SQL_READ_DATAGRID(DataGridView1, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL) Then
                Dim tCOMMAND As String = ""
                Label_NumberCommands0.Text = Str(DataGridView1.RowCount - 1)
                'Me.DataGridView1.Visible = True
                For i1 = 0 To Me.DataGridView1.RowCount - 1
                    tCOMMAND = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COMMAND", i1)
                    Select Case tCOMMAND
                        Case "EMAIL"
                            Label_NumberEMail0.Text = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1).PadLeft(10)
                        Case "EMAILSEC", "EMAILAR"
                            Label_NumberEMailSecure0.Text = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1).PadLeft(10)
                        Case "STATUS"
                            Label_NumberStatus0.Text = Val(rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1)).ToString.PadLeft(10)
                        Case "NOTE"
                            Label_NumberNote0.Text = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1).PadLeft(10)
                        Case "TRACK"
                            Label_NumberTrack0.Text = rkutils.DataGridView_ValueByColumnName(DataGridView1, "COUNT", i1).PadLeft(10)
                    End Select
                Next
            End If
            '******************************************************************
            '* 2021-01-25 RFK: 
            If Val(Label_NumberStatus0.Text) > 0 Then
                tSEL = "SELECT STATUS,COUNT(*) AS COUNT"
                Select Case sSITE
                    Case "AnnuityOne"
                        tSEL += " FROM RevMD.dbo.commands"
                    Case Else
                        tSEL += " FROM iTeleCollect.dbo.commands"
                End Select
                tSEL += " WHERE COMPLETED_DATE IS NULL"
                tSEL += " GROUP BY STATUS"
                DataGridView_Statuses.Visible = SQL_READ_DATAGRID(DataGridView_Statuses, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSEL)
            Else
                DataGridView_Statuses.Visible = False
            End If
            '******************************************************************
            '* 2015-04-28 RFK: 
            If Text.Contains("[STATUS") Then
                If CheckBox_RUN.Checked = False Then ExitApp()
                '**************************************************************
                '* 2015-04-28 RFK: Exit This One
                If Me.Text.Contains("[STATUS 2]") And Val(Label_NumberStatus0.Text) < 500 Then ExitApp()
                If Me.Text.Contains("[STATUS 3]") And Val(Label_NumberStatus0.Text) < 1000 Then ExitApp()
                If Me.Text.Contains("[STATUS 4]") And Val(Label_NumberStatus0.Text) < 1500 Then ExitApp()
                If Me.Text.Contains("[STATUS 5]") And Val(Label_NumberStatus0.Text) < 2000 Then ExitApp()
                If Me.Text.Contains("[STATUS 6]") And Val(Label_NumberStatus0.Text) < 2500 Then ExitApp()
                If Me.Text.Contains("[STATUS 7]") And Val(Label_NumberStatus0.Text) < 3000 Then ExitApp()
                If Me.Text.Contains("[STATUS 8]") And Val(Label_NumberStatus0.Text) < 3500 Then ExitApp()
                If Me.Text.Contains("[STATUS 9]") And Val(Label_NumberStatus0.Text) < 4000 Then ExitApp()
                '**************************************************************
                '* 2018-11-14 RFK: Exit This One
                If Me.Text.Contains("[STATUS 10]") And Val(Label_NumberStatus0.Text) < 4500 Then ExitApp()
                If Me.Text.Contains("[STATUS 11]") And Val(Label_NumberStatus0.Text) < 5000 Then ExitApp()
                If Me.Text.Contains("[STATUS 12]") And Val(Label_NumberStatus0.Text) < 5500 Then ExitApp()
                If Me.Text.Contains("[STATUS 13]") And Val(Label_NumberStatus0.Text) < 6000 Then ExitApp()
                If Me.Text.Contains("[STATUS 14]") And Val(Label_NumberStatus0.Text) < 6500 Then ExitApp()
                If Me.Text.Contains("[STATUS 15]") And Val(Label_NumberStatus0.Text) < 7000 Then ExitApp()
                If Me.Text.Contains("[STATUS 16]") And Val(Label_NumberStatus0.Text) < 7500 Then ExitApp()
                If Me.Text.Contains("[STATUS 17]") And Val(Label_NumberStatus0.Text) < 8000 Then ExitApp()
                If Me.Text.Contains("[STATUS 18]") And Val(Label_NumberStatus0.Text) < 8500 Then ExitApp()
                If Me.Text.Contains("[STATUS 19]") And Val(Label_NumberStatus0.Text) < 9000 Then ExitApp()
            End If
            '******************************************************************
            '* 2015-09-28 RFK: 
            If Me.Text.Contains("[NOTE") Then
                '**************************************************************
                '* 2015-04-28 RFK: Exit This One
                If Me.Text.Contains("[NOTE 2]") And Val(Label_NumberNote0.Text) < 500 Then ExitApp()
                If Me.Text.Contains("[NOTE 3]") And Val(Label_NumberNote0.Text) < 1000 Then ExitApp()
                If Me.Text.Contains("[NOTE 4]") And Val(Label_NumberNote0.Text) < 1500 Then ExitApp()
                If Me.Text.Contains("[NOTE 5]") And Val(Label_NumberNote0.Text) < 2000 Then ExitApp()
                If Me.Text.Contains("[NOTE 6]") And Val(Label_NumberNote0.Text) < 2500 Then ExitApp()
                If Me.Text.Contains("[NOTE 7]") And Val(Label_NumberNote0.Text) < 3000 Then ExitApp()
                If Me.Text.Contains("[NOTE 8]") And Val(Label_NumberNote0.Text) < 3500 Then ExitApp()
                If Me.Text.Contains("[NOTE 9]") And Val(Label_NumberNote0.Text) < 4000 Then ExitApp()
            End If
            '******************************************************************
            '* 2017-10-06 RFK: 
            If Me.Text.Contains("[TRACK") Then
                '**************************************************************
                '* 2015-04-28 RFK: Exit This One
                If Me.Text.Contains("[TRACK 2]") And Val(Label_NumberTrack0.Text) < 500 Then ExitApp()
                If Me.Text.Contains("[TRACK 3]") And Val(Label_NumberTrack0.Text) < 1000 Then ExitApp()
                If Me.Text.Contains("[TRACK 4]") And Val(Label_NumberTrack0.Text) < 1500 Then ExitApp()
                If Me.Text.Contains("[TRACK 5]") And Val(Label_NumberTrack0.Text) < 2000 Then ExitApp()
                If Me.Text.Contains("[TRACK 6]") And Val(Label_NumberTrack0.Text) < 2500 Then ExitApp()
                If Me.Text.Contains("[TRACK 7]") And Val(Label_NumberTrack0.Text) < 3000 Then ExitApp()
                If Me.Text.Contains("[TRACK 8]") And Val(Label_NumberTrack0.Text) < 3500 Then ExitApp()
                If Me.Text.Contains("[TRACK 9]") And Val(Label_NumberTrack0.Text) < 4000 Then ExitApp()
            End If
            '******************************************************************
            '* 2015-10-26 RFK: 
            If Me.Text.Contains("[QRY") Then
                '**************************************************************
                '* 2015-04-28 RFK: Exit This One
                If Me.Text.Contains("[QRY 2]") And Val(Label_NumberQueries0.Text) < 50 Then ExitApp()
                If Me.Text.Contains("[QRY 3]") And Val(Label_NumberQueries0.Text) < 100 Then ExitApp()
                If Me.Text.Contains("[QRY 4]") And Val(Label_NumberQueries0.Text) < 150 Then ExitApp()
                If Me.Text.Contains("[QRY 5]") And Val(Label_NumberQueries0.Text) < 200 Then ExitApp()
                If Me.Text.Contains("[QRY 6]") And Val(Label_NumberQueries0.Text) < 250 Then ExitApp()
                If Me.Text.Contains("[QRY 7]") And Val(Label_NumberQueries0.Text) < 300 Then ExitApp()
                If Me.Text.Contains("[QRY 8]") And Val(Label_NumberQueries0.Text) < 350 Then ExitApp()
                If Me.Text.Contains("[QRY 9]") And Val(Label_NumberQueries0.Text) < 400 Then ExitApp()
            End If
            '******************************************************************
            '* 2015-04-28 RFK: look for other eProcessor [STATUS
            If Me.Text.Contains("[1]") Then
                If Val(Label_NumberStatus0.Text) > 500 Then
                    Dim sRunFile As String = "c:\tele\aoProcess.EXE"
                    If IS_File(sRunFile) Then
                        Dim swFound As Boolean = False, swFound2 As Boolean = False, swFound3 As Boolean = False, swFound4 As Boolean = False, swFound5 As Boolean = False
                        Dim swFound6 As Boolean = False, swFound7 As Boolean = False, swFound8 As Boolean = False, swFound9 As Boolean = False
                        Dim swFound10 As Boolean = False, swFound11 As Boolean = False, swFound12 As Boolean = False, swFound13 As Boolean = False
                        Dim swFound14 As Boolean = False, swFound15 As Boolean = False, swFound16 As Boolean = False, swFound17 As Boolean = False
                        Dim swFound18 As Boolean = False, swFound19 As Boolean = False
                        For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                            '**************************************************
                            '* 2015-04-28 RFK:
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 2]") Then swFound2 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 3]") Then swFound3 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 4]") Then swFound4 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 5]") Then swFound5 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 6]") Then swFound6 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 7]") Then swFound7 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 8]") Then swFound8 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 9]") Then swFound9 = True
                            '**************************************************
                            '* 2018-11-14 RFK:
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 10]") Then swFound10 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 11]") Then swFound11 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 12]") Then swFound12 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 13]") Then swFound13 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 14]") Then swFound14 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 15]") Then swFound15 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 16]") Then swFound16 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 17]") Then swFound17 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 18]") Then swFound18 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[STATUS 19]") Then swFound19 = True
                        Next
                        '******************************************************
                        '* 2015-04-28 RFK:
                        Dim sMSG As String = "aoProcessor started additional eProcessor" + vbCrLf
                        Dim swEmail As Boolean = False
                        If Val(Label_NumberStatus0.Text) > 500 And swFound2 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 2", 1)
                            sMSG += "/STATUS 2" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 1000 And swFound3 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 3", 1)
                            sMSG += "/STATUS 3" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 1500 And swFound4 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 4", 1)
                            sMSG += "/STATUS 4" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 2000 And swFound5 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 5", 1)
                            sMSG += "/STATUS 5" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 2500 And swFound6 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 6", 1)
                            sMSG += "/STATUS 6" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 3000 And swFound7 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 7", 1)
                            sMSG += "/STATUS 7" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 3500 And swFound8 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 8", 1)
                            sMSG += "/STATUS 8" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 4000 And swFound9 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 9", 1)
                            sMSG += "/STATUS 9" + vbCrLf
                            swEmail = True
                        End If
                        '******************************************************
                        '* 2018-11-14 RFK:
                        If Val(Label_NumberStatus0.Text) > 4500 And swFound10 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 10", 1)
                            sMSG += "/STATUS 10" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 5000 And swFound11 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 11", 1)
                            sMSG += "/STATUS 11" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberStatus0.Text) > 5500 And swFound12 = False Then
                            rkutils.ShellIT(sRunFile, "/STATUS 12", 1)
                            sMSG += "/STATUS 12" + vbCrLf
                            swEmail = True
                        End If
                        '******************************************************
                        '* 2018-11-14 RFK: too much for 1 machine CPU is FINE / NETWORK SQL is too much
                        '* 2018-11-14 RFK: look into VM:eProcessor2
                        'If Val(Label_NumberStatus0.Text) > 6000 And swFound13 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 13", 1)
                        '    sMSG += "/STATUS 13" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 6500 And swFound14 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 14", 1)
                        '    sMSG += "/STATUS 14" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 7000 And swFound15 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 15", 1)
                        '    sMSG += "/STATUS 15" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 7500 And swFound16 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 16", 1)
                        '    sMSG += "/STATUS 16" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 8000 And swFound17 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 17", 1)
                        '    sMSG += "/STATUS 17" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 8500 And swFound18 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 18", 1)
                        '    sMSG += "/STATUS 18" + vbCrLf
                        '    swEmail = True
                        'End If
                        'If Val(Label_NumberStatus0.Text) > 9000 And swFound19 = False Then
                        '    rkutils.ShellIT(sRunFile, "/STATUS 19", 1)
                        '    sMSG += "/STATUS 19" + vbCrLf
                        '    swEmail = True
                        'End If
                        '******************************************************
                        If swEmail Then
                            'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply_eProcessor@AnnuityHealth.com", "aoProcessor", "IT@AnnuityHealth.com", "IT", "aoProcessor", "aoProcessor Resources", sMSG, "", "")
                        End If
                    End If
                End If
                '******************************************************************
                '* 2015-09-28 RFK: 
                If Val(Label_NumberNote0.Text) > 500 Then
                    Dim sRunFile As String = "c:\tele\aoProcess.EXE"
                    If IS_File(sRunFile) Then
                        Dim swFound As Boolean = False, swFound2 As Boolean = False, swFound3 As Boolean = False, swFound4 As Boolean = False, swFound5 As Boolean = False
                        Dim swFound6 As Boolean = False, swFound7 As Boolean = False, swFound8 As Boolean = False, swFound9 As Boolean = False
                        For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                            '**************************************************
                            '* 2015-04-28 RFK:
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 2]") Then swFound2 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 3]") Then swFound3 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 4]") Then swFound4 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 5]") Then swFound5 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 6]") Then swFound6 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 7]") Then swFound7 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 8]") Then swFound8 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[NOTE 9]") Then swFound9 = True
                        Next
                        '******************************************************
                        '* 2015-04-28 RFK:
                        Dim sMSG As String = "aoProcessor started additional eProcessor" + vbCrLf
                        Dim swEmail As Boolean = False
                        If Val(Label_NumberNote0.Text) > 500 And swFound2 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 2", 1)
                            sMSG += "/NOTE 2" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 1000 And swFound3 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 3", 1)
                            sMSG += "/NOTE 3" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 1500 And swFound4 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 4", 1)
                            sMSG += "/NOTE 4" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 2000 And swFound5 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 5", 1)
                            sMSG += "/NOTE 5" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 2500 And swFound6 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 6", 1)
                            sMSG += "/NOTE 6" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 3000 And swFound7 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 7", 1)
                            sMSG += "/NOTE 7" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 3500 And swFound8 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 8", 1)
                            sMSG += "/NOTE 8" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberNote0.Text) > 4000 And swFound9 = False Then
                            rkutils.ShellIT(sRunFile, "/NOTE 9", 1)
                            sMSG += "/NOTE 9" + vbCrLf
                            swEmail = True
                        End If
                        If swEmail Then
                            'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply_eProcessor@AnnuityHealth.com", "aoProcessor", "IT@AnnuityHealth.com", "IT", "aoProcessor", "aoProcessor Resources", sMSG, "", "")
                        End If
                    End If
                End If
                '******************************************************************
                '* 2017-10-06 RFK: 
                If Val(Label_NumberTrack0.Text) > 500 Then
                    Dim sRunFile As String = "c:\tele\aoProcess.EXE"
                    If IS_File(sRunFile) Then
                        Dim swFound As Boolean = False, swFound2 As Boolean = False, swFound3 As Boolean = False, swFound4 As Boolean = False, swFound5 As Boolean = False
                        Dim swFound6 As Boolean = False, swFound7 As Boolean = False, swFound8 As Boolean = False, swFound9 As Boolean = False
                        For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                            '**************************************************
                            '* 2015-04-28 RFK:
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 2]") Then swFound2 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 3]") Then swFound3 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 4]") Then swFound4 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 5]") Then swFound5 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 6]") Then swFound6 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 7]") Then swFound7 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 8]") Then swFound8 = True
                            If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[TRACK 9]") Then swFound9 = True
                        Next
                        '******************************************************
                        '* 2015-04-28 RFK:
                        Dim sMSG As String = "aoProcessor started additional eProcessor" + vbCrLf
                        Dim swEmail As Boolean = False
                        If Val(Label_NumberTrack0.Text) > 500 And swFound2 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 2", 1)
                            sMSG += "/TRACK 2" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 1000 And swFound3 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 3", 1)
                            sMSG += "/TRACK 3" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 1500 And swFound4 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 4", 1)
                            sMSG += "/TRACK 4" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 2000 And swFound5 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 5", 1)
                            sMSG += "/TRACK 5" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 2500 And swFound6 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 6", 1)
                            sMSG += "/TRACK 6" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 3000 And swFound7 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 7", 1)
                            sMSG += "/TRACK 7" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 3500 And swFound8 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 8", 1)
                            sMSG += "/TRACK 8" + vbCrLf
                            swEmail = True
                        End If
                        If Val(Label_NumberTrack0.Text) > 4000 And swFound9 = False Then
                            rkutils.ShellIT(sRunFile, "/TRACK 9", 1)
                            sMSG += "/TRACK 9" + vbCrLf
                            swEmail = True
                        End If
                        If swEmail Then
                            'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply_eProcessor@AnnuityHealth.com", "aoProcessor", "IT@AnnuityHealth.com", "IT", "aoProcessor", "aoProcessor Resources", sMSG, "", "")
                        End If
                    End If
                End If
            End If
            '******************************************************************
            '* 2015-10-26 RFK: 
            If Val(Label_NumberNote0.Text) > 500 Then
                Dim sRunFile As String = "c:\tele\aoProcess.EXE"
                If IS_File(sRunFile) Then
                    Dim swFound As Boolean = False, swFound2 As Boolean = False, swFound3 As Boolean = False, swFound4 As Boolean = False, swFound5 As Boolean = False
                    Dim swFound6 As Boolean = False, swFound7 As Boolean = False, swFound8 As Boolean = False, swFound9 As Boolean = False
                    For Each clsProcess As Process In System.Diagnostics.Process.GetProcesses()
                        '**************************************************
                        '* 2015-04-28 RFK:
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 2]") Then swFound2 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 3]") Then swFound3 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 4]") Then swFound4 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 5]") Then swFound5 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 6]") Then swFound6 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 7]") Then swFound7 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 8]") Then swFound8 = True
                        If clsProcess.MainWindowTitle.Contains("aoProcessor") And clsProcess.MainWindowTitle.Contains("[QRY 9]") Then swFound9 = True
                    Next
                    '******************************************************
                    '* 2015-04-28 RFK:
                    Dim sMSG As String = "aoProcessor started additional eProcessor" + vbCrLf
                    Dim swEmail As Boolean = False
                    If Val(Label_NumberQueries0.Text) > 50 And swFound2 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 2", 1)
                        sMSG += "/QRY 2" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 100 And swFound3 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 3", 1)
                        sMSG += "/QRY 3" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 150 And swFound4 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 4", 1)
                        sMSG += "/QRY 4" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 200 And swFound5 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 5", 1)
                        sMSG += "/QRY 5" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 250 And swFound6 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 6", 1)
                        sMSG += "/QRY 6" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 300 And swFound7 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 7", 1)
                        sMSG += "/QRY 7" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 350 And swFound8 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 8", 1)
                        sMSG += "/QRY 8" + vbCrLf
                        swEmail = True
                    End If
                    If Val(Label_NumberQueries0.Text) > 4000 And swFound9 = False Then
                        rkutils.ShellIT(sRunFile, "/QRY 9", 1)
                        sMSG += "/QRY 9" + vbCrLf
                        swEmail = True
                    End If
                    If swEmail Then
                        'rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply_eProcessor@AnnuityHealth.com", "aoProcessor", "IT@AnnuityHealth.com", "IT", "aoProcessor", "aoProcessor Resources", sMSG, "", "")
                    End If
                End If
            End If
            '******************************************************************
            '* 2015-04-28 RFK: 
            If Me.Text.Contains("[STATUS]") = False And Val(Label_NumberStatus0.Text) > 0 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 2]") And Val(Label_NumberStatus0.Text) > 500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 3]") And Val(Label_NumberStatus0.Text) > 1000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 4]") And Val(Label_NumberStatus0.Text) > 1500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 5]") And Val(Label_NumberStatus0.Text) > 2000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 6]") And Val(Label_NumberStatus0.Text) > 2500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 7]") And Val(Label_NumberStatus0.Text) > 3000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 8]") And Val(Label_NumberStatus0.Text) > 3500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 9]") And Val(Label_NumberStatus0.Text) > 4000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            '******************************************************************
            '* 2018-11-14 RFK: 
            If Me.Text.Contains("[STATUS 10]") And Val(Label_NumberStatus0.Text) > 4500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 11]") And Val(Label_NumberStatus0.Text) > 5000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 12]") And Val(Label_NumberStatus0.Text) > 5500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 13]") And Val(Label_NumberStatus0.Text) > 6000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 14]") And Val(Label_NumberStatus0.Text) > 6500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 15]") And Val(Label_NumberStatus0.Text) > 7000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 16]") And Val(Label_NumberStatus0.Text) > 7500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 17]") And Val(Label_NumberStatus0.Text) > 8000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 18]") And Val(Label_NumberStatus0.Text) > 8500 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            If Me.Text.Contains("[STATUS 19]") And Val(Label_NumberStatus0.Text) > 9000 And CheckBox_Status.Checked Then ProcessCommand("STATUS")
            '******************************************************************
            '* 2015-09-28 RFK: 
            If Me.Text.Contains("[NOTE]") = False And Val(Label_NumberNote0.Text) > 0 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 2]") And Val(Label_NumberNote0.Text) > 500 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 3]") And Val(Label_NumberNote0.Text) > 1000 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 4]") And Val(Label_NumberNote0.Text) > 1500 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 5]") And Val(Label_NumberNote0.Text) > 2000 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 6]") And Val(Label_NumberNote0.Text) > 2500 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 7]") And Val(Label_NumberNote0.Text) > 3000 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 8]") And Val(Label_NumberNote0.Text) > 3500 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Me.Text.Contains("[NOTE 9]") And Val(Label_NumberNote0.Text) > 4000 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            '******************************************************************
            '* 2017-10-06 RFK: 
            If Me.Text.Contains("[TRACK]") = False And Val(Label_NumberTrack0.Text) > 0 And CheckBox_Track.Checked Then ProcessCommand("track")
            If Me.Text.Contains("[TRACK 2]") And Val(Label_NumberTrack0.Text) > 500 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 3]") And Val(Label_NumberTrack0.Text) > 1000 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 4]") And Val(Label_NumberTrack0.Text) > 1500 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 5]") And Val(Label_NumberTrack0.Text) > 2000 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 6]") And Val(Label_NumberTrack0.Text) > 2500 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 7]") And Val(Label_NumberTrack0.Text) > 3000 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 8]") And Val(Label_NumberTrack0.Text) > 3500 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            If Me.Text.Contains("[TRACK 9]") And Val(Label_NumberTrack0.Text) > 4000 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            '******************************************************************
            '* 2015-10-26 RFK: 
            If Me.Text.Contains("[QRY]") = False And Val(Label_NumberQueries0.Text) > 0 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 2]") And Val(Label_NumberQueries0.Text) > 50 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 3]") And Val(Label_NumberQueries0.Text) > 100 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 4]") And Val(Label_NumberQueries0.Text) > 150 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 5]") And Val(Label_NumberQueries0.Text) > 200 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 6]") And Val(Label_NumberQueries0.Text) > 250 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 7]") And Val(Label_NumberQueries0.Text) > 300 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 8]") And Val(Label_NumberQueries0.Text) > 350 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            If Me.Text.Contains("[QRY 9]") And Val(Label_NumberQueries0.Text) > 400 And CheckBox_Note.Checked Then ProcessCommand("QRY")
            '******************************************************************
            'If Val(Label_NumberEMail0.Text) > 0 And CheckBox_email.Checked Then ProcessCommand("EMAIL")
            '* 2021-07-21 RFK: 
            If Val(Label_NumberEMailSecure0.Text) > 0 And CheckBox_EMailSecure.Checked Then ProcessCommand("EMAILAR")
            'If Val(Label_NumberEMailSecure0.Text) > 0 And CheckBox_EMailSecure.Checked Then ProcessCommand("EMAILSEC")
            If Val(Label_NumberNote0.Text) > 0 And CheckBox_Note.Checked Then ProcessCommand("NOTE")
            If Val(Label_NumberTrack0.Text) > 0 And CheckBox_Track.Checked Then ProcessCommand("TRACK")
            '**************************************
        Catch ex As Exception
            MsgError("ProcessInit", ex.ToString)
        End Try
    End Sub

    Private Sub ProcessCommand(ByVal tCommandType As String)
        'Try
        Dim tUpdateString As String = "", tSendString As String = "", tFileName As String = ""
        Dim tSQLcommand As String = "SELECT TOP 25 *"
        Dim tQRY_tunique As String = "", tQRY_name As String = ""
        '**********************************************************************
        '* 2015-07-09 RFK: 
        Select Case sSITE
            Case "AnnuityOne"
                tSQLcommand += " FROM RevMD.dbo.commands"
            Case Else
                tSQLcommand += " FROM iTeleCollect.dbo.commands"
        End Select
        '**********************************************************************
        Dim iCommandsLoop As Integer = 0
        Dim sUniques As String = ""
        Dim iTotal As Integer = 25
        Dim iSubProcessor As Integer = 1
        Label_TimeStart.Text = rkutils.STR_format("TODAY", "HH:MM:SS")
        '**********************************************************************
        If Me.Text.Contains("[NOTE") Or Me.Text.Contains("[STATUS") Then
            iTotal = 100
            iSubProcessor = iTotal
            If Me.Text.Contains(" 2]") Then
                iSubProcessor = 500
            Else
                If Me.Text.Contains(" 3]") Then
                    iSubProcessor = 1000
                Else
                    If Me.Text.Contains(" 4]") Then
                        iSubProcessor = 1500
                    Else
                        If Me.Text.Contains(" 5]") Then
                            iSubProcessor = 2000
                        Else
                            If Me.Text.Contains(" 6]") Then
                                iSubProcessor = 2500
                            Else
                                If Me.Text.Contains(" 7]") Then
                                    iSubProcessor = 3000
                                Else
                                    If Me.Text.Contains(" 8]") Then
                                        iSubProcessor = 3500
                                    Else
                                        If Me.Text.Contains(" 9]") Then
                                            iSubProcessor = 4000
                                        Else
                                            If Me.Text.Contains(" 10]") Then
                                                iSubProcessor = 4500
                                            Else
                                                If Me.Text.Contains(" 11]") Then
                                                    iSubProcessor = 5000
                                                Else
                                                    If Me.Text.Contains(" 12]") Then
                                                        iSubProcessor = 5500
                                                    Else
                                                        If Me.Text.Contains(" 13]") Then
                                                            iSubProcessor = 6000
                                                        Else
                                                            If Me.Text.Contains(" 14]") Then
                                                                iSubProcessor = 6500
                                                            Else
                                                                If Me.Text.Contains(" 15]") Then
                                                                    iSubProcessor = 7000
                                                                Else
                                                                    If Me.Text.Contains(" 16]") Then
                                                                        iSubProcessor = 7500
                                                                    Else
                                                                        If Me.Text.Contains(" 17]") Then
                                                                            iSubProcessor = 8000
                                                                        Else
                                                                            If Me.Text.Contains(" 19]") Then
                                                                                iSubProcessor = 8500
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            '**************************************************************
            '* 2015-07-09 RFK: 
            '* 2018-11-14 RFK: 
            Select Case sSITE
                Case "AnnuityOne"
                    tSQLcommand = "SELECT TOP " + iSubProcessor.ToString.Trim + " * FROM RevMD.dbo.commands"
                Case Else
                    tSQLcommand = "SELECT TOP " + iSubProcessor.ToString.Trim + " * FROM iTeleCollect.dbo.commands"
            End Select
            '******************************************************************
            iCommandsLoop = iSubProcessor - iTotal
            '******************************************************************
        End If
        '**********************************************************************
        '* 2015-10-26 RFK: 
        If Me.Text.Contains("[QRY") Then
            If Me.Text.Contains(" 2]") Then
                Select Case sSITE
                    Case "AnnuityOne"
                        tSQLcommand = "SELECT TOP 50 * FROM RevMD.dbo.commands"
                    Case Else
                        tSQLcommand = "SELECT TOP 50 * FROM iTeleCollect.dbo.commands"
                End Select
                '**************************************************************
                iTotal = 100
                iCommandsLoop = 500 - iTotal
            Else
                If Me.Text.Contains(" 3]") Then
                    Select Case sSITE
                        Case "AnnuityOne"
                            tSQLcommand = "SELECT TOP 100 * FROM RevMD.dbo.commands"
                        Case Else
                            tSQLcommand = "SELECT TOP 100 * FROM iTeleCollect.dbo.commands"
                    End Select
                    '**********************************************************
                    iTotal = 100
                    iCommandsLoop = 1000 - iTotal
                Else
                    If Me.Text.Contains(" 4]") Then
                        Select Case sSITE
                            Case "AnnuityOne"
                                tSQLcommand = "SELECT TOP 150 * FROM RevMD.dbo.commands"
                            Case Else
                                tSQLcommand = "SELECT TOP 150 * FROM iTeleCollect.dbo.commands"
                        End Select
                        '******************************************************
                        iTotal = 100
                        iCommandsLoop = 1500 - iTotal
                    Else
                        If Me.Text.Contains(" 5]") Then
                            Select Case sSITE
                                Case "AnnuityOne"
                                    tSQLcommand = "SELECT TOP 200 * FROM RevMD.dbo.commands"
                                Case Else
                                    tSQLcommand = "SELECT TOP 200 * FROM iTeleCollect.dbo.commands"
                            End Select
                            '**************************************************
                            iTotal = 100
                            iCommandsLoop = 2000 - iTotal
                        Else
                            If Me.Text.Contains(" 6]") Then
                                Select Case sSITE
                                    Case "AnnuityOne"
                                        tSQLcommand = "SELECT TOP 250 * FROM RevMD.dbo.commands"
                                    Case Else
                                        tSQLcommand = "SELECT TOP 250 * FROM iTeleCollect.dbo.commands"
                                End Select
                                '**********************************************
                                iTotal = 100
                                iCommandsLoop = 2500 - iTotal
                            Else
                                If Me.Text.Contains(" 7]") Then
                                    Select Case sSITE
                                        Case "AnnuityOne"
                                            tSQLcommand = "SELECT TOP 300 * FROM RevMD.dbo.commands"
                                        Case Else
                                            tSQLcommand = "SELECT TOP 300 * FROM iTeleCollect.dbo.commands"
                                    End Select
                                    '******************************************
                                    iTotal = 100
                                    iCommandsLoop = 3000 - iTotal
                                Else
                                    If Me.Text.Contains(" 8]") Then
                                        Select Case sSITE
                                            Case "AnnuityOne"
                                                tSQLcommand = "SELECT TOP 350 * FROM RevMD.dbo.commands"
                                            Case Else
                                                tSQLcommand = "SELECT TOP 350 * FROM iTeleCollect.dbo.commands"
                                        End Select
                                        '**************************************
                                        iTotal = 100
                                        iCommandsLoop = 3500 - iTotal
                                    Else
                                        If Me.Text.Contains(" 9]") Then
                                            Select Case sSITE
                                                Case "AnnuityOne"
                                                    tSQLcommand = "SELECT TOP 400 * FROM RevMD.dbo.commands"
                                                Case Else
                                                    tSQLcommand = "SELECT TOP 400 * FROM iTeleCollect.dbo.commands"
                                            End Select
                                            '**********************************
                                            iTotal = 100
                                            iCommandsLoop = 4000 - iTotal
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        '**********************************************************************
        tSQLcommand += " WHERE COMMAND='" + tCommandType + "' AND COMPLETED_DATE IS NULL ORDER BY PRIORITY DESC, TUNIQUE"
        '******************************************************************
        If swTEST = True Then Exit Sub
        '**********************************************************************
        If SQL_READ_DATAGRID(DataGridView_COMMANDS, "MSSQL", "*", msSQLConnectionString, msSQLuser, tSQLcommand) Then
            MsgStatus(tCommandType + " " + Me.DataGridView_COMMANDS.RowCount.ToString, True)
            '******************************************************************
            Do While iCommandsLoop < Me.DataGridView_COMMANDS.RowCount - 1 And CheckBox_RUN.Checked
                System.Windows.Forms.Application.DoEvents()
                '**************************************************************
                tUNIQUE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "TUNIQUE", iCommandsLoop)
                tCLIENT = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "CLIENT", iCommandsLoop)
                tTOB = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "TOB", iCommandsLoop)
                tFACILITY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "FACILITY", iCommandsLoop)
                tLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "LOCX", iCommandsLoop)
                tACCOUNTNUMBER = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "ACCOUNTNUMBER", iCommandsLoop)
                tSUFFIX = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "SUFFIX", iCommandsLoop)
                '**************************************************************
                Select Case tCommandType
                    Case "EMAIL"
                        tEMAILfrom = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILFROM", iCommandsLoop)
                        tEMAILfromname = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILFROMNAME", iCommandsLoop)
                        tEMAILto = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILTO", iCommandsLoop)
                        tEMAILtoname = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILTONAME", iCommandsLoop)
                        tEMAILcc = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILCC", iCommandsLoop)
                        tEMAILbcc = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILBCC", iCommandsLoop)
                        tEMAILsubject = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILSUBJECT", iCommandsLoop)
                        tEMAILmessage = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILMESSAGE", iCommandsLoop)
                        tEMAILattach = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILATTACH", iCommandsLoop)
                        'Handled in the Scheduler still
                    Case "EMAILSEC", "EMAILAR"
                        tEMAILfrom = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILFROM", iCommandsLoop)
                        tEMAILfromname = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILFROMNAME", iCommandsLoop)
                        tEMAILto = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILTO", iCommandsLoop)
                        tEMAILtoname = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILTONAME", iCommandsLoop)
                        tEMAILcc = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILCC", iCommandsLoop)
                        tEMAILbcc = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILBCC", iCommandsLoop)
                        tEMAILsubject = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILSUBJECT", iCommandsLoop)
                        tEMAILmessage = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILMESSAGE", iCommandsLoop)
                        tEMAILattach = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "EMAILATTACH", iCommandsLoop)
                        If tEMAILfrom.Contains("@") And tEMAILto.Contains("@") Then
                            email_SECURE()
                        End If
                        UpdateCommand(tUNIQUE)
                    Case "STATUS"
                        Label_NumberStatus0.Text = Str(Val(Label_NumberStatus0.Text) - 1).Trim
                        Label_NumberStatus0.Refresh()
                        tSTATUS = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "STATUS", iCommandsLoop).Replace(" ", "")
                        tNOTE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "TNOTE", iCommandsLoop)
                        tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iCommandsLoop)
                        tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "INSERT_DATE", iCommandsLoop)
                        If IsDate(tDATE) = False Then tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_DATE", iCommandsLoop)
                        '******************************************************
                        '* 2018-01-25 RFK:
                        If IsDate(tDATE) = False Then tDATE = rkutils.STR_format("TODAY", "mm/dd/ccyy HH:MM:SS")
                        '******************************************************
                        tQRY_tunique = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "QRY_TUNIQUE", iCommandsLoop)
                        tQRY_name = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "QRY_NAME", iCommandsLoop)
                        If Val(tLOCX) > 0 Then
                            MsgStatus(tCommandType + " {" + tSTATUS + "}(" + tLOCX + ")[" + tDATE + "][" + tBY + "](" + iCommandsLoop.ToString + ")", True)
                            '**************************************************
                            '* 2016-01-26 RFK:
                            If tQRY_tunique.Length > 0 And tQRY_name.Length > 0 Then
                                rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", "QRY[" + tQRY_tunique + "] " + tQRY_name + " STATUS:" + tSTATUS, tDATE, tBY)
                            End If
                            rkutils.LOCX_STATUS("DB2", DB2SQLConnectionString, DB2SQLuser, msSQLConnectionString, msSQLuser, DataGridView2, DataGridView3, tLOCX, tCC, "", tSTATUS, tNOTE, tDATE, tBY, tQRY_name)
                            '**************************************************
                        End If
                        '**************************************************
                        '* 2015-04-28 RFK: UpdateCommand(tUNIQUE)
                        If sUniques.Length > 0 Then sUniques += ","
                        sUniques += "'" + tUNIQUE + "'"
                        '**************************************************
                    Case "NOTE"
                        Label_NumberNote0.Text = Str(Val(Label_NumberNote0.Text) - 1).Trim
                        Label_NumberNote0.Refresh()
                        tNOTE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "TNOTE", iCommandsLoop)
                        tMC = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MC", iCommandsLoop)
                        tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iCommandsLoop)
                        tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "INSERT_DATE", iCommandsLoop)
                        If IsDate(tDATE) = False Then
                            tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_DATE", iCommandsLoop)
                        End If
                        If tBY.Length < 1 Then tBY = rkutils.WhoAmI()
                        If Val(tLOCX) > 0 Then
                            MsgStatus(tCommandType + " [" + tLOCX + "]{" + tNOTE + "}" + tDATE + "(" + tBY + ")", True)
                            rkutils.NOTES_ADD("DB2", DB2SQLConnectionString, DB2SQLuser, tBY, Me.DataGridView2, tLOCX, "1", tMC, tSTATUS, "", "", "", tNOTE, tDATE, tBY, "")
                        End If
                        '**************************************************
                        '* 2015-04-28 RFK: UpdateCommand(tUNIQUE)
                        If sUniques.Length > 0 Then sUniques += ","
                        sUniques += "'" + tUNIQUE + "'"
                        '**************************************************
                    Case "TRACK"
                        Label_NumberTrack0.Text = Val(Label_NumberTrack0.Text) - 1
                        Label_NumberTrack0.Refresh()
                        tNOTE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "TNOTE", iCommandsLoop)
                        tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iCommandsLoop)
                        tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "INSERT_DATE", iCommandsLoop)
                        If IsDate(tDATE) = False Then
                            tDATE = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_DATE", iCommandsLoop)
                        End If
                        If tBY.Length < 1 Then tBY = rkutils.WhoAmI()
                        If Val(tLOCX) > 0 Then
                            MsgStatus(tCommandType + " [" + tLOCX + "]{" + tNOTE + "}" + tDATE + " " + tBY, True)
                            rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", tNOTE, tDATE, tBY)
                        End If
                        '**************************************************
                        '* 2015-04-28 RFK: UpdateCommand(tUNIQUE)
                        If sUniques.Length > 0 Then sUniques += ","
                        sUniques += "'" + tUNIQUE + "'"
                        '**************************************************
                End Select
                '**********************************************************
                iCommandsLoop += 1
            Loop
            '**************************************************************
            '* 2015-04-28 RFK: 
            Select Case tCommandType
                Case "NOTE", "STATUS", "TRACK"
                    If sUniques.Length > 0 Then UpdateCommands(sUniques)
            End Select
            '**************************************************************
            Label_TimeFinish.Text = rkutils.STR_format("TODAY", "HH:MM:SS")
            Dim lSeconds As Long = rkutils.TimeToSeconds(Label_TimeFinish.Text) - rkutils.TimeToSeconds(Label_TimeStart.Text)
            Label_TimeDuration.Text = rkutils.SecondsToTime(lSeconds, 8)
            Dim lNumberPerSecond As Long = 0
            If lSeconds > 0 Then
                lNumberPerSecond = iTotal / lSeconds
            End If
            Dim lNumberPerHour As Long = (lNumberPerSecond * 60) * 60
            Label_AnHour.Text = Str(lNumberPerHour) + " per hour"
            '**************************************************************
        End If
        'Catch ex As Exception
        '    MsgError("ProcessCommand", ex.ToString)
        'End Try
    End Sub

    Private Sub QRYread(ByVal iRow As Integer)
        Try
            Label_QRY_tUnique.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TUNIQUE", iRow)
            Label_QRY_Name.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "NAME", iRow)
            Label_QRY_Type.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "TYPE", iRow)
            Label_QRY_Who.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "WHO", iRow)
            Label_QRY_Description.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "DESCRIPTION", iRow)
            Label_QRY_Status.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "STATUSCODE", iRow)
            Label_SetField.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "SET_FIELD", iRow)
            Label_SetValue.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "SET_VALUE", iRow)
            Label_QRY_StatusMatch.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "STATUSMATCH", iRow)
            Label_QRY_EmailType.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL_TYPE", iRow)
            Label_QRY_Emailblank.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL_BLANK", iRow)
            Label_QRY_IncludeQuery.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL_INCLUDEQUERY", iRow)      '* 2020-01-08 RFK: 
            Label_QRY_EMail.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL", iRow)
            Label_QRY_EmailMessage.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL_MESSAGE", iRow)
            Label_QRY_Letter_Number.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "LETTER_NUMBER", iRow)
            Label_QRY_Letter_Date.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "LETTER_DATE", iRow)
            Label_QRY_Output_Type.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "OUTPUT_TYPE", iRow)
            Label_QRY_Output_CopyTo.Text = STR_convert_Macro(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "OUTPUT_COPYTO", iRow))
            Label_QRY_Output_CopyTo2.Text = STR_convert_Macro(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "OUTPUT_COPYTO2", iRow))
            Label_QRY_Output_CopyToEMailDir.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "OUTPUT_COPYTOEMAILDIR", iRow)
            Label_QRY.Text = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "QRY", iRow)
            ResizeApp()
        Catch ex As Exception
            MsgError("QRYread", ex.ToString)
        End Try
    End Sub

    Private Sub COMMANDrun()
        Try
            If Label_QRY.Text.Contains("ROIDATA.") Then
                MsgStatus("DB2:" + Label_QRY.Text, True)
                DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, Label_QRY.Text)
            Else
                '**********************************************************************************
                '* 2021-12-29 RFK:
                If Label_QRY.Text.ToUpper.Contains("[SQLAO]") Or Label_QRY.Text.ToUpper.Contains("[SQLMSAO]") Then
                    MsgStatus("SQLAO:" + Label_QRY.Text, True)
                    DB_COMMAND("MSSQL", msSQLAOconnection, msSQLAOuser, Label_QRY.Text.Replace("[SQLAO]", "").Replace("[SQLMSAO]", ""))
                Else
                    '******************************************************************************
                    '* 2020-01-08 RFK:
                    If Label_QRY.Text.ToUpper.Contains("[SQLMS1]") Or Label_QRY.Text.ToUpper.Contains("[SQLMS1]") Then
                        MsgStatus("MSSQL:" + Label_QRY.Text, True)
                        DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, Label_QRY.Text.Replace("[SQLMS]", "").Replace("[SQLMS1]", ""))
                    Else
                        '**************************************************************************
                        '* 2020-01-08 RFK:
                        If Label_QRY.Text.ToUpper.Contains("[SQLMS2]") Then
                            MsgStatus("SQLMS2:" + Label_QRY.Text, True)
                            DB_COMMAND("MSSQL", msSQL2ConnectionString, msSQL2user, Label_QRY.Text.Replace("[SQLMS2]", ""))
                        Else
                            '**********************************************************************
                            '* 2020-01-08 RFK:
                            If Label_QRY.Text.ToUpper.Contains("[SQLMS3]") Then
                                MsgStatus("SQLMS3:" + Label_QRY.Text, True)
                                DB_COMMAND("MSSQL", msSQL3ConnectionString, msSQL3user, Label_QRY.Text.Replace("[SQLMS3]", ""))
                            Else
                                '******************************************************************
                                '* 2020-01-08 RFK:
                                If Label_QRY.Text.ToUpper.Contains("REVMD.DBO.") Or Label_QRY.Text.ToUpper.Contains("[REVMD].[DBO].") Then
                                    MsgStatus("MSSQL:" + Label_QRY.Text, True)
                                    DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, Label_QRY.Text)
                                Else
                                    '**************************************************************
                                    '* 2020-01-08 RFK:
                                    If Label_QRY.Text.ToUpper.Contains("[SQLMSENHANCE]") Then
                                        MsgStatus("SQLMSENHANCE:" + Label_QRY.Text, True)
                                        DB_COMMAND("MSSQL", msSQL3ConnectionString, msSQL3user, Label_QRY.Text.Replace("[SQLMSENHANCE]", ""))
                                    Else
                                        '**********************************************************
                                        '* 2018-09-05 RFK:
                                        If Label_QRY.Text.ToUpper.Contains("TELESERVER.DBO.") Or Label_QRY.Text.ToUpper.Contains("[TELESERVER].[DBO].") Then
                                            MsgStatus("MSSQL:" + Label_QRY.Text, True)
                                            DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, Label_QRY.Text)
                                        Else
                                            '******************************************************
                                            '* 2018-09-05 RFK:
                                            If Label_QRY.Text.ToUpper.Contains("DIALER.DBO.") Or Label_QRY.Text.ToUpper.Contains("[DIALER].[DBO].") Then
                                                MsgStatus("MSSQL2:" + Label_QRY.Text, True)
                                                DB_COMMAND("MSSQL2", msSQL2ConnectionString, msSQL2user, Label_QRY.Text)
                                            Else
                                                MsgStatus("COMMAND UNABLE:" + Label_QRY.Text, True)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgStatus("COMMANDrun" + ex.ToString, True)
        End Try
    End Sub

    Private Sub QRYselect()
        Try
            'File.AppendAllText("c:\temp\temp.txt", Label_QRY_Description.Text + vbCrLf + Label_QRY.Text.Replace(vbCr, vbCrLf) + vbCrLf)
            If Label_QRY.Text.ToUpper.Contains("ROIDATA.") Then
                MsgStatus("QRY DB2:" + Label_QRY_Name.Text, True)
                MsgStatus(Label_QRY.Text, True)
                '*********************************
                If swQRYtable Then
                    rkutils.SQL_READ_DATATABLE(DTqry, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, Label_QRY.Text)
                    If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                Else
                    DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "DB2", "*", DB2SQLConnectionString, DB2SQLuser, Label_QRY.Text)
                    DataGridView_QRYoutput.Visible = True
                    Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                End If
                MsgStatus("QRY DB2:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
            Else
                If Label_QRY.Text.ToUpper.Contains("[SQLAO]") Or Label_QRY.Text.ToUpper.Contains("[SQLMSAO]") Then
                    If swQRYtable Then
                        rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQLAOconnection, msSQLAOuser, Label_QRY.Text.Replace("[SQLAO]", "").Replace("[SQLMSAO]", ""))
                        If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                    Else
                        DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQLAOconnection, msSQLAOuser, Label_QRY.Text.Replace("[SQLAO]", "").Replace("[SQLMSAO]", ""))
                        DataGridView_QRYoutput.Visible = True
                        Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                    End If
                    MsgStatus("QRY MSSQLAO:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                Else
                    If Label_QRY.Text.ToUpper.Contains("[SQLMS]".ToUpper) Or Label_QRY.Text.ToUpper.Contains("[SQLMS1]".ToUpper) Then
                        If swQRYtable Then
                            rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQLConnectionString, msSQLuser, Label_QRY.Text.Replace("[SQLMS]", "").Replace("[SQLMS1]", ""))
                            If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                        Else
                            DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQLConnectionString, msSQLuser, Label_QRY.Text.Replace("[SQLMS]", "").Replace("[SQLMS1]", ""))
                            DataGridView_QRYoutput.Visible = True
                            Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                        End If
                        MsgStatus("QRY MSSQL:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                    Else
                        If Label_QRY.Text.ToUpper.Contains("[SQLMS2]".ToUpper) Then
                            If swQRYtable Then
                                rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text.Replace("[SQLMS2]", ""))
                                If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                            Else
                                DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text.Replace("[SQLMS2]", ""))
                                DataGridView_QRYoutput.Visible = True
                                Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                            End If
                            MsgStatus("QRY MSSQL2:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                        Else
                            If Label_QRY.Text.ToUpper.Contains("[SQLMS3]".ToUpper) Then
                                If swQRYtable Then
                                    rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQL3ConnectionString, msSQL3user, Label_QRY.Text.Replace("[SQLMS3]", ""))
                                    If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                                Else
                                    DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQL3ConnectionString, msSQL3user, Label_QRY.Text.Replace("[SQLMS3]", ""))
                                    DataGridView_QRYoutput.Visible = True
                                    Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                                End If
                                MsgStatus("QRY MSSQL3:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                            Else
                                If Label_QRY.Text.ToUpper.Contains("[SQLMSENHANCE]".ToUpper) Then
                                    'If swQRYtable Then
                                    '    rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQLEConnectionString, msSQLEuser, Label_QRY.Text.Replace("[SQLMSENHANCE]", ""))
                                    '    If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                                    'Else
                                    '    DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQLEConnectionString, msSQLEuser, Label_QRY.Text.Replace("[SQLMSENHANCE]", ""))
                                    '    DataGridView_QRYoutput.Visible = True
                                    '    Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                                    'End If
                                    MsgStatus("QRY MSSQLENHANCE:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                                Else
                                    If Label_QRY.Text.ToUpper.Contains("DIALER.DBO.".ToUpper) Then
                                        If swQRYtable Then
                                            rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text)
                                            If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                                        Else
                                            DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text)
                                            DataGridView_QRYoutput.Visible = True
                                            Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                                        End If
                                        MsgStatus("QRY MSSQL2:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                                    Else
                                        '**********************************************************
                                        '* 2019-02-04 RFK: Added iIncident.dbo.
                                        If Label_QRY.Text.ToUpper.Contains("iINCIDENT.DBO.".ToUpper) Then
                                            If swQRYtable Then
                                                rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text)
                                                If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                                            Else
                                                DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQL2ConnectionString, msSQL2user, Label_QRY.Text)
                                                DataGridView_QRYoutput.Visible = True
                                                Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                                            End If
                                            MsgStatus("QRY MSSQL2:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                                        Else
                                            If Label_QRY.Text.ToUpper.Contains(".DBO.".ToUpper) Then
                                                If swQRYtable Then
                                                    rkutils.SQL_READ_DATATABLE(DTqry, "MSSQL", "*", msSQLConnectionString, msSQLuser, Label_QRY.Text)
                                                    If DTqry IsNot Nothing Then Label_QueriesAccounts.Text = Trim(Str(DTqry.Rows.Count - 1))
                                                Else
                                                    DataGridView_QRYoutput.Visible = SQL_READ_DATAGRID(DataGridView_QRYoutput, "MSSQL", "*", msSQLConnectionString, msSQLuser, Label_QRY.Text)
                                                    DataGridView_QRYoutput.Visible = True
                                                    Label_QueriesAccounts.Text = Trim(Str(DataGridView_QRYoutput.RowCount - 1))
                                                End If
                                                MsgStatus("QRY MSSQL:" + Label_QRY_Name.Text + " selected:" + Label_QueriesAccounts.Text, True)
                                            Else
                                                If swQRYtable Then

                                                Else
                                                    Me.DataGridView_QRYoutput.Visible = False
                                                End If
                                                MsgStatus("QRY ELSE:" + Label_QRY_Name.Text + " " + Label_QRY.Text, True)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            '*******************************************
            Label_QueriesAccountsCount.Text = Label_QueriesAccounts.Text
            ResizeApp()
        Catch ex As Exception
            MsgError("QRYselect", ex.ToString)
        End Try
    End Sub

    Private Sub QRYrunDT(ByVal sType As String, ByVal bLiveRun As Boolean)
        Try
            Dim tMessage As String = "", tSQL As String = "", sSQL As String = ""
            Dim iQRYLoop As Integer = 0
            Dim tReportName As String = ""
            Dim tHTML As String = "", tHTMLline As String = "", tTEMP As String = ""
            Dim tReportLine As String = "", tReportAllLines As String = "", tDelimiter As String = vbTab
            Dim sColumn As String = ""
            Dim iCol As Integer = 0, iCol2 As Integer = 0, iColTotal As Integer = 0, iColAverage As Integer = 0, iRow As Integer = 0, iSum As Integer = 0, iTotal As Integer = 0
            Dim swTotal As Boolean = False, swIncludeHTML As Boolean = False
            '******************************************************************
            '* 2015-05-22 RFK:
            Select Case Label_QRY_Output_Type.Text
                Case "C"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".CSV"
                    tDelimiter = ","
                    MsgStatus("QRYrunDT/Comma-CSV", True)
                Case "P"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".TXT"
                    tDelimiter = "|"
                    MsgStatus("QRYrunDT/Pipe", True)
                Case "E"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".XLS"
                Case Else
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".XLS"
            End Select
            '*******************************************************************
            '* 2013-09-09 RFK: 
            MsgStatus("QRYrunDT:" + sType + " [" + Label_RUNNING.Text + "][" + Label_QueriesAccounts.Text + "][" + tReportName + "]", True)
            If File.Exists(tReportName) Then
                MsgStatus(tReportName + " already exists", True)
                Exit Sub
            End If
            '******************************************************************
            '* 2015-05-22 RFK: StreamWriter
            Dim sw As New StreamWriter(tReportName, False)
            '******************************************************************
            Label_QRY_Running.Text = sType + "-" + Label_RUNNING.Text
            '******************************************************************
            '* 2013-09-09 RFK: 
            'If Label_RUNNING.Text <> "Ready" Then Exit Sub
            Label_RUNNING.Text = "Running"
            Label_RUNNING.Refresh()
            System.Windows.Forms.Application.DoEvents()
            If Val(Label_QueriesAccounts.Text) > 0 Then
                '**************************************************************
                '* 2013-03-21 RFK:
                MsgStatus("Creating " + sType + " " + tReportName, True)
                MsgStatus(Label_RUNNING.Text, True)
                '**************************************************************
                '* 2015-05-22 RFK: Is it being included HTML
                swIncludeHTML = False
                Select Case Label_QRY_EmailType.Text
                    Case "A"    'Attached / Include
                        If Val(Label_QueriesAccountsCount.Text) < 1000 Then
                            swIncludeHTML = True
                        End If
                End Select
                '**************************************************************
                '* 2015-05-22 RFK: Does it contain Columns
                swTotal = True
                '**************************************************************
                For iCol = 0 To DTqry.Columns.Count - 1
                    If DTqry.Rows(iQRYLoop)(iCol) IsNot Nothing Then
                        Select Case DTqry.Rows(iQRYLoop)(iCol).ToString.ToUpper
                            Case "Total".ToUpper
                                swTotal = True
                        End Select
                    End If
                Next
                '**************************************************************
                If swIncludeHTML Then
                    tHTMLline = "<table width=100% cellpadding=2 cellspacing=2 border=1>"
                End If
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                Do While Label_RUNNING.Text = "Running" And iQRYLoop < DTqry.Rows.Count - 1
                    rkutils.DoEvents()
                    tLOCX = rkutils.DataTable_ValueByColumnName(DTqry, "RALOCX", iQRYLoop).Trim
                    '**********************************************************
                    tBY = rkutils.DataTable_ValueByColumnName(DTqry, "MODIFIED_BY", iQRYLoop).Trim
                    'If tBY.Length = 0 Then
                    '    tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iQRYLoop).Trim
                    'End If
                    '**********************************************************
                    '* 2013-03-21 RFK:
                    '* 2013-10-22 RFK: Allow report if no LOCX in it
                    Select Case sType
                        Case "Report"
                            '**************************************************
                            '* 2012-10-12 RFK: The Header Row
                            tReportLine = ""
                            If iQRYLoop = 0 Then
                                If tLOCX.Length > 0 Then
                                    '* 2017-06-07 RFK:
                                    '* 2021-11-03 RFK:
                                    Select Case Label_QRY_Output_Type.Text
                                        Case "C"
                                            tReportLine += Chr(34) + "LOCX" + Chr(34) + tDelimiter
                                        Case "P"
                                            '* 2021-09-30 RFK: No Header in PIPE DELIMITED (NEED TO CREATE HEADER SWITCH)
                                        Case "E"
                                            tReportLine += "LOCX" + tDelimiter
                                        Case Else
                                            tReportLine += "LOCX" + tDelimiter
                                    End Select
                                End If
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "<tr>"
                                End If
                                '**********************************************
                                For iCol = 0 To DTqry.Columns.Count - 1
                                    If DTqry.Columns(iCol).ColumnName IsNot Nothing Then
                                        MsgStatus(DTqry.Columns(iCol).ColumnName.ToUpper, True)
                                        sColumn = DTqry.Columns(iCol).ColumnName.ToString.ToUpper
                                        Select Case sColumn
                                            Case "TODAY"
                                                sColumn = rkutils.STR_format("TODAY", "mm-dd-ccyy")
                                            Case "DAY1"
                                                sColumn = rkutils.STR_format("TODAY-1", "DOW")
                                            Case "DAY2"
                                                sColumn = rkutils.STR_format("TODAY-2", "DOW")
                                            Case "DAY3"
                                                sColumn = rkutils.STR_format("TODAY-3", "DOW")
                                            Case "DAY4"
                                                sColumn = rkutils.STR_format("TODAY-4", "DOW")
                                            Case "DAY5"
                                                sColumn = rkutils.STR_format("TODAY-5", "DOW")
                                            Case "DAY6"
                                                sColumn = rkutils.STR_format("TODAY-6", "DOW")
                                            Case "DAY7"
                                                sColumn = rkutils.STR_format("TODAY-7", "DOW")
                                        End Select
                                        '**************************************
                                        '* 2017-06-07 RFK:
                                        Select Case Label_QRY_Output_Type.Text
                                            Case "C"
                                                tReportLine += Chr(34) + sColumn + Chr(34) + tDelimiter
                                            Case "P"
                                                '* 2021-09-30 RFK: No Header in PIPE DELIMITED (NEED TO CREATE HEADER SWITCH)
                                            Case Else
                                                MsgStatus("DEBUG4:" + sColumn, True)
                                                tReportLine += sColumn + tDelimiter
                                        End Select
                                        '**************************************
                                        If swIncludeHTML Then
                                            tHTMLline += "<td>" + sColumn + "</td>"
                                        End If
                                        '**************************************
                                    End If
                                Next
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "</tr>"
                                End If
                                '**********************************************
                                '* 2015-04-22 RFK:
                                '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                                'sw.WriteLine(tReportLine)
                                '* 2021-09-30 RFK: No Header in PIPE DELIMITED (NEED TO CREATE HEADER SWITCH)
                                Select Case Label_QRY_Output_Type.Text
                                    Case "P"
                                        MsgStatus("DEBUG5:" + DTqry.Columns(0).ColumnName, True)
                                        If DTqry.Columns.Count > 1 Then
                                            If DTqry.Columns(0).ColumnName IsNot Nothing Then
                                                MsgStatus("DEBUG6:" + DTqry.Columns(0).ColumnName, True)
                                            End If
                                        End If
                                    Case Else
                                        sw.WriteLine(tReportLine)
                                End Select
                            End If
                            '**************************************************
                            '* 2012-10-12 RFK: Date Line
                            tReportLine = ""
                            If tLOCX.Length > 0 Then
                                '**********************************************
                                '* 2017-06-07 RFK:
                                Select Case Label_QRY_Output_Type.Text
                                    Case "C"
                                        tReportLine += Chr(34) + tLOCX + Chr(34) + tDelimiter
                                    Case "P"
                                        '* 2021-09-30 RFK: No Header in PIPE DELIMITED (NEED TO CREATE HEADER SWITCH)
                                    Case Else
                                        tReportLine += tLOCX + tDelimiter
                                End Select
                                '**********************************************
                            End If
                            '**************************************************
                            If swIncludeHTML Then
                                tHTMLline += "<tr>"
                            End If
                            '**************************************************
                            '* 2015-05-21 RFK:
                            If swIncludeHTML Or swTotal Then
                                For iCol = 0 To DTqry.Columns.Count - 1
                                    If DTqry.Rows(0)(iCol).ToString IsNot Nothing Then
                                        Select Case DTqry.Rows(0)(iCol).ToString.ToUpper
                                            Case "Total".ToUpper
                                                '**********************************
                                                swTotal = True
                                                iColTotal = iCol
                                                iSum = 0
                                                For iCol2 = 1 To iCol - 1
                                                    If DTqry.Rows(iQRYLoop)(iCol2).ToString IsNot Nothing Then
                                                        iSum += Val(DTqry.Rows(iQRYLoop)(iCol2).ToString)
                                                    End If
                                                Next
                                                '******************************
                                                If DTqry.Rows(iQRYLoop)(iCol2).ToString IsNot Nothing Then
                                                    '**************************
                                                    '* 2017-06-07 RFK:
                                                    Select Case Label_QRY_Output_Type.Text
                                                        Case "C"
                                                            tReportLine += Chr(34) + DTqry.Rows(iQRYLoop)(iCol2).ToString + Chr(34)
                                                        Case "P"
                                                            MsgStatus("DEBUG8:" + DTqry.Columns(0).ColumnName, True)
                                                            tReportLine += DTqry.Rows(iQRYLoop)(iCol2).ToString
                                                        Case Else
                                                            tReportLine += DTqry.Rows(iQRYLoop)(iCol2).ToString
                                                    End Select
                                                    '**************************
                                                End If
                                                tReportLine += tDelimiter
                                                '**********************************
                                                If swIncludeHTML Then
                                                    tHTMLline += "<td>"
                                                    If DTqry.Rows(iQRYLoop)(iCol).ToString IsNot Nothing Then
                                                        tHTMLline += DTqry.Rows(iQRYLoop)(iCol).ToString
                                                    End If
                                                    tHTMLline += "</td>"
                                                End If
                                                '**********************************
                                            Case "Average".ToUpper
                                                '**********************************
                                                iColAverage = iCol
                                                iSum = 0
                                                For iCol2 = 1 To iCol - 1
                                                    If DTqry.Rows(iQRYLoop)(iCol2).ToString IsNot Nothing Then
                                                        iSum += Val(DTqry.Rows(iQRYLoop)(iCol2).ToString)
                                                    End If
                                                Next
                                                '******************************
                                                '* 2017-06-07 RFK:
                                                Select Case Label_QRY_Output_Type.Text
                                                    Case "C"
                                                        tReportLine += Chr(34) + rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + Chr(34) + tDelimiter
                                                    Case "P"
                                                        MsgStatus("DEBUG9:" + DTqry.Columns(0).ColumnName, True)
                                                        tReportLine += rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + tDelimiter
                                                    Case Else
                                                        tReportLine += rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + tDelimiter
                                                End Select
                                                '******************************
                                                If swIncludeHTML Then
                                                    tHTMLline += "<td>" + rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + "</td>"
                                                End If
                                                '******************************
                                            Case Else
                                                If DTqry.Rows(iQRYLoop)(iCol).ToString IsNot Nothing Then
                                                    '**************************
                                                    '* 2017-06-07 RFK:
                                                    Select Case Label_QRY_Output_Type.Text
                                                        Case "C"
                                                            tReportLine += Chr(34) + DTqry.Rows(iQRYLoop)(iCol).ToString + Chr(34) + tDelimiter
                                                        Case "P"
                                                            MsgStatus("DEBUG10:" + DTqry.Rows(iQRYLoop)(iCol).ToString, True)
                                                            tReportLine += DTqry.Rows(iQRYLoop)(iCol).ToString + tDelimiter
                                                        Case Else
                                                            tReportLine += DTqry.Rows(iQRYLoop)(iCol).ToString + tDelimiter
                                                    End Select
                                                    '**************************
                                                    If swIncludeHTML Then
                                                        tHTMLline += "<td>" + DTqry.Rows(iQRYLoop)(iCol).ToString + "</td>"
                                                    End If
                                                    '**************************
                                                End If
                                        End Select
                                    End If
                                Next
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "</tr>"
                                End If
                            End If
                            '**************************************************
                            '* 2015-04-22 RFK: 
                            '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                            sw.WriteLine(tReportLine)
                            '**************************************************
                        Case "CallList"
                            'Should NOT get to this point
                    End Select
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '* 2013-10-22 RFK: Only if a LOCX in it
                    If tLOCX.Length > 0 Then
                        '******************************************************
                        '* 2013-04-03 RFK:
                        If Label_QRY_Letter_Number.Text.Trim.Length > 0 Then
                            If Val(Label_QRY_Letter_Number.Text) > 0 And Val(Label_QRY_Letter_Number.Text) <= 999 Then
                                Select Case Label_QRY_Letter_Date.Text
                                    Case "TODAY"
                                        Label_QRY_Letter_Date.Text = STR_format("TODAY", "mm/dd/ccyy")
                                    Case Else
                                End Select
                                If IsDate(Label_QRY_Letter_Date.Text) Then
                                    tSQL = "UPDATE ROIDATA.RACCTP"
                                    tSQL += " SET RALNAC=" + Label_QRY_Letter_Number.Text
                                    tSQL += ",RANLMO=" + STR_format(Label_QRY_Letter_Date.Text, "mm")
                                    tSQL += ",RANLDY=" + STR_format(Label_QRY_Letter_Date.Text, "dd")
                                    tSQL += ",RANLYR=" + STR_format(Label_QRY_Letter_Date.Text, "ccyy")
                                    tSQL += " WHERE RALOCX='" + tLOCX + "'"
                                    rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQL)
                                    MsgStatus(tSQL, False)
                                    '******************************************
                                    tMessage = "QRY:" + Label_QRY_Name.Text + " set letter to:" + Label_QRY_Letter_Number.Text + " [" + Label_QRY_Letter_Date.Text + "]"
                                    rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", tMessage, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                                Else
                                    MsgStatus(Label_QRY_Letter_Date.Text + " is NOT a valid date, can not set to letter " + Label_QRY_Letter_Number.Text, True)
                                End If
                            End If
                        End If
                        '******************************************************
                        '* 2012-10-12 RFK:
                        If Label_QRY_Status.Text.Trim.Length > 0 Then
                            '**************************************************
                            '* 2013-08-22 RFK: TCode/Status if checkbox allowed.
                            If CheckBox_Queries_TCode.Checked Then
                                '**********************************************
                                '* 2013-08-22 RFK: TCode/Status Matched Accounts.
                                '* 2015-05-21 RFK:
                                rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tLOCX, Label_QRY_Status.Text, "", "", "")   ', Label_QRY_tUnique.Text, Label_QRY_Name.Text)
                                '**********************************************
                                'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", "QRY[" + Label_QRY_tUnique.Text + "] " + Label_QRY_Name.Text + " STATUS:" + Label_QRY_Status.Text, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                                '**********************************************
                            Else
                                MsgStatus("Status Account [Not checked to run]" + tLOCX + " " + Label_QRY_Status.Text, True)
                            End If
                        End If
                        '******************************************************
                        '* 2019-03-18 RFK: Set Field/Value
                        MsgStatus("Label_SetField.Text:" + Label_SetField.Text, True)
                        If Label_SetField.Text.Trim.Length > 0 And Label_SetValue.Text.Trim.Length > 0 And CheckBox_Queries_TCode.Checked Then
                            '**************************************************
                            sSQL = "UPDATE ROIDATA.RACCTP SET " + Label_SetField.Text + "='" + Label_SetValue.Text + "'"
                            sSQL += rkutils.WhereAnd(sSQL, "RALOCX='" + tLOCX + "'")
                            MsgStatus(sSQL, True)
                            rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", "QRY[" + Label_QRY_tUnique.Text + "] " + Label_QRY_Name.Text + " SET:" + Label_SetField.Text + "=" + Label_SetValue.Text, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                            '**************************************************
                        End If
                        '******************************************************
                    End If
                    iQRYLoop += 1
                    Label_QueriesAccountsCount.Text = Trim(Str(Val(Label_QueriesAccountsCount.Text) - 1))
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                Loop
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '* 2015-05-14 RFK: Sum Columns
                If swTotal = True And iColTotal > 0 Then
                    '**************************
                    '* 2017-06-07 RFK:
                    Select Case Label_QRY_Output_Type.Text
                        Case "C"
                        Case "P"
                        Case "E"
                        Case Else
                    End Select
                    '**************************
                    tReportLine = "Total" + tDelimiter
                    '**********************************************************
                    Select Case Label_QRY_EmailType.Text
                        Case "A"    'Attached / Include
                            tHTMLline += "<tr>"
                            tHTMLline += "<td>Total</td>"
                    End Select
                    '**********************************************************
                    iTotal = 0
                    For iCol = 1 To iColTotal - 1
                        iSum = 0
                        For iRow = 0 To DTqry.Rows.Count - 1
                            If DTqry.Rows(iRow)(iCol).ToString IsNot Nothing Then
                                iSum += Val(DTqry.Rows(iRow)(iCol).ToString)
                            End If
                        Next
                        '******************************************************
                        iTotal += iSum
                        '******************************************************
                        '* 2017-06-07 RFK:
                        Select Case Label_QRY_Output_Type.Text
                            Case "C"
                                tReportLine += Chr(34) + iSum.ToString + Chr(34) + tDelimiter
                            Case Else
                                tReportLine += iSum.ToString + tDelimiter
                        End Select
                        '******************************************************
                        Select Case Label_QRY_EmailType.Text
                            Case "A"    'Attached / Include
                                tHTMLline += "<td>" + iSum.ToString + "</td>"
                        End Select
                    Next
                    '**********************************************************
                    '**************************
                    '* 2017-06-07 RFK:
                    Select Case Label_QRY_Output_Type.Text
                        Case "C"
                            tReportLine += Chr(34) + iTotal.ToString + Chr(34) + tDelimiter
                        Case Else
                            tReportLine += iTotal.ToString + tDelimiter
                    End Select
                    '**************************
                    '**********************************************************
                    If iColAverage > 0 Then
                        '**************************
                        '* 2017-06-07 RFK:
                        Select Case Label_QRY_Output_Type.Text
                            Case "C"
                                tReportLine += Chr(34) + (iTotal / (iColTotal - 1)).ToString + Chr(34) + tDelimiter
                            Case Else
                                tReportLine += (iTotal / (iColTotal - 1)).ToString + tDelimiter
                        End Select
                        '**************************
                    End If
                    '**********************************************************
                    '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                    sw.WriteLine(tReportLine)
                    sw.Close()
                    '**********************************************************
                    Select Case Label_QRY_EmailType.Text
                        Case "A"    'Attached / Include
                            tHTMLline += "<td>" + iTotal.ToString + "</td>"
                            '**********************************************************
                            If iColAverage > 0 Then
                                tHTMLline += "<td>" + (iTotal / (iColTotal - 1)).ToString + "</td>"
                            End If
                            '**********************************************************
                            tHTMLline += "</tr>"
                    End Select
                End If
                '**************************************************************
                '* 2015-04-22 RFK: 
                'If tReportAllLines.Length > 0 Then
                '    File.AppendAllText(tReportName, tReportAllLines)
                'End If
                '**************************************************************
            Else
                MsgStatus("QRYrun:" + Label_QRY_Name.Text + " selected NOTHING", True)
            End If
            '******************************************************************
            '* 2015-11-09 RFK: Close the streamwriter
            If sw IsNot Nothing Then sw.Close()
            '******************************************************************
            '* 2012-10-12 RFK: Email
            '* 2015-03-30 RFK: Blank (No Results) Email
            MsgStatus("Label_QueriesAccounts.Text=" + Label_QueriesAccounts.Text, True)
            MsgStatus("Label_QRY_Emailblank.Text =" + Label_QRY_Emailblank.Text, True)
            If Val(Label_QueriesAccounts.Text) > 0 Or Label_QRY_Emailblank.Text = "Y" Then
                MsgStatus("Emailing" + sType + " " + tReportName, True)
                '**************************************************************
                tEMAILfrom = "production@AnnuityHealth.com"
                tEMAILfromname = "Production"
                '**************************************************************
                '* 2013-02-27 RFK: Email
                If bLiveRun Then
                    '**********************************************************
                    '* 2015-07-09 RFK: 
                    tSQL = "SELECT LIVE_EMAIL,LIVE_ATTACH"
                    Select Case sSITE
                        Case "AnnuityOne"
                            tSQL += " FROM RevMD.dbo.query "
                        Case Else
                            tSQL += " FROM iTeleCollect.dbo.query"
                    End Select
                    '**********************************************************
                    tSQL += " WHERE TUNIQUE='" + Label_QRY_tUnique.Text + "'"
                    '**********************************************************
                    tEMAILto = rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "LIVE_EMAIL", msSQLConnectionString, msSQLuser, tSQL)
                    tEMAILtoname = tEMAILto
                    Label_QRY_EmailType.Text = rkutils.DataGridView_ValueByColumnName(DataGridView2, "LIVE_ATTACH", 0)
                Else
                    tEMAILto = Label_QRY_EMail.Text
                    tEMAILtoname = Label_QRY_EMail.Text
                End If
                tEMAILcc = ""
                tEMAILbcc = ""
                If Label_QRY_Description.Text.Length > 0 Then
                    tEMAILsubject = Label_QRY_Description.Text
                Else
                    tEMAILsubject = Label_QRY_Name.Text
                End If
                tEMAILattach = tReportName
                '**************************************************************
                '* 2013-08-23 RFK: Setup HTML
                tHTML = "<html><body>"
                tHTML += Label_QRY_EmailMessage.Text
                If swIncludeHTML Then
                    tHTML += tHTMLline + "</table>"
                End If
                tHTML += "<br><br>"
                tHTML += "Selected " + Label_QueriesAccounts.Text
                tHTML += "<br><br>"
                '**************************************************************
                '* 2013-09-09 RFK: 
                If Val(Label_QueriesAccounts.Text) > 0 Then tHTML += tReportName + "<br>"
                tHTML += "Type:" + Label_QRY_Type.Text + "<br>"
                tHTML += "Who:" + Label_QRY_Who.Text + "<br>"
                If Label_QRY_Status.Text.Trim.Length > 0 Then tHTML += "Statused:" + Label_QRY_Status.Text + "<br><br>"
                If Label_QRY_IncludeQuery.Text = "Y" Then tHTML += Label_QRY.Text   '* 2020-01-08 RFK:
                tHTML += "</body></html>"
                '**************************************************************
                '* 2013-09-09 RFK: 
                If Val(Label_QueriesAccounts.Text) > 0 Then
                    Select Case Label_QRY_EmailType.Text
                        Case "A"
                            MsgStatus("EMAIL [" + tEMAILto + "] Attach :" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, tEMAILattach)
                            End If
                            '**************************************************
                        Case "L"
                            MsgStatus("EMAIL [" + tEMAILto + "]  Link:" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, "")
                            End If
                            '**************************************************
                        Case "S"
                            MsgStatus("EMAIL [" + tEMAILto + "] Secure:" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                tEMAILmessage = tHTML
                                email_SECURE()
                            End If
                            '**************************************************
                    End Select
                Else
                    MsgStatus("EMAIL NO ATTACHMENT:" + tEMAILattach, True)
                    If tEMAILto.Contains("@") Then
                        If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, "")
                    End If
                    '**********************************************************
                End If
            End If
            '******************************************************************
            '* 2015-05-22 RFK:
            If Label_QRY_Output_CopyTo.Text.Length > 0 Then
                If Label_QRY_Output_CopyTo.Text.Contains(".") Then
                    MsgStatus("Output CopyTo:" + Label_QRY_Output_CopyTo.Text, True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo.Text)
                Else
                    If Label_QRY_Output_CopyTo.Text.Trim.EndsWith("\") = False Then
                        Label_QRY_Output_CopyTo.Text += "\"
                    End If
                    MsgStatus("Output CopyTo:" + Label_QRY_Output_CopyTo.Text + Path.GetFileName(tReportName), True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo.Text + Path.GetFileName(tReportName))
                End If
            End If
            '******************************************************************
            '* 2021-11-04 RFK:
            If Label_QRY_Output_CopyTo2.Text.Length > 0 Then
                If Label_QRY_Output_CopyTo2.Text.Contains(".") Then
                    MsgStatus("Output CopyTo2:" + Label_QRY_Output_CopyTo2.Text, True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo2.Text)
                Else
                    If Label_QRY_Output_CopyTo2.Text.Trim.EndsWith("\") = False Then
                        Label_QRY_Output_CopyTo2.Text += "\"
                    End If
                    MsgStatus("Output CopyTo2:" + Label_QRY_Output_CopyTo2.Text + Path.GetFileName(tReportName), True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo2.Text + Path.GetFileName(tReportName))
                End If
            End If
            '******************************************************************
            '* 2015-05-22 RFK:
            If Label_QRY_Output_CopyToEMailDir.Text = "Y" And tEMAILto.Contains("@") Then
                Dim dI As DirectoryInfo = New DirectoryInfo(Label_QRY_Output_CopyTo.Text)
                Dim allFiles() As FileInfo = dI.GetFiles
                tHTML = "<html><body>"
                tHTML += "<table>"
                For Each fl As FileInfo In allFiles
                    tHTML += "<tr>"
                    tHTML += "<td>" + fl.FullName.ToString + "</td><td>" + fl.CreationTime.ToString + "</td>"
                    tHTML += "</tr>"
                Next
                tHTML += "</table>"
                tHTML += "</body></html>"
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", "Directory of " + Label_QRY_Output_CopyTo.Text, "", tHTML, "")
            End If
            '******************************************************************
            '* 2015-01-22 RFK:
            Label_QueriesAccountsCount.Text = "0"
            Label_QRY_Running.Text = ""
            '******************************************************************
            '* 2012-10-12 RFK: 
            MsgStatus("Completed:" + Label_QueriesAccounts.Text, True)
            If Label_RUNNING.Text = "Running" Then Label_RUNNING.Text = "Ready"
        Catch ex As Exception
            MsgError("QRYrun", ex.ToString)
        End Try
    End Sub

    Private Sub QRYrun(ByVal sType As String, ByVal bLiveRun As Boolean)
        Try
            Dim tMessage As String = "", tSQL As String = "", sSQL As String = ""
            Dim iQRYLoop As Integer = 0
            Dim tReportName As String = ""
            Dim tHTML As String = "", tHTMLline As String = "", tTEMP As String = ""
            Dim tReportLine As String = "", tReportAllLines As String = "", tDelimiter As String = vbTab
            Dim sColumn As String = ""
            Dim iCol As Integer = 0, iCol2 As Integer = 0, iColTotal As Integer = 0, iColAverage As Integer = 0, iRow As Integer = 0, iSum As Integer = 0, iTotal As Integer = 0
            Dim swTotal As Boolean = False, swIncludeHTML As Boolean = False
            '******************************************************************
            '* 2015-05-22 RFK:
            Select Case Label_QRY_Output_Type.Text
                Case "C"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".CSV"
                    tDelimiter = ","
                    MsgStatus("QRYrun/Comma-CSV", True)
                Case "P"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".TXT"
                    tDelimiter = "|"
                    MsgStatus("QRYrun/Pipe", True)
                Case "E"
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".XLS"
                Case Else
                    tReportName = dir_REPORT + Label_QRY_Name.Text.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + "_" + rkutils.WhoAmI().Replace(".", "_") + ".XLS"
            End Select
            '*******************************************************************
            '* 2013-09-09 RFK: 
            MsgStatus("QRYrun:" + sType + " [" + Label_RUNNING.Text + "][" + Label_QueriesAccounts.Text + "][" + tReportName + "]", True)
            If File.Exists(tReportName) Then
                MsgStatus(tReportName + " already exists", True)
                Exit Sub
            End If
            '******************************************************************
            '* 2015-05-22 RFK: StreamWriter
            Dim sw As New StreamWriter(tReportName, False)
            '******************************************************************
            Label_QRY_Running.Text = sType + "-" + Label_RUNNING.Text
            '******************************************************************
            '* 2013-09-09 RFK: 
            'If Label_RUNNING.Text <> "Ready" Then Exit Sub
            Label_RUNNING.Text = "Running"
            Label_RUNNING.Refresh()
            System.Windows.Forms.Application.DoEvents()
            If Val(Label_QueriesAccounts.Text) > 0 Then
                '**************************************************************
                '* 2013-03-21 RFK:
                MsgStatus("Creating " + sType + " " + tReportName, True)
                MsgStatus(Label_RUNNING.Text, True)
                '**************************************************************
                '* 2015-05-22 RFK: Is it being included HTML
                swIncludeHTML = False
                Select Case Label_QRY_EmailType.Text
                    Case "A"    'Attached / Include
                        If Val(Label_QueriesAccountsCount.Text) < 1000 Then
                            swIncludeHTML = True
                        End If
                End Select
                '**************************************************************
                '* 2015-05-22 RFK: Does it contain Columns
                swTotal = True
                '**************************************************************
                For iCol = 0 To Me.DataGridView_QRYoutput.ColumnCount - 1
                    If Me.DataGridView_QRYoutput.Columns(iCol).Name IsNot Nothing Then
                        Select Case Me.DataGridView_QRYoutput.Columns(iCol).Name.ToUpper
                            Case "Total".ToUpper
                                swTotal = True
                        End Select
                    End If
                Next
                '**************************************************************
                If swIncludeHTML Then
                    tHTMLline = "<table width=100% cellpadding=2 cellspacing=2 border=1>"
                End If
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                Do While Label_RUNNING.Text = "Running" And iQRYLoop < Me.DataGridView_QRYoutput.RowCount - 1
                    rkutils.DoEvents()
                    tLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RALOCX", iQRYLoop).Trim
                    If tLOCX.Length = 0 Then tLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LOCX", iQRYLoop).Trim
                    'If tLOCX.Length = 0 Then MsgStatus("WARNING RALOCX IS BLANK", True)
                    '**********************************************************
                    tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iQRYLoop).Trim
                    If tBY.Length = 0 Then
                        tBY = rkutils.DataGridView_ValueByColumnName(DataGridView_COMMANDS, "MODIFIED_BY", iQRYLoop).Trim
                    End If
                    '**********************************************************
                    '* 2013-03-21 RFK:
                    '* 2013-10-22 RFK: Allow report if no LOCX in it
                    Select Case sType
                        Case "Report"
                            '**************************************************
                            '* 2012-10-12 RFK: The Header Row
                            tReportLine = ""
                            If iQRYLoop = 0 Then
                                If tLOCX.Length > 0 Then
                                    '* 2017-06-07 RFK:
                                    Select Case Label_QRY_Output_Type.Text
                                        Case "C"
                                            tReportLine += Chr(34) + "LOCX" + Chr(34) + tDelimiter
                                        Case "P"
                                            '*
                                        Case "E"
                                            tReportLine += "LOCX" + tDelimiter
                                        Case Else
                                            tReportLine += "LOCX" + tDelimiter
                                    End Select
                                End If
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "<tr>"
                                End If
                                '**********************************************
                                For iCol = 0 To Me.DataGridView_QRYoutput.ColumnCount - 1
                                    If Me.DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value IsNot Nothing Then
                                        sColumn = Me.DataGridView_QRYoutput.Columns(iCol).Name.ToString.ToUpper
                                        Select Case sColumn
                                            Case "TODAY"
                                                sColumn = rkutils.STR_format("TODAY", "mm-dd-ccyy")
                                            Case "DAY1"
                                                sColumn = rkutils.STR_format("TODAY-1", "DOW")
                                            Case "DAY2"
                                                sColumn = rkutils.STR_format("TODAY-2", "DOW")
                                            Case "DAY3"
                                                sColumn = rkutils.STR_format("TODAY-3", "DOW")
                                            Case "DAY4"
                                                sColumn = rkutils.STR_format("TODAY-4", "DOW")
                                            Case "DAY5"
                                                sColumn = rkutils.STR_format("TODAY-5", "DOW")
                                            Case "DAY6"
                                                sColumn = rkutils.STR_format("TODAY-6", "DOW")
                                            Case "DAY7"
                                                sColumn = rkutils.STR_format("TODAY-7", "DOW")
                                        End Select
                                        '**************************************
                                        '* 2017-06-07 RFK:
                                        Select Case Label_QRY_Output_Type.Text
                                            Case "C"
                                                tReportLine += Chr(34) + sColumn + Chr(34) + tDelimiter
                                            Case "P"
                                                '*
                                            Case Else
                                                tReportLine += sColumn + tDelimiter
                                        End Select
                                        '**************************************
                                        If swIncludeHTML Then
                                            tHTMLline += "<td>" + sColumn + "</td>"
                                        End If
                                        '**************************************
                                    End If
                                Next
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "</tr>"
                                End If
                                '**********************************************
                                '* 2015-04-22 RFK:
                                '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                                '* 2021-09-30 RFK: No Header in PIPE DELIMITED (NEED TO CREATE HEADER SWITCH)
                                Select Case Label_QRY_Output_Type.Text
                                    Case "C"
                                        sw.WriteLine(tReportLine)
                                    Case "P"
                                        '*
                                    Case Else
                                        sw.WriteLine(tReportLine)
                                End Select
                            End If
                            '**************************************************
                            '* 2012-10-12 RFK: Date Line
                            tReportLine = ""
                            If tLOCX.Length > 0 Then
                                '**************************************
                                '* 2017-06-07 RFK:
                                Select Case Label_QRY_Output_Type.Text
                                    Case "C"
                                        tReportLine += Chr(34) + tLOCX + Chr(34) + tDelimiter
                                    Case "P"
                                        '* NO 
                                    Case Else
                                        tReportLine += tLOCX + tDelimiter
                                End Select
                            End If
                            '**************************************************
                            If swIncludeHTML Then
                                tHTMLline += "<tr>"
                            End If
                            '**************************************************
                            '* 2015-05-21 RFK:
                            If swIncludeHTML Or swTotal Then
                                For iCol = 0 To Me.DataGridView_QRYoutput.ColumnCount - 1
                                    If Me.DataGridView_QRYoutput.Columns(iCol).Name IsNot Nothing Then
                                        Select Case Me.DataGridView_QRYoutput.Columns(iCol).Name.ToUpper
                                            Case "Total".ToUpper
                                                '**********************************
                                                swTotal = True
                                                iColTotal = iCol
                                                iSum = 0
                                                For iCol2 = 1 To iCol - 1
                                                    If Me.DataGridView_QRYoutput.Item(iCol2, iQRYLoop).Value IsNot Nothing Then
                                                        iSum += Val(DataGridView_QRYoutput.Item(iCol2, iQRYLoop).Value.ToString)
                                                    End If
                                                Next
                                                '**********************************
                                                If Me.DataGridView_QRYoutput.Item(iCol2, iQRYLoop).Value IsNot Nothing Then
                                                    '**************************************
                                                    '* 2017-06-07 RFK:
                                                    Select Case Label_QRY_Output_Type.Text
                                                        Case "C"
                                                            tReportLine += Chr(34) + DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString.Trim + Chr(34)
                                                        Case Else
                                                            tReportLine += DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString
                                                    End Select
                                                End If
                                                tReportLine += tDelimiter
                                                '**********************************
                                                If swIncludeHTML Then
                                                    tHTMLline += "<td>"
                                                    If Me.DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value IsNot Nothing Then
                                                        tHTMLline += Me.DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString
                                                    End If
                                                    tHTMLline += "</td>"
                                                End If
                                                '**********************************
                                            Case "Average".ToUpper
                                                '**********************************
                                                iColAverage = iCol
                                                iSum = 0
                                                For iCol2 = 1 To iCol - 1
                                                    If Me.DataGridView_QRYoutput.Item(iCol2, iQRYLoop).Value IsNot Nothing Then
                                                        iSum += Val(DataGridView_QRYoutput.Item(iCol2, iQRYLoop).Value.ToString())
                                                    End If
                                                Next
                                                '**********************************
                                                '******************************
                                                '* 2017-06-07 RFK:
                                                Select Case Label_QRY_Output_Type.Text
                                                    Case "C"
                                                        tReportLine += Chr(34) + rkutils.STR_format((iSum / (iCol - 1)).ToString, "0").Trim + Chr(34) + tDelimiter
                                                    Case Else
                                                        tReportLine += rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + tDelimiter
                                                End Select
                                                '**********************************
                                                If swIncludeHTML Then
                                                    tHTMLline += "<td>" + rkutils.STR_format((iSum / (iCol - 1)).ToString, "0") + "</td>"
                                                End If
                                                '**********************************
                                            Case Else
                                                If Me.DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value IsNot Nothing Then
                                                    '******************************
                                                    '* 2017-06-07 RFK:
                                                    Select Case Label_QRY_Output_Type.Text
                                                        Case "C"
                                                            tReportLine += Chr(34) + DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString.Trim + Chr(34) + tDelimiter
                                                        Case Else
                                                            tReportLine += DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString + tDelimiter
                                                    End Select
                                                    '******************************
                                                    If swIncludeHTML Then
                                                        tHTMLline += "<td>" + Me.DataGridView_QRYoutput.Item(iCol, iQRYLoop).Value.ToString + "</td>"
                                                    End If
                                                    '**********************************************
                                                End If
                                        End Select
                                    End If
                                Next
                                '**********************************************
                                If swIncludeHTML Then
                                    tHTMLline += "</tr>"
                                End If
                            End If
                            '**************************************************
                            '* 2015-04-22 RFK: 
                            '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                            sw.WriteLine(tReportLine)
                            '**************************************************
                        Case "CallList"
                            MsgStatus("CallList should NOT get to this point in code", True)
                    End Select
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '**********************************************************
                    '* 2013-10-22 RFK: Only if a LOCX in it
                    If tLOCX.Length > 0 Then
                        '******************************************************
                        '* 2013-04-03 RFK:
                        If Label_QRY_Letter_Number.Text.Trim.Length > 0 Then
                            If Val(Label_QRY_Letter_Number.Text) > 0 And Val(Label_QRY_Letter_Number.Text) <= 999 Then
                                Select Case Label_QRY_Letter_Date.Text
                                    Case "TODAY"
                                        Label_QRY_Letter_Date.Text = STR_format("TODAY", "mm/dd/ccyy")
                                    Case Else
                                End Select
                                If IsDate(Label_QRY_Letter_Date.Text) Then
                                    tSQL = "UPDATE ROIDATA.RACCTP"
                                    tSQL += " SET RALNAC=" + Label_QRY_Letter_Number.Text
                                    tSQL += ",RANLMO=" + STR_format(Label_QRY_Letter_Date.Text, "mm")
                                    tSQL += ",RANLDY=" + STR_format(Label_QRY_Letter_Date.Text, "dd")
                                    tSQL += ",RANLYR=" + STR_format(Label_QRY_Letter_Date.Text, "ccyy")
                                    tSQL += " WHERE RALOCX='" + tLOCX + "'"
                                    rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, tSQL)
                                    MsgStatus(tSQL, False)
                                    '******************************************
                                    tMessage = "QRY:" + Label_QRY_Name.Text + " set letter to:" + Label_QRY_Letter_Number.Text + " [" + Label_QRY_Letter_Date.Text + "]"
                                    rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", tMessage, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                                Else
                                    MsgStatus(Label_QRY_Letter_Date.Text + " is NOT a valid date, can not set to letter " + Label_QRY_Letter_Number.Text, True)
                                End If
                            End If
                        End If
                        '******************************************************
                        '* 2012-10-12 RFK:
                        If Label_QRY_Status.Text.Trim.Length > 0 Then
                            '**************************************************
                            '* 2013-08-22 RFK: TCode/Status if checkbox allowed.
                            If CheckBox_Queries_TCode.Checked Then
                                '**********************************************
                                '* 2015-05-21 RFK:
                                rkutils.COMMAND_STATUS(msSQLConnectionString, msSQLuser, tLOCX, Label_QRY_Status.Text, "", "", "")
                                '**********************************************
                                '* 2013-08-22 RFK: TCode/Status Matched Accounts.
                                'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", "QRY[" + Label_QRY_tUnique.Text + "] " + Label_QRY_Name.Text + " STATUS:" + Label_QRY_Status.Text, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                            Else
                                MsgStatus("Status Account [Not checked to run]" + tLOCX + " " + Label_QRY_Status.Text, True)
                            End If
                        End If
                        '******************************************************
                        '* 2019-03-18 RFK: Set Field/Value
                        'MsgStatus("Label_SetField.Text=" + Label_SetField.Text, False)
                        If Label_SetField.Text.Trim.Length > 0 And Label_SetValue.Text.Trim.Length > 0 And CheckBox_Queries_TCode.Checked Then
                            '**************************************************
                            sSQL = "UPDATE ROIDATA.RACCTP SET " + Label_SetField.Text + "='" + Label_SetValue.Text + "'"
                            sSQL += rkutils.WhereAnd(sSQL, "RALOCX='" + tLOCX + "'")
                            MsgStatus(sSQL, False)
                            rkutils.DB_COMMAND("DB2", DB2SQLConnectionString, DB2SQLuser, sSQL)
                            '**************************************************
                            rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", "QRY[" + Label_QRY_tUnique.Text + "] " + Label_QRY_Name.Text + " SET:" + Label_SetField.Text + "=" + Label_SetValue.Text, STR_format("TODAY", "mm/dd/ccyy HH:MM:SS"), "ePROC")
                            '**************************************************
                        End If
                        '******************************************************
                    End If
                    iQRYLoop += 1
                    Label_QueriesAccountsCount.Text = Trim(Str(Val(Label_QueriesAccountsCount.Text) - 1))
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                    '**************************************************************
                Loop
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '**************************************************************
                '* 2015-05-14 RFK: Sum Columns
                If swTotal = True And iColTotal > 0 Then
                    '******************************
                    '* 2017-06-07 RFK:
                    Select Case Label_QRY_Output_Type.Text
                        Case "C"
                            tReportLine = Chr(34) + "Total" + Chr(34) + tDelimiter
                        Case Else
                            tReportLine = "Total" + tDelimiter
                    End Select
                    '**********************************************************
                    Select Case Label_QRY_EmailType.Text
                        Case "A"    'Attached / Include
                            tHTMLline += "<tr>"
                            tHTMLline += "<td>Total</td>"
                    End Select
                    '**********************************************************
                    iTotal = 0
                    For iCol = 1 To iColTotal - 1
                        iSum = 0
                        For iRow = 0 To Me.DataGridView_QRYoutput.RowCount - 1
                            If Me.DataGridView_QRYoutput.Item(iCol, iRow).Value IsNot Nothing Then
                                iSum += Val(DataGridView_QRYoutput.Item(iCol, iRow).Value.ToString())
                            End If
                        Next
                        '******************************************************
                        iTotal += iSum
                        '******************************************************
                        '******************************
                        '* 2017-06-07 RFK:
                        Select Case Label_QRY_Output_Type.Text
                            Case "C"
                                tReportLine += Chr(34) + iSum.ToString + Chr(34) + tDelimiter
                            Case Else
                                tReportLine += iSum.ToString + tDelimiter
                        End Select
                        '******************************************************
                        Select Case Label_QRY_EmailType.Text
                            Case "A"    'Attached / Include
                                tHTMLline += "<td>" + iSum.ToString + "</td>"
                        End Select
                    Next
                    '**********************************************************
                    '******************************
                    '* 2017-06-07 RFK:
                    Select Case Label_QRY_Output_Type.Text
                        Case "C"
                            tReportLine += Chr(34) + iTotal.ToString + Chr(34) + tDelimiter
                        Case Else
                            tReportLine += iTotal.ToString + tDelimiter
                    End Select
                    '**********************************************************
                    If iColAverage > 0 Then
                        '******************************
                        '* 2017-06-07 RFK:
                        Select Case Label_QRY_Output_Type.Text
                            Case "C"
                                tReportLine += Chr(34) + (iTotal / (iColTotal - 1)).ToString + Chr(34) + tDelimiter
                            Case Else
                                tReportLine += (iTotal / (iColTotal - 1)).ToString + tDelimiter
                        End Select
                    End If
                    '**********************************************************
                    '* 2015-05-22 RFK: Changed to sw 'tReportAllLines += tReportLine + vbCrLf
                    sw.WriteLine(tReportLine)
                    sw.Close()
                    '**********************************************************
                    Select Case Label_QRY_EmailType.Text
                        Case "A"    'Attached / Include
                            tHTMLline += "<td>" + iTotal.ToString + "</td>"
                            '**********************************************************
                            If iColAverage > 0 Then
                                tHTMLline += "<td>" + (iTotal / (iColTotal - 1)).ToString + "</td>"
                            End If
                            '**********************************************************
                            tHTMLline += "</tr>"
                    End Select
                End If
                '**************************************************************
                '* 2015-04-22 RFK: 
                'If tReportAllLines.Length > 0 Then
                '    File.AppendAllText(tReportName, tReportAllLines)
                'End If
                '**************************************************************
            Else
                MsgStatus("QRYrun:" + Label_QRY_Name.Text + " selected NOTHING", True)
            End If
            '******************************************************************
            '* 2015-11-09 RFK: Close the streamwriter
            If sw IsNot Nothing Then sw.Close()
            '******************************************************************
            '* 2012-10-12 RFK: Email
            '* 2015-03-30 RFK: Blank (No Results) Email
            MsgStatus("Label_QueriesAccounts.Text=" + Label_QueriesAccounts.Text, True)
            MsgStatus("Label_QRY_Emailblank.Text =" + Label_QRY_Emailblank.Text, True)
            If Val(Label_QueriesAccounts.Text) > 0 Or Label_QRY_Emailblank.Text = "Y" Then
                MsgStatus("Emailing" + sType + " " + tReportName, True)
                '**************************************************************
                tEMAILfrom = "production@AnnuityHealth.com"
                tEMAILfromname = "Production"
                '**************************************************************
                '* 2013-02-27 RFK: Email
                If bLiveRun Then
                    '**********************************************************
                    '* 2015-07-09 RFK: 
                    tSQL = "SELECT LIVE_EMAIL,LIVE_ATTACH"
                    Select Case sSITE
                        Case "AnnuityOne"
                            tSQL += " FROM RevMD.dbo.query "
                        Case Else
                            tSQL += " FROM iTeleCollect.dbo.query"
                    End Select
                    '**********************************************************
                    tSQL += " WHERE TUNIQUE='" + Label_QRY_tUnique.Text + "'"
                    '**********************************************************
                    tEMAILto = rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "LIVE_EMAIL", msSQLConnectionString, msSQLuser, tSQL)
                    tEMAILtoname = tEMAILto
                    Label_QRY_EmailType.Text = rkutils.DataGridView_ValueByColumnName(DataGridView2, "LIVE_ATTACH", 0)
                Else
                    tEMAILto = Label_QRY_EMail.Text
                    tEMAILtoname = Label_QRY_EMail.Text
                End If
                tEMAILcc = ""
                tEMAILbcc = ""
                If Label_QRY_Description.Text.Length > 0 Then
                    tEMAILsubject = Label_QRY_Description.Text
                Else
                    tEMAILsubject = Label_QRY_Name.Text
                End If
                tEMAILattach = tReportName
                '**************************************************************
                '* 2013-08-23 RFK: Setup HTML
                tHTML = "<html><body>"
                tHTML += Label_QRY_EmailMessage.Text
                If swIncludeHTML Then
                    tHTML += tHTMLline + "</table>"
                End If
                tHTML += "<br><br>"
                tHTML += "Selected " + Label_QueriesAccounts.Text
                tHTML += "<br><br>"
                '**************************************************************
                '* 2013-09-09 RFK: 
                If Val(Label_QueriesAccounts.Text) > 0 Then tHTML += tReportName + "<br>"
                tHTML += "Type:" + Label_QRY_Type.Text + "<br>"
                tHTML += "Who:" + Label_QRY_Who.Text + "<br>"
                If Label_QRY_Status.Text.Trim.Length > 0 Then tHTML += "Statused:" + Label_QRY_Status.Text + "<br><br>"
                If Label_QRY_IncludeQuery.Text = "Y" Then tHTML += Label_QRY.Text   '* 2020-01-08 RFK:
                tHTML += "</body></html>"
                '**************************************************************
                '* 2013-09-09 RFK: 
                If Val(Label_QueriesAccounts.Text) > 0 Then
                    Select Case Label_QRY_EmailType.Text
                        Case "A"
                            MsgStatus("EMAIL [" + tEMAILto + "] Attach :" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, tEMAILattach)
                            End If
                            '**************************************************
                        Case "L"
                            MsgStatus("EMAIL [" + tEMAILto + "]  Link:" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, "")
                            End If
                            '**************************************************
                        Case "S"
                            MsgStatus("EMAIL [" + tEMAILto + "] Secure:" + tEMAILattach, True)
                            If tEMAILto.Contains("@") Then
                                If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                                tEMAILmessage = tHTML
                                email_SECURE()
                            End If
                            '**************************************************
                    End Select
                Else
                    MsgStatus("EMAIL NO ATTACHMENT:" + tEMAILattach, True)
                    If tEMAILto.Contains("@") Then
                        If swTEST Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                        rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", tEMAILsubject, "", tHTML, "")
                    End If
                    '**********************************************************
                End If
            End If
            '******************************************************************
            '* 2015-05-22 RFK:
            If Label_QRY_Output_CopyTo.Text.Length > 0 Then
                If Label_QRY_Output_CopyTo.Text.Contains(".") Then
                    MsgStatus("Output CopyTo:" + Label_QRY_Output_CopyTo.Text, True)
                    If File.Exists(Label_QRY_Output_CopyTo.Text) Then
                        MsgStatus("Already Exists:" + Label_QRY_Output_CopyTo.Text, True)
                        File.Move(Label_QRY_Output_CopyTo.Text, Path.GetFileNameWithoutExtension(Label_QRY_Output_CopyTo.Text) + "_" + rkutils.STR_format("TODAY", "HHMMSS") + rkutils.STR_RIGHT(Label_QRY_Output_CopyTo.Text, 4))
                    End If
                    File.Copy(tReportName, Label_QRY_Output_CopyTo.Text)
                Else
                    If Label_QRY_Output_CopyTo.Text.Trim.EndsWith("\") = False Then
                        Label_QRY_Output_CopyTo.Text += "\"
                    End If
                    MsgStatus("Output CopyTo:" + Label_QRY_Output_CopyTo.Text + Path.GetFileName(tReportName), True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo.Text + Path.GetFileName(tReportName))
                End If
            End If
            '******************************************************************
            '* 2021-11-04 RFK:
            If Label_QRY_Output_CopyTo2.Text.Length > 0 Then
                If Label_QRY_Output_CopyTo2.Text.Contains(".") Then
                    MsgStatus("Output CopyTo2:" + Label_QRY_Output_CopyTo2.Text, True)
                    If File.Exists(Label_QRY_Output_CopyTo2.Text) Then
                        MsgStatus("Already Exists:" + Label_QRY_Output_CopyTo2.Text, True)
                        File.Move(Label_QRY_Output_CopyTo2.Text, Path.GetFileNameWithoutExtension(Label_QRY_Output_CopyTo2.Text) + "_" + rkutils.STR_format("TODAY", "HHMMSS") + rkutils.STR_RIGHT(Label_QRY_Output_CopyTo2.Text, 4))
                    End If
                    File.Copy(tReportName, Label_QRY_Output_CopyTo2.Text)
                Else
                    If Label_QRY_Output_CopyTo2.Text.Trim.EndsWith("\") = False Then
                        Label_QRY_Output_CopyTo2.Text += "\"
                    End If
                    MsgStatus("Output CopyTo2:" + Label_QRY_Output_CopyTo2.Text + Path.GetFileName(tReportName), True)
                    File.Copy(tReportName, Label_QRY_Output_CopyTo2.Text + Path.GetFileName(tReportName))
                End If
            End If
            '******************************************************************
            '* 2015-05-22 RFK:
            If Label_QRY_Output_CopyToEMailDir.Text = "Y" And tEMAILto.Contains("@") Then
                Dim dI As DirectoryInfo = New DirectoryInfo(Label_QRY_Output_CopyTo.Text)
                Dim allFiles() As FileInfo = dI.GetFiles
                tHTML = "<html><body>"
                tHTML += "<table>"
                For Each fl As FileInfo In allFiles
                    tHTML += "<tr>"
                    tHTML += "<td>" + fl.FullName.ToString + "</td><td>" + fl.CreationTime.ToString + "</td>"
                    tHTML += "</tr>"
                Next
                tHTML += "</table>"
                tHTML += "</body></html>"
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, tEMAILfrom, tEMAILfromname, tEMAILto, tEMAILtoname, "", "Directory of " + Label_QRY_Output_CopyTo.Text, "", tHTML, "")
            End If
            '******************************************************************
            '* 2015-01-22 RFK:
            Label_QueriesAccountsCount.Text = "0"
            Label_QRY_Running.Text = ""
            '******************************************************************
            '* 2012-10-12 RFK: 
            MsgStatus("Completed:" + Label_QueriesAccounts.Text, True)
            If Label_RUNNING.Text = "Running" Then Label_RUNNING.Text = "Ready"
        Catch ex As Exception
            MsgError("QRYrun", ex.ToString)
        End Try
    End Sub

    Private Sub email_SECURE()
        Try
            '******************************************************************
            If tEMAILto.Contains("@") = False Then
                MsgStatus("email_SECURE can not send to:" + tEMAILto, True)
                Exit Sub
            End If
            If File.Exists(tEMAILattach) = False Then
                MsgStatus("email_SECURE can not attach :" + tEMAILattach, True)
                Exit Sub
            End If
            '******************************************************************
            MsgStatus("email_SECURE From:" + tEMAILfrom + " To:" + tEMAILto + " Subject:" + tEMAILsubject, True)
            Dim sSendString As String = tEMAILto
            sSendString += " " + Chr(34) + tEMAILcc + Chr(34)
            sSendString += " " + Chr(34) + tEMAILbcc + Chr(34)
            sSendString += " " + Chr(34) + tEMAILsubject + Chr(34)
            sSendString += " " + Chr(34) + tEMAILmessage + Chr(34)
            sSendString += " " + Chr(34) + tEMAILattach + Chr(34)
            '******************************************************************
            '* 2018-12-04 RFK: Modified iAppRiverSend.EXE
            If IS_File(dir_EMAIL + "aoSecureEMail\aoSecureEMail.EXE") Then
                MsgStatus(dir_EMAIL + "aoSecureEMail\aoSecureEMail " + sSendString, True)
                Shell(dir_EMAIL + "aoSecureEMail\aoSecureEMail" + " " + sSendString, AppWinStyle.NormalNoFocus)
            Else
                MsgStatus("ERROR:" + dir_EMAIL + "aoSecureEMail\aoSecureEMail " + sSendString, True)
            End If
        Catch ex As Exception
            MsgError("email_SECURE", ex.ToString)
        End Try
    End Sub

    Private Sub CallList_ADD(ByVal tCallList As String, ByVal sDescription As String, ByVal sGroup As String, ByVal sDialer As String, ByVal sAppType As String, ByVal sRatio As String, ByVal sStart As String, ByVal sStop As String, ByVal sInsert As String, ByVal sAccountNumberOutput As String, ByVal sClientOutput As String)
        '************************************************
        '* 2011-11-01 RFK:
        Try
            Dim iQRYLoop As Integer = 0
            Dim tPhone As String = "", tPhoneFlag As String = "", tLOCX As String = ""
            Dim tClient As String = "", tFacility As String = "", tFacilityName As String = "", tFacilityGroup As String = "", tTOB As String = ""
            Dim tClientVMB As String = ""
            Dim sCompany As String = "" '* 2021-11-04 RFK:
            Dim sCallListDirectory As String = "" '* 2021-11-05 RFK:
            Dim tBalance As String = "", tDOS As String = ""
            Dim SQLcommandstring As String = "", ReportString As String = ""
            Dim sDialerString As String = ""
            Dim iAdded As Integer = 0, iNOTphone As Integer = 0
            Dim tTimeZone As String = "", tState As String = "", tRAMTTP As String = ""
            Dim iAreaCodeRow As Integer = 0, iStateBlocked As Integer = 0, iClient As Integer = 0
            Dim swOK As Boolean
            ListBox_Phones.Items.Clear()   'Used For Phones in this build
            ListBox_Dialer.Items.Clear()
            '******************************************************************
            '* 2015-01-26 RFK:
            tCallList = tCallList.Replace("-", "_")
            tCallList = tCallList.Replace(" ", "_")
            '******************************************************************
            '* 2015-12-30 RFK: 
            MsgStatus("CallList:" + tCallList + " Description:" + sDescription + " Group:" + sGroup, True)
            MsgStatus("Dialer:" + sDialer + " AppType:" + sAppType, True)
            MsgStatus("Ratio:" + sRatio + " Start:" + sStart + " Stop:" + sStop + " Insert:" + sInsert, True)
            MsgStatus("AccountNumberOutput:" + sAccountNumberOutput, True)
            MsgStatus("ClientOutput:" + sClientOutput, True)
            '******************************************************************
            '* 2017-08-22 RFK: 
            If DataGridView_QRYoutput.RowCount - 1 < 1 Then
                MsgStatus("CallList:" + tCallList + " No Records Selected", True)
                '**********************************************************
                '* 2017-09-14 RFK:
                MsgStatus("EmailTo:" + tEMAILto + "-" + tEMAILtoname, True)
                tEMAILto = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYs, "EMAIL", 0).Trim
                MsgStatus("EmailTo:" + tEMAILto, True)
                If tEMAILto.Contains("@") = False Then tEMAILto = "Ryan.Kiechle@AnnuityHealth.com"
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "aoProcessor", tEMAILto, tEMAILto, "", "CallList:" + tCallList + "/No Records Selected", tCallList + " did NOT select any records [" + tClient + "]", "", "")
                Exit Sub
            End If
            '******************************************************************
            '* 2017-09-14 RFK:
            MsgStatus("Contains:" + DataGridView_QRYoutput.RowCount.ToString, True)
            '******************************************************************
            '* 2012-09-20 RFK: First Record Default for Client
            tClient = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RACL#", 1).Trim
            MsgStatus("Client:" + tClient, True)
            tClientVMB = rkutils.WhatIsClientVMBValue(DataGridView2, msSQLConnectionString, msSQLuser, tClient, "", "", "VMB_MSG")
            MsgStatus("VMB:" + tClientVMB, True)
            If Val(tClientVMB) < 1000 Then
                '**********************************************************
                '* 2017-07-27 RFK:
                MsgStatus(tCallList + " does NOT have a VMB defined in dbo.clientsVMB [" + tClient + "]", True)
                rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "aoProcessor", "HELPDESK@AnnuityHealth.com", "HelpDesk", "", "CallList_Add/ERROR", tCallList + " does NOT have a VMB defined in dbo.clientsVMB [" + tClient + "]", "", "")
                Exit Sub
                '**********************************************************
            End If
            '************************************
            If CallList_exists(tCallList) = False Then
                '**************************************************************
                SQLcommandstring = "INSERT INTO TeleServer.dbo.CallLists"
                SQLcommandstring += "(apptype, callList, description"
                SQLcommandstring += ", tgroup"
                SQLcommandstring += ", torder"
                SQLcommandstring += ", dialahead"
                SQLcommandstring += ", timestart, timestop"
                SQLcommandstring += ", vmb, vmbcid, vmbmsg"
                SQLcommandstring += ", modified_date, modified_by)"
                SQLcommandstring += "values('" + sAppType + "'"
                SQLcommandstring += ", '" + STR_LEFT(tCallList, 20) + "'"
                SQLcommandstring += ", '" + sDescription + "'"
                SQLcommandstring += ", '" + sGroup + "'"
                SQLcommandstring += ", 'A'"
                SQLcommandstring += ", '" + sRatio + "'"
                SQLcommandstring += ", '" + sStart + "'"
                SQLcommandstring += ", '" + sStop + "'"
                SQLcommandstring += ", '" + tClientVMB + "'"
                SQLcommandstring += ", '" + tClientVMB + "'"
                SQLcommandstring += ", '" + tClientVMB + "'"
                SQLcommandstring += ", '" + Date.Now + "'"
                SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
                SQLcommandstring += ")" + vbCr
                '2015-01-23 RFK: 
                If swTEST = False Then rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
                SQLcommandstring = ""
                '*************************************************************************************************************
                '2015-01-23 RFK: 
                If swTEST = False Then
                    Label_Unique_Name.Text = rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "TUNIQUE", msSQLConnectionString, msSQLuser, "SELECT TUNIQUE FROM TeleServer.dbo.calllists WHERE CALLLIST='" + tCallList + "'")
                Else
                    Label_Unique_Name.Text = "1"
                End If
            End If
            '**************************************************************************************
            If Val(Label_Unique_Name.Text) = 0 Then
                MsgStatus("Can NOT build CallList [TUNIQUE]", True)
                Exit Sub
            End If
            '**************************************************************************************
            Dim iRecNo As Integer = Val(rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "MAXNO", msSQLConnectionString, msSQLuser, "SELECT MAX(TUNIQUE) AS MAXNO FROM TeleServer.dbo.calllist WHERE CALLLIST='" + tCallList + "'"))
            '**************************************************************************************
            '* 2011-07-01 RFK: add to a calllist
            '* 2011-10-01 RFK: If already dialed today do not add
            '* 2012-01-11 RFK: If already in this calllist do not add (GridViewPhone)
            '* 2012-11-27 RFK: If already in UNIQUE_NAME then do not add (GridViewPhone)
            rkutils.SQL_READ_FIELD(DataGridView_Phone, "MSSQL", "PHONE", msSQLConnectionString, msSQLuser, "SELECT PHONE FROM TeleServer.dbo.CallList WHERE UNIQUE_NAME='" + Label_Unique_Name.Text + "'")
            '**************************************************************************************
            '* 2012-11-12 RFK: AreaCode
            rkutils.SQL_READ_FIELD(DataGridView_AreaCode, "MSSQL", "AREACODE", msSQLConnectionString, msSQLuser, "SELECT AreaCode,Exchange,Time_Zone,Call_Type,state,block_call,block_messages FROM TeleServer.dbo.areacode")
            '**************************************************************************************
            '* 2012-11-12 RFK: State Blocking
            rkutils.SQL_READ_FIELD(DataGridView_StateBlock, "DB2", "STATE", DB2SQLConnectionString, DB2SQLuser, "SELECT PostalCode,Call_Active,Call_Collections,UPPER(sGroup) AS sGroup FROM ROIDATA.StateBlocking WHERE CALL_Active='N' or CALL_Collections='N' ORDER BY SGROUP, POSTALCODE")
            '**************************************************************************************
            '* 2021-11-04 RFK: Client / Company-Agency
            rkutils.SQL_READ_DATAGRID(DataGridView_Clients, "MSSQL", "ClientName", msSQLConnectionString, msSQLuser, "SELECT ClientName, Company FROM RevMD.dbo.Clients WHERE Active='Y' ORDER BY ClientName")
            '**************************************************************************************
            '* 2012-12-21 RFK: GHOST Calls
            '* 2012-12-21 RFK: Read 1st Client info from actual calllist
            tClient = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RACL#", 1).Trim
            tTOB = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RATOB", 1).Trim
            tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", 1).Trim
            If tFacility.Length = 0 Then rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAFACL", 1).Trim()
            tFacilityGroup = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACGRP", 1).Trim
            tClientVMB = rkutils.WhatIsClientVMBValue(DataGridView2, msSQLConnectionString, msSQLuser, tClient, tTOB, tFacility, "VMB_MSG")
            '**************************************************************************************
            '* 2012-12-21 RFK: GHOST Calls
            rkutils.SQL_READ_FIELD(DataGridView2, "MSSQL", "*", msSQLConnectionString, msSQLuser, "SELECT * FROM TeleServer.dbo.CallsGhost WHERE TGROUP='*' OR TGROUP='CALL'")
            For i1 = 0 To Me.DataGridView2.Rows.Count - 1
                tPhone = rkutils.DataGridView_ValueByColumnName(DataGridView2, "PHONE", i1).Trim
                MsgStatus(tPhone, True)
                If tPhone.Length = 10 Then
                    '******************************************************************************
                    '* 2013-01-08 RFK: Not Already in the CallList
                    If rkutils.DataGridViewContains(DataGridView_Phone, "PHONE", tPhone) = -1 Then
                        '*****************************************************************************************************
                        '* 2013-01-08 RFK: Not Already in this list
                        If rkutils.Listbox_Contains(ListBox_Phones, tPhone, False) = False Then
                            iRecNo += 1
                            tTimeZone = rkutils.DataGridView_ValueByColumnName(DataGridView2, "TIMEZONE", i1).Trim
                            '*********************************************************************************************
                            SQLcommandstring = "INSERT INTO TeleServer.dbo.CallList"
                            SQLcommandstring += " (tunique"
                            SQLcommandstring += ", calllist"
                            SQLcommandstring += ", phone"
                            SQLcommandstring += ", timezone"
                            SQLcommandstring += ", vmb"
                            SQLcommandstring += ", client"
                            SQLcommandstring += ", account"
                            SQLcommandstring += ", contact"
                            SQLcommandstring += ", route"
                            SQLcommandstring += ", ref1"
                            SQLcommandstring += ", ref2"
                            SQLcommandstring += ", ref3"
                            SQLcommandstring += ", ref4"
                            SQLcommandstring += ", unique_name"
                            SQLcommandstring += ", modified_date, modified_by)"
                            SQLcommandstring += " values('" + Str(iRecNo).Trim + "'"
                            SQLcommandstring += ", '" + rkutils.STR_LEFT(tCallList, 20) + "'"
                            SQLcommandstring += ", '" + tPhone + "'"
                            SQLcommandstring += ", '" + tTimeZone + "'"
                            SQLcommandstring += ", '" + tClientVMB + "'"
                            SQLcommandstring += ", '" + tClient + "'"
                            SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView2, "ACCOUNTNUMBER", i1).Trim + "'"
                            SQLcommandstring += ", '" + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView2, "FIRSTNAME", i1).Trim) + " " + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView2, "LASTNAME", i1).Trim) + "'"
                            SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView2, "ACCOUNTNUMBER", i1).Trim + "'"
                            SQLcommandstring += ", ''"    'Ref1
                            SQLcommandstring += ", ''"    'Ref2
                            SQLcommandstring += ", ''"    'Ref3
                            SQLcommandstring += ", ''"    'Ref4
                            SQLcommandstring += ", " + Val(Label_Unique_Name.Text).ToString.Trim + ""
                            SQLcommandstring += ", '" + Date.Now + "'"
                            SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
                            SQLcommandstring += ")" + vbCr
                            '***************************************
                            '2015-01-23 RFK: 
                            If swTEST = False Then rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
                            '***************************************
                            '* 2013-01-08 RFK: In this list now
                            ListBox_Phones.Items.Add(tPhone)
                            '***************************************
                        End If
                    End If
                End If
            Next
            '******************************************************************
            '* 2013-05-17 RFK:
            '* 2015-01-23 RFK: 
            Dim tReportName As String = dir_REPORT + "CLST_" + tCallList.Replace(" ", "_").Trim + "_" + STR_format("TODAY", "ccyymmdd_HHMM") + ".XLS"
            '******************************************************************
            Dim tDelimiter As String = vbTab
            ReportString = "CallListName"
            ReportString += tDelimiter + "LOCX"
            ReportString += tDelimiter + "Phone"
            ReportString += tDelimiter + "TimeZone"
            ReportString += tDelimiter + "VMB"
            ReportString += tDelimiter + "Client"
            ReportString += tDelimiter + "AccountNumber"
            ReportString += tDelimiter + "Name"
            ReportString += tDelimiter + "Balance"
            File.AppendAllText(tReportName, ReportString + vbCrLf)
            '**************************************************************************************
            '* 2013-03-21 RFK: Changed to CheckBox_Queries
            iQRYLoop = 0
            Do While iQRYLoop < DataGridView_QRYoutput.RowCount 'and CheckBox_Queries.checked 
                '**********************************************************************************
                '* 2011-12-12 RFK: Verify it is within ranges
                '**********************************************************************************
                tLOCX = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RALOCX", iQRYLoop).Trim
                If tLOCX.Length > 0 Then
                    '******************************************************************************
                    '* 2012-11-12 RFK: STATE BLOCKED
                    '* 2012-12-18 RFK: look at A / C seperately 
                    tState = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAGSTATE", iQRYLoop).Trim
                    tRAMTTP = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAMTTP", iQRYLoop).Trim
                    '******************************************************************************
                    '* 2021-11-04 RFK: COMPANY (AGENCY) SGROUP
                    tClient = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RACL#", iQRYLoop).Trim
                    iClient = rkutils.DataGridViewContains(DataGridView_Clients, "CLIENTNAME", tClient)
                    If iClient >= 0 Then
                        sCompany = rkutils.DataGridView_ValueByColumnName(DataGridView_Clients, "COMPANY", iClient).Trim.ToUpper
                    Else
                        sCompany = "INVALID"
                    End If
                    swOK = True
                    If swOK Then
                        '**************************************************************************
                        '* 2012-11-12 RFK: STATE BLOCKED
                        '* 2021-11-04 RFK: COMPANY (AGENCY) SGROUP
                        iStateBlocked = rkutils.DataGridViewContains2(DataGridView_StateBlock, "POSTALCODE", tState, "SGROUP", sCompany, True)
                        If iStateBlocked >= 0 Then
                            MsgStatus(tState + " " + sCompany + " " + tRAMTTP, True)
                            Select Case tRAMTTP
                                Case "A"
                                    If rkutils.DataGridView_ValueByColumnName(DataGridView_StateBlock, "CALL_ACTIVE", iStateBlocked).Trim = "N" Then
                                        MsgStatus("STATE BLOCKED [" + tState + "][" + sCompany + "]", True)
                                        CallList_Error(tCallList, tState, "STATE BLOCKED")
                                        swOK = False
                                    End If
                                Case "C"
                                    If rkutils.DataGridView_ValueByColumnName(DataGridView_StateBlock, "CALL_COLLECTIONS", iStateBlocked).Trim = "N" Then
                                        MsgStatus("STATE BLOCKED [" + tState + "][" + sCompany + "]", True)
                                        CallList_Error(tCallList, tState, "STATE BLOCKED")
                                        swOK = False
                                    End If
                            End Select
                        End If
                    End If
                    If swOK Then
                        '*********************************************************************************************************
                        '* 2012-11-12 RFK: PHONE
                        tPhone = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "PHONE", iQRYLoop).Trim
                        If tPhone.Length = 10 Then
                            '*****************************************************************************************************
                            '* 2012-11-12 RFK: ALREADY IN CALLED TODAY
                            If rkutils.DataGridViewContains(DataGridView_Phone, "PHONE", tPhone) >= 0 Then
                                'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", rkutils.WhoAmI() + " could NOT add to CallList " + tCallList + " [Already In CallList]", rkutils.STR_format("TODAY", "mm/dd/ccyy"), rkutils.WhoAmI())
                                MsgStatus("Could NOT add to CallList " + tCallList + " [Already In CallList Today] Phone:" + tPhone + "]", False)
                                swOK = False
                            End If
                            '*****************************************************************************************************
                            '* 2012-11-12 RFK: ALREADY IN CALLLIST FROM THIS LIST
                            If rkutils.Listbox_Contains(ListBox_Phones, tPhone, False) Then
                                'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", rkutils.WhoAmI() + " could NOT add to CallList " + tCallList + " [Already In List]", rkutils.STR_format("TODAY", "mm/dd/ccyy"), rkutils.WhoAmI())
                                MsgStatus("Could NOT add to CallList " + tCallList + " [Already In This List] Phone:" + tPhone + "]", False)
                                swOK = False
                            End If
                            If swOK Then
                                '*****************************************************************************************************
                                '* 2012-11-12 RFK: SET TIMEZONE
                                iAreaCodeRow = rkutils.DataGridViewContains(DataGridView_AreaCode, "AREACODE", rkutils.STR_LEFT(tPhone, 3))
                                If iAreaCodeRow >= 0 Then
                                    tTimeZone = rkutils.DataGridView_ValueByColumnName(DataGridView_AreaCode, "TIME_ZONE", iAreaCodeRow).Trim
                                    '*************************************************************************************************
                                    '* 2012-11-12 RFK: AREACODE BLOCKED 
                                    Select Case tRAMTTP
                                        Case "A"
                                            If rkutils.DataGridView_ValueByColumnName(DataGridView_AreaCode, "BLOCK_CALL_ACTIVE", iAreaCodeRow).Trim = "Y" Then
                                                MsgStatus("AREACODE BLOCKED [" + rkutils.STR_LEFT(tPhone, 3) + "]", True)
                                                CallList_Error(tCallList, tPhone, "AREACODE BLOCKED")
                                                swOK = False
                                            End If
                                        Case "C"
                                            If rkutils.DataGridView_ValueByColumnName(DataGridView_AreaCode, "BLOCK_CALL", iAreaCodeRow).Trim = "Y" Then
                                                MsgStatus("AREACODE BLOCKED [" + rkutils.STR_LEFT(tPhone, 3) + "]", True)
                                                CallList_Error(tCallList, tPhone, "AREACODE BLOCKED")
                                                swOK = False
                                            End If
                                    End Select
                                End If
                            End If
                            If swOK Then
                                '******************************************************
                                '* 2012-12-27 RFK: Invalid TimeZone
                                If tTimeZone.Trim.Length < 3 Then
                                    Select Case rkutils.STR_TRIM(sDialer, 1)
                                        Case "T"    'TeleServer
                                            MsgStatus("TeleServer/INVALID TIMEZONE", True)
                                            swOK = False
                                        Case Else   'Defaults to IAT
                                            '**************************************************
                                            '* 2017-09-20 RFK: IAT Dialer OK to SEND
                                            MsgStatus("IAT/INVALID TIMEZONE/SENDANYWAY", True)
                                    End Select
                                End If
                            End If
                        Else
                            'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", rkutils.WhoAmI() + " could NOT add to CallList " + tCallList + " [BAD PHONE](" + tPhone + ")", rkutils.STR_format("TODAY", "mm/dd/ccyy"), rkutils.WhoAmI())
                            MsgStatus("BAD PHONE [" + tPhone + "]", True)
                            CallList_Error(tCallList, tPhone, "BAD PHONE")
                            swOK = False
                        End If
                        If swOK Then
                            '**********************************************************
                            '* 2011-12-12 RFK: Balance
                            tBalance = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "MATCHBAL", iQRYLoop).Trim
                            If tBalance.Length > 0 Then
                                If swOK And CDec(tBalance) <= 0 Then
                                    '**************************************************
                                    '* 2021-01-20 RFK: Look at account Balance
                                    tBalance = Trim(Str(Val(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", iQRYLoop).Trim)))
                                End If
                                'MsgStatus("Balance:[" + tBalance + "]", True)
                                If swOK And CDec(tBalance) <= 0 Then
                                    '**************************************************
                                    'rkutils.TRACKS_update(msSQLConnectionString, msSQLuser, "", tLOCX, "", "T", rkutils.WhoAmI() + " could NOT add to CallList " + tCallList + " [MATCH BALANCE](" + tBalance + "]", rkutils.STR_format("TODAY", "mm/dd/ccyy"), rkutils.WhoAmI())
                                    MsgStatus("BALANCE [" + tBalance + "]", True)
                                    CallList_Error(tCallList, tBalance, "BALANCE")
                                    swOK = False
                                End If
                                If swOK And Label_BalanceGreaterThan.Text.Length > 0 Then
                                    If CDec(tBalance) < CDec(Label_BalanceGreaterThan.Text) Then
                                        '**********************************************
                                        MsgStatus("BALANCE [" + tBalance + "]", True)
                                        swOK = False
                                    End If
                                End If
                            End If
                        End If
                        If swOK Then
                            '************************************************
                            '* 2011-12-12 RFK: DateOfService
                            tDOS = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "DOS", iQRYLoop).Trim
                            'If swOK And IsDate(tDOS) Then swOK = False
                            '************************************************
                            '* 2012-04-12 RFK: Check The Bad Phone Flag
                            tPhoneFlag = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "PHONEFLAG", iQRYLoop).Trim
                            If tPhoneFlag = "B" Then
                                CallList_Error(tCallList, tPhoneFlag, "PHONEFLAG")
                                swOK = False
                            End If
                        End If
                    End If
                    '************************************************
                    'MsgStatus("tPhone:" + tPhone + "] Balance:" + tBalance + "] tPhoneFlag:" + tPhoneFlag + "] DOS" + tDOS + "][" + swOK.ToString + "]", False)
                    '************************************************
                    '* 2011-12-12 RFK: Still OK to GO
                    If swOK Then
                        '********************************************
                        '* 2012-04-0  RFK: for each client
                        '* 2012-09-21 RFK: by TOB/FACILITY
                        '* 2016-01-13 RFK: Facility Group
                        tClient = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RACL#", iQRYLoop).Trim
                        tTOB = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RATOB", iQRYLoop).Trim
                        tFacility = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", iQRYLoop).Trim
                        If tFacility.Length = 0 Then rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAFACL", iQRYLoop).Trim()
                        If tFacility.Length = 0 Then
                            MsgStatus("FACILITY WARNING!, EXITING", True)
                            Exit Sub
                        End If
                        tFacilityGroup = rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACGRP", iQRYLoop).Trim
                        tClientVMB = rkutils.WhatIsClientVMBValue(DataGridView2, msSQLConnectionString, msSQLuser, tClient, tTOB, tFacility, "VMB_MSG")
                        '********************************************
                        iRecNo += 1
                        SQLcommandstring += "INSERT INTO TeleServer.dbo.CallList"
                        SQLcommandstring += " (tunique"
                        SQLcommandstring += ", calllist"
                        SQLcommandstring += ", phone"
                        SQLcommandstring += ", timezone"
                        SQLcommandstring += ", vmb"
                        SQLcommandstring += ", client"
                        SQLcommandstring += ", id"
                        SQLcommandstring += ", account"
                        SQLcommandstring += ", contact"
                        SQLcommandstring += ", route"
                        SQLcommandstring += ", ref1"
                        SQLcommandstring += ", ref2"
                        SQLcommandstring += ", ref3"
                        SQLcommandstring += ", ref4"
                        SQLcommandstring += ", unique_name"
                        SQLcommandstring += ", modified_date, modified_by)"
                        SQLcommandstring += " values('" + Str(iRecNo).Trim + "'"
                        SQLcommandstring += ", '" + rkutils.STR_LEFT(tCallList, 20) + "'"
                        SQLcommandstring += ", '" + tPhone + "'"
                        SQLcommandstring += ", '" + tTimeZone + "'"
                        SQLcommandstring += ", '" + tClientVMB + "'"
                        SQLcommandstring += ", '" + tClient + "'"
                        SQLcommandstring += ", '" + tLOCX + "'"           'ID
                        Select Case STR_TRIM(sAccountNumberOutput, 1)
                            Case "M"    'MedRec
                                SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "MEDRECORD", iQRYLoop).Trim + "'"
                            Case "A"    'AccountNumber
                                SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim + "'"
                            Case Else
                                SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim + "'"
                        End Select
                        SQLcommandstring += ", '" + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FIRSTNAME", iQRYLoop).Trim) + " " + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LASTNAME", iQRYLoop).Trim) + "'"
                        SQLcommandstring += ", '" + tLOCX + "'"           'Route
                        SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", iQRYLoop).Trim + "'"          'Ref1
                        SQLcommandstring += ", '" + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", iQRYLoop).Trim) + "'"         'Ref2
                        SQLcommandstring += ", '" + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RATOB", iQRYLoop).Trim + "'"            'Ref3
                        SQLcommandstring += ", ''"    'Ref4
                        SQLcommandstring += ", " + Val(Label_Unique_Name.Text).ToString.Trim + ""
                        SQLcommandstring += ", '" + Date.Now + "'"
                        SQLcommandstring += ", '" + rkutils.WhoAmI() + "'"
                        SQLcommandstring += ")" + vbCr
                        '*******************************************
                        iAdded += 1
                        If iAdded Mod 100 = 0 Then
                            '* 2015-01-23 RFK: 
                            If swTEST = False Then rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
                            SQLcommandstring = ""
                        End If
                        '********************************************
                        '* 2012-12-27 RFK: Phone in this list
                        ListBox_Phones.Items.Add(tPhone)
                        '********************************************
                        '* 2013-05-17 RFK: 
                        ReportString = rkutils.STR_LEFT(tCallList, 20)
                        ReportString += tDelimiter + tLOCX
                        ReportString += tDelimiter + tPhone
                        ReportString += tDelimiter + tTimeZone
                        ReportString += tDelimiter + tClientVMB
                        ReportString += tDelimiter + tClient
                        ReportString += tDelimiter + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim
                        ReportString += tDelimiter + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FIRSTNAME", iQRYLoop).Trim) + " " + rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LASTNAME", iQRYLoop).Trim)
                        ReportString += tDelimiter + rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", iQRYLoop).Trim
                        File.AppendAllText(tReportName, ReportString + vbCrLf)
                        '********************************************
                        '* 2021-08-26 RFK: 
                        'MsgStatus("Creating Dialer Info " + sDialer, false)
                        Select Case rkutils.STR_TRIM(sDialer, 1)
                            Case "T"    'TCN
                                '********************************************
                                '* 2014-10-20 RFK: 
                                sDialerString = Chr(34) + "AnnuityHealth-1234" + Chr(34) 'TCN Account Number
                                sDialerString += "," + Chr(34) + "SelfPay" + Chr(34) 'Location
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tLOCX, 16) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LASTNAME", iQRYLoop).Trim), 40) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FIRSTNAME", iQRYLoop).Trim), 40) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAGZIP", iQRYLoop).Trim, 10) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAGSTATE", iQRYLoop).Trim, 10) + Chr(34)
                                '**********************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tPhone, 16) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "PHONEFLAG", iQRYLoop).Trim, 16) + Chr(34)
                                '**********************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 40) + Chr(34)                                                                                    'Client Full Name
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 40) + Chr(34)                                                                                    'Client Friendly Name
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClientVMB, 16) + Chr(34)                                                                                 'Office ID
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClientVMB, 16) + Chr(34)                                                                                 'Outpulse CID
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RATOB", iQRYLoop).Trim, 10) + Chr(34)
                                '**********************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", iQRYLoop).Trim, 10) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", iQRYLoop).Trim, 10) + Chr(34)
                                '**********************************************
                                '* 2016-01-13 RFK: Facility Group
                                Select Case sClientOutput
                                    Case "0"    'Client [CCC]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 16) + Chr(34)
                                    Case "1"    'Client-Facility [CCC-Facility]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + tFacility, 16) + Chr(34)
                                    Case "2"    'Client-Facility [CCC-FF]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + rkutils.STR_TRIM(tFacility, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "3"    'Client-FacilityGroup [CCC-GG]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + rkutils.STR_TRIM(tFacilityGroup, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "4"    'Client-Facility [C-FF]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 1) + "-" + rkutils.STR_TRIM(tFacility, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "5"    'Client-FacilityGroup [C-GG]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 1) + "-" + rkutils.STR_TRIM(tFacilityGroup, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case Else   'Client (CCC)
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 16) + Chr(34)
                                End Select
                                '**********************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ReportGroup", iQRYLoop).Trim, 10) + Chr(34)
                                '**********************************************
                                Select Case STR_TRIM(sAccountNumberOutput, 1)
                                    Case "M"    'MedRec
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "MEDRECORD", iQRYLoop).Trim, 20) + Chr(34)      'User Defined
                                    Case "A"    'AccountNumber
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim, 15) + Chr(34)  'User Defined
                                    Case Else
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim, 15) + Chr(34)  'User Defined [Account Number]
                                End Select
                                '**********************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", iQRYLoop).Trim, "0.00"), 10) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_format(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "MATCHBAL", iQRYLoop).Trim, "0.00"), 10) + Chr(34)
                            Case Else
                                '********************************************
                                '* 2014-10-20 RFK: 
                                sDialerString = Chr(34) + "1234" + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tLOCX, 16) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "LASTNAME", iQRYLoop).Trim), 40) + Chr(34)
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_NORMALIZE(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FIRSTNAME", iQRYLoop).Trim), 40) + Chr(34)
                                '******************************************************
                                '* 2016-01-13 RFK: Facility Group
                                Select Case sClientOutput
                                    Case "0"    'Client [CCC]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 16) + Chr(34)
                                    Case "1"    'Client-Facility [CCC-Facility]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + tFacility, 16) + Chr(34)
                                    Case "2"    'Client-Facility [CCC-FF]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + rkutils.STR_TRIM(tFacility, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "3"    'Client-FacilityGroup [CCC-GG]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 3) + "-" + rkutils.STR_TRIM(tFacilityGroup, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "4"    'Client-Facility [C-FF]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 1) + "-" + rkutils.STR_TRIM(tFacility, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case "5"    'Client-FacilityGroup [C-GG]
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.STR_TRIM(tClient, 1) + "-" + rkutils.STR_TRIM(tFacilityGroup, 2).PadLeft(2, "0"), 16) + Chr(34)
                                    Case Else   'Client (CCC)
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 16) + Chr(34)
                                End Select
                                '******************************************************
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClient, 40) + Chr(34)                                                                                    'Client Full Name
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tClientVMB, 16) + Chr(34)                                                                                 'Office ID
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                         'Desk ID
                                Select Case STR_TRIM(sAccountNumberOutput, 1)
                                    Case "M"    'MedRec
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "MEDRECORD", iQRYLoop).Trim, 10) + Chr(34)      'User Defined
                                    Case "A"    'AccountNumber
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim, 10) + Chr(34)  'User Defined
                                    Case Else
                                        sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "ACCOUNTNUMBER", iQRYLoop).Trim, 10) + Chr(34)  'User Defined [Account Number]
                                End Select
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 10) + Chr(34)                                                                                           'User Defined 2
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 10) + Chr(34)                                                                                           'User Defined 3
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(tPhone, 16) + Chr(34)                                                                                       'Phone
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "PHONEFLAG", iQRYLoop).Trim, 16) + Chr(34)           'Phone Flag 
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone2
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone2 Flag
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone3
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone3 Flag
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone4
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone4 Falg
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone5
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 16) + Chr(34)                                                                                           'Phone5 Flag
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 10) + Chr(34)                                                                                           'User Defined 4 [22]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM("", 10) + Chr(34)                                                                                           'User Defined 5 [23]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "BALANCE", iQRYLoop).Trim, 10) + Chr(34)             'User Defined 6 [24]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "FACILITY", iQRYLoop).Trim, 10) + Chr(34)            'User Defined 7 [25]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RATOB", iQRYLoop).Trim, 10) + Chr(34)               'User Defined 8 [26]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAGZIP", iQRYLoop).Trim, 10) + Chr(34)              'User Defined 9 [27]
                                sDialerString += "," + Chr(34) + rkutils.STR_TRIM(rkutils.DataGridView_ValueByColumnName(DataGridView_QRYoutput, "RAGSTATE", iQRYLoop).Trim, 10) + Chr(34)            'User Defined 10 [28]
                        End Select
                        ListBox_Dialer.Items.Add(sDialerString)
                    End If
                End If
                '**************************************************************************
                '* 2013-03-21 RFK: 
                Label_QueriesAccounts.Text = Trim(Str(Val(Label_QueriesAccounts.Text) - 1))
                iQRYLoop += 1
                System.Windows.Forms.Application.DoEvents()
            Loop
            If SQLcommandstring.Length > 0 Then
                '* 2015-01-23 RFK: 
                If swTEST = False Then rkutils.DB_COMMAND("MSSQL", msSQLConnectionString, msSQLuser, SQLcommandstring)
            End If
            '******************************************************************
            '* 2014-10-20 RFK: 
            '* 2015-12-30 RFK:
            MsgStatus("Creating Dialer Info " + sDialer, True)
            Select Case rkutils.STR_TRIM(sDialer, 1)
                Case "-"    '-TeleServer
                    MsgStatus("TeleServer", True)
                Case "T"    'TCN
                    '**********************************************************
                    '* 2021-08-26 RFK:
                    '* 2021-11-05 RFK:
                    sCallListDirectory = "\\production\Automation$\CallList\"
                    Select Case tTOB
                        Case "3"
                            sCallListDirectory += sCompany + "-SP\"
                        Case "5"
                            sCallListDirectory += sCompany + "-Ins\"
                        Case "222"
                            sCallListDirectory += sCompany + "-BD\"
                    End Select
                    '**********************************************************
                    MsgStatus("TCN:" + sCallListDirectory + tCallList, True)
                    '**********************************************************
                    If File.Exists(sCallListDirectory + tCallList + ".CSV") Then
                        File.Delete(sCallListDirectory + tCallList + ".CSV")
                    End If
                    '**********************************************************
                    Dim sHeaderRow As String = Chr(34) + "TCNnumber" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "Location" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "LOCX" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "LastName" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "FirstName" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "GZip" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "GState" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "Phone" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "PhoneFlag" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "Client" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "ClientFriendly" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "ClientVMB" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "OutPulseCID" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "TypeOfBusiness" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "Facility" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "FacilityFriendly" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "FacilityGroup" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "ReportGroup" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "MedRecOrAccountNumber" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "Balance" + Chr(34)
                    sHeaderRow += "," + Chr(34) + "BalanceMatched" + Chr(34)
                    File.AppendAllText(sCallListDirectory + tCallList + ".CSV", sHeaderRow + vbCrLf)
                    '**********************************************************
                    For i1 = 0 To ListBox_Dialer.Items.Count - 1
                        'MsgStatus(ListBox_Dialer.Items(i1).ToString, True)
                        If ListBox_Dialer.Items(i1).ToString.Length > 0 Then
                            File.AppendAllText(sCallListDirectory + tCallList + ".CSV", ListBox_Dialer.Items(i1).ToString.Trim + vbCrLf)
                        End If
                    Next
                    '**********************************************************
                Case Else   'Defaults to IAT
                    MsgStatus("DEFAULT/IAT", True)
                    If File.Exists("\\production\Reports\CallList\" + tCallList + ".TXT") Then
                        File.Delete("\\production\Reports\CallList\" + tCallList + ".TXT")
                    End If
                    For i1 = 0 To ListBox_Dialer.Items.Count - 1
                        If ListBox_Dialer.Items(i1).ToString.Length > 0 Then
                            '* 2015-01-23 RFK:
                            File.AppendAllText("\\production\Reports\CallList\" + tCallList + ".TXT", ListBox_Dialer.Items(i1).ToString.Trim + vbCrLf)
                        End If
                    Next
                    If File.Exists("\\production\Reports\CallList\" + tCallList + ".TXT") Then
                        If swTEST = False Then FileCopy("\\production\Reports\CallList\" + tCallList + ".TXT", "\\IAT_dialer\COLLECTIONS\" + tCallList + ".TXT")
                    End If
            End Select
            '******************************************************************
            '* 2013-02-21 RFK:
            Label_QueriesAccounts.Text = "0"
            '******************************************************************
            '* 2013-05-17 RFK: 
            If IS_File(tReportName) Then
                If Label_QRY_EMail.Text.Contains("@") Then
                    ReportString = "<html><body>"
                    ReportString += tCallList + " contains " + iAdded.ToString.Trim
                    ReportString += "<br><br>"
                    ReportString += Label_QRY_Name.Text
                    ReportString += "<br>"
                    ReportString += Label_QRY_Description.Text
                    ReportString += "<br>"
                    ReportString += tReportName.Trim
                    ReportString += "<br><br>"
                    ReportString += Label_QRY.Text
                    ReportString += "</body></html>"
                    rkutils.EMAILIT(msSQLConnectionString, msSQLuser, "DoNotReply@AnnuityHealth.com", "aoProcessor", Label_QRY_EMail.Text, Label_QRY_EMail.Text, "", "CallList " + tCallList + " created", "", ReportString, "")
                    '********************************************************
                    MsgStatus("Emailed to :" + Label_QRY_EMail.Text, True)
                End If
            End If
            '*******************************
            '* 2013-02-21 RFK: 
            If iAdded > 0 Then
                MsgStatus("Added " + iAdded.ToString.Trim + " to CallList:" + tCallList, True)
            Else
                MsgStatus("Added NONE to CallList:" + tCallList, True)
            End If
            '********************************************
            '* 2013-05-17 RFK:
            If sInsert = "Y" Then
                MsgStatus("Inserted CallList:" + tCallList, True)
            End If
            '************************************************************************************************************************
        Catch ex As Exception
            MsgStatus("CallList_ADD:" + ex.ToString, True)
        End Try
    End Sub

End Class
