Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.ComponentModel
Public Class Control
    Private setcom, com_barcode As String 'cai dat com
    Dim setlogmachine As String ' cai dat dia chi folder log ma may o cong doan tu sinh ra
    Dim setlogforwip As String ' cai dat dia chi luu file log cho he thong wip
    Dim setprocess As String ' cat dat ten cong doan
    Dim setsoftware As String ' cai dat ten soft can kich hoat
    Dim setfolder_report As String
    Dim stringtime As String ' bien ghep cac gia tri thoi gian
    Dim serial As String ' ma serial khi no dung tieu chuan
    Dim writeresult As String ' ghi ket qua vao file log cho he thong wip
    Dim setfilenamereport As String ' cai dat ten lien quan den file name của file report may chuc nang
    Dim filenamelog As String
    Dim namemodel, namemodel1, namemodel2 As String
    Dim barcode_rule, barcode_rule1, barcode_rule2 As String
    Dim ngay, thang, gio, phut, giay As String
    Dim start As Boolean
    Dim npass As Integer = 0
    Dim nfail As Integer = 0
    Dim ntotal As Integer = 0
    Dim labelcheck As Boolean
    'Dim j As Integer
    Dim strnamesoft, linereport As String
    Dim new_filename_report As String
    Dim countread As Integer
    Dim countlinecurrent, countlinebegin As Integer
    Dim length_label As Integer
    Dim retry As Integer
    Dim filename_passrate As String = ""
    Private Property filename_sys_setting As String
    Delegate Sub SetTextCallback(ByVal [text] As String) 'Added to prevent threading errors during receiveing of data
    '======lien quan den process task manager
    Public Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Dim h As String
    Dim i, x As Boolean
    Private Sub Control_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If SerialPort2.IsOpen = True Then
            SerialPort2.Writeline("F") ' tat role truoc khi dong cong com
            Serialport2.Close() ' dong cong com
        Else
            SerialPort2.Open()
            SerialPort2.Writeline("F") ' tat role truoc khi dong cong com
            Serialport2.Close() ' dong cong com
        End If
        End
    End Sub
    Private Sub setup_display()
        Labelresult.BackColor = Color.Blue
        Labelresult.ForeColor = Color.White
        Labelresult.Text = "Wait"
        textserial.Enabled = True
        textserial.Text = ""
        textserial.Focus()
    End Sub

    Private Sub Control_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadsetting() ' load thong tin cai dat
        passpercent()
        setup_display()
        SerialPort2.Writeline("A")
        For j As Integer = 1 To 100
            Datagrid.Rows.Add()
        Next
        Datagrid.Enabled = False

    End Sub
    '-------------------------------------
    Private Sub CommPortSetup()
        ' setup com control
        With SerialPort2
            .PortName = setcom
            .BaudRate = 9600
            .DataBits = 8
            .Parity = Parity.None
            .StopBits = StopBits.One
            .Handshake = Handshake.None
            .ReceivedBytesThreshold = 1
        End With
        Try
            SerialPort2.Open()
        Catch ex As Exception
            MsgBox(setcom & "not connect. Byby and See you later !")
            End
        End Try

        'setup com barcode
        With SerialPort1
            .PortName = com_barcode
            .BaudRate = 115200
            .DataBits = 8
            .Parity = Parity.Even
            .StopBits = StopBits.One
            .Handshake = Handshake.None
            .ReceivedBytesThreshold = 1
        End With
        Try
            SerialPort1.Open()
        Catch ex As Exception
            MsgBox("com_barcode: " & com_barcode & "not connect. Byby and See you later !")
            End
        End Try
        SerialPort1.DiscardInBuffer()
        SerialPort1.DiscardOutBuffer()
        SerialPort2.DiscardInBuffer()
        SerialPort2.DiscardOutBuffer()
    End Sub



    Private Sub passpercent()

        ntotal = npass + nfail
        TextTotal.Text = ntotal
        TextPass.Text = npass
        TextNg.Text = nfail
        If ntotal <> 0 Then
            Passrate.Text = "Pass Rate: " & Format((npass / ntotal) * 100, "0.00") & "%"
        Else
            Passrate.Text = "Pass Rate: " & "0%"
        End If
        If System.IO.File.Exists("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt") = False Then ' kiem tra xem file text cua lot no nay da duoc tao hay chua
            npass = 0
            nfail = 0
            ntotal = 0
            filename_passrate = "C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt"
            Dim text_passrate As New System.IO.StreamWriter(filename_passrate)
            text_passrate.WriteLine("# Total")
            text_passrate.WriteLine(ntotal)
            text_passrate.WriteLine("# Pass")
            text_passrate.WriteLine(npass)
            text_passrate.WriteLine("# NG")
            text_passrate.WriteLine(nfail)
            text_passrate.WriteLine("# PASSRATE")
            text_passrate.WriteLine(Passrate.Text)
            text_passrate.Close()
        End If
    End Sub
    '===========================================
    '--------------------------------------------------------------------------------------------------------
    Private Sub loadsetting()
        dinhdangthoigian()

        If System.IO.File.Exists("C:\setting support fct\setting.txt") = True Then
            com_barcode = Mid(ReadTextFile("C:\setting support fct\setting.txt", 2), 1, 4) ' cai dat cong com
            setcom = Mid(ReadTextFile("C:\setting support fct\setting.txt", 2), 6, 4) ' cai dat cong com
            setlogforwip = ReadTextFile("C:\setting support fct\setting.txt", 4) ' cai dat dia folder lưu file log cho hệ thống wip
            setlogmachine = ReadTextFile("C:\setting support fct\setting.txt", 6) ' cai dat dia chi folder log mà máy chuc nang se sinh ra.
            setprocess = ReadTextFile("C:\setting support fct\setting.txt", 8) ' cai dat ten cong doan
            barcode_rule1 = Mid(ReadTextFile("C:\setting support fct\setting.txt", 10), 1, 3) ' Cai dat barcode rule model 1

            barcode_rule2 = Mid(ReadTextFile("C:\setting support fct\setting.txt", 12), 1, 3) ' Cai dat barcode rule model 2

            namemodel1 = Mid(ReadTextFile("C:\setting support fct\setting.txt", 10), 5, 9) ' cai dat name model code 1
            namemodel2 = Mid(ReadTextFile("C:\setting support fct\setting.txt", 12), 5, 9) ' cai dat name model code 2

            setsoftware = ReadTextFile("C:\setting support fct\setting.txt", 14) ' cai dat ten phan mem can kich hoat
            setfilenamereport = ReadTextFile("C:\setting support fct\setting.txt", 16) ' cai dat ten file log: TMP1.LOG
            Timer1.Interval = ReadTextFile("C:\setting support fct\setting.txt", 18) ' cai dat thoi gian cho timer 1
            Timer2.Interval = ReadTextFile("C:\setting support fct\setting.txt", 20) ' cai dat thoi gian cho timer 2
            Timer3.Interval = ReadTextFile("C:\setting support fct\setting.txt", 22) ' cai dat thoi gian cho timer 3
            length_label = ReadTextFile("C:\setting support fct\setting.txt", 24) ' doc cai dat do dai cho phep cua barcode
            Txt_Retry.Text = ReadTextFile("C:\setting support fct\setting.txt", 26) ' CAI DAT SO LAN RETRAY

            Me.Text = setprocess
            CommPortSetup() ' set up comport 
            Labelnameprocess.Text = setprocess
            Timer3.Enabled = True
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Log_Report") = False Then ' chac chan da co folder report trong folder luu log of machine
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Log_Report") ' tao folder luu file report
            End If
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Log_Report" & Now.Year) = False Then
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Log_Report\" & Now.Year) ' tao folder nam moi: C:\Report\2014
            End If
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang) = False Then
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang) ' tao folder nam moi: C:\setting support fct\Report\2014\01
            End If
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\OK") = False Then
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\OK") ' tao folder nam moi: C:\Report\2014\01\OK
            End If
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\NG") = False Then
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\NG") ' tao folder nam moi: C:\setting support fct\Report\2014\01\NG
            End If
            If My.Computer.FileSystem.DirectoryExists("C:\setting support fct\Passrate") = False Then
                My.Computer.FileSystem.CreateDirectory("C:\setting support fct\Passrate") ' tao folder nam moi: C:\setting support fct\Report\2014\01\NG
            End If
            '---------------------------------------------------------
            If System.IO.File.Exists("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt") = True Then
                ntotal = ReadTextFile("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt", 2)
                npass = ReadTextFile("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt", 4)
                nfail = ReadTextFile("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt", 6)
            Else
                npass = 0
                nfail = 0
                ntotal = 0
                filename_passrate = "C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt"
                Dim text_passrate As New System.IO.StreamWriter(filename_passrate)
                text_passrate.WriteLine("# Total")
                text_passrate.WriteLine(ntotal)
                text_passrate.WriteLine("# Pass")
                text_passrate.WriteLine(npass)
                text_passrate.WriteLine("# NG")
                text_passrate.WriteLine(nfail)
                text_passrate.WriteLine("# Passrate")
                text_passrate.WriteLine(Passrate.Text)
                text_passrate.Close()
            End If
        Else
            MsgBox(" File setting.txt in C:\settingforwip not found ")
            End
        End If
    End Sub

    Private Sub dinhdangthoigian()
        If Now.Month < 10 Then
            thang = "0" & Now.Month
        Else
            thang = Now.Month
        End If
        If Now.Day < 10 Then
            ngay = "0" & Now.Day
        Else
            ngay = Now.Day
        End If
        If Now.Hour < 10 Then
            gio = "0" & Now.Hour
        Else
            gio = Now.Hour
        End If
        If Now.Minute < 10 Then
            phut = "0" & Now.Minute
        Else
            phut = Now.Minute
        End If
        If Now.Second < 10 Then
            giay = "0" & Now.Second
        Else
            giay = Now.Second
        End If
    End Sub


    Private Sub textserial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textserial.KeyPress

        If Asc(e.KeyChar) = 13 Then
            SerialPort1.WriteLine(Chr(&H16) & "U" & Chr(&HD))
            TextReceive.Text = ""
            '=====================================
            textserial.Text = StrConv(textserial.Text, VbStrConv.Uppercase) ' chuyen thanh chu in hoa

            If textserial.TextLength > 10 Then
                textserial.Text = Microsoft.VisualBasic.Left(textserial.Text, 10)
            End If

            If textserial.TextLength <> 0 And textserial.TextLength = length_label Then
                If (InStr(textserial.Text, barcode_rule1, 0) <> 0 And (namemodel1 = namemodel1)) Or (InStr(textserial.Text, barcode_rule2, 0) <> 0 And (namemodel2 = namemodel2)) Then
                    ' kiem tra xem co ten model 1 trong serial hay khong Or InStrRev(textserial.Text, namemodel2) = True Then ' kiem tra xem co ten model 1 trong serial hay khong
                    labelcheck = True
                    serial = textserial.Text
                    barcode_rule = Mid(textserial.Text, 1, 3) ' lay ki tu barcode_rule
                Else
                    labelcheck = False
                    MsgBox("Barcode Rule was Wrong")
                End If

            Else
                labelcheck = False
                MsgBox(" Length Barcode wrong !")
                'setup_display()
            End If
            If System.IO.File.Exists(setlogmachine & "\" & Now.Year & thang & ngay & "." & setfilenamereport) = True Then
                ' dong luu thong tin log
                countlinebegin = CounterlineTextFile(setlogmachine & "\" & Now.Year & thang & ngay & "." & setfilenamereport)
            End If
            If labelcheck = True Then
                For k As Integer = 1 To Datagrid.RowCount - 1
                    Datagrid.Rows.Item(Datagrid.RowCount - k).Cells(0).Value = Datagrid.Rows.Item(Datagrid.RowCount - 1 - k).Cells(0).Value
                    Datagrid.Rows.Item(Datagrid.RowCount - k).Cells(1).Value = Datagrid.Rows.Item(Datagrid.RowCount - 1 - k).Cells(1).Value
                    Datagrid.Rows.Item(Datagrid.RowCount - k).Cells(2).Value = Datagrid.Rows.Item(Datagrid.RowCount - 1 - k).Cells(2).Value
                    Datagrid.Rows.Item(Datagrid.RowCount - k).Cells(3).Value = Datagrid.Rows.Item(Datagrid.RowCount - 1 - k).Cells(3).Value
                Next
                Datagrid.Rows.Item(0).Cells(0).Value = "" ' xoa thong tin cua dong dau trong datagrid
                Datagrid.Rows.Item(0).Cells(1).Value = ""
                Datagrid.Rows.Item(0).Cells(2).Value = ""
                Datagrid.Rows.Item(0).Cells(3).Value = ""
                Labelresult.BackColor = Color.Yellow
                Labelresult.ForeColor = Color.Red
                Labelresult.Text = "Wait"
                textserial.Enabled = False

                AppActivate(Findapplication(setsoftware))
                Thread.Sleep(100)
                SendKeys.SendWait(textserial.Text)
                Thread.Sleep(100)
                Labelresult.BackColor = Color.Yellow
                Labelresult.ForeColor = Color.Red
                Labelresult.Text = "Busy"
                SendKeys.SendWait("{ENTER}")
                start = True
                Timer3.Enabled = True
                Timer2.Enabled = True
            End If
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lbtime.Text = Now
        passpercent()
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If start = True Then ' khi may da chay roi 
            'dinh dang thoi gian ve form 2014 =14; thang <10 = 0x, thang >9 = xx ; ngay , gio , phut , giay : tuong tu
            dinhdangthoigian()
            ' stringtime la bien thoi gian de ghep voi so serial tao thanh file name cho log
            stringtime = Mid(Now.Year, 3, 2) & thang & ngay & gio & phut & giay
            ' neu file log ton tai
            If System.IO.File.Exists(setlogmachine & "\" & Now.Year & thang & ngay & "." & setfilenamereport) Then
                ' dong luu thong tin log
                countlinecurrent = CounterlineTextFile(setlogmachine & "\" & Now.Year & thang & ngay & "." & setfilenamereport)
                If countlinecurrent > countlinebegin Then
                    countlinebegin = countlinecurrent
                    linereport = ReadTextFile(setlogmachine & "\" & Now.Year & thang & ngay & "." & setfilenamereport, countlinecurrent)
                    '''''''''.............Kiem tra model check co dung hay khong?...............................
                    If InStr(linereport, namemodel1, 0) <> 0 And barcode_rule <> barcode_rule1 Then
                        Timer2.Enabled = False
                        Timer3.Enabled = True
                        AppActivate(Findapplication("FCT_Power_Brother"))
                        Me.Hide()
                        Form1.Show()
                        Exit Sub
                    End If

                    If InStr(linereport, namemodel2, 0) <> 0 And barcode_rule <> barcode_rule2 Then
                        Timer2.Enabled = False
                        Timer3.Enabled = True
                        AppActivate(Findapplication("FCT"))
                        Me.Hide()
                        Form1.Show()
                        Exit Sub
                    End If
                    '------------------------------------------------------------------------------------------
                    If InStr(linereport, "NG", 0) <> 0 Then
                        If retry < Txt_Retry.Text Then
                            retry = retry + 1
                            Thread.Sleep(1000)
                            SendKeys.SendWait(textserial.Text)
                            SendKeys.SendWait("{ENTER}")
                            Exit Sub
                        End If
                        Datagrid.Rows.Item(0).Cells(0).Value = ngay & "/" & thang & "/" & Now.Year
                        Datagrid.Rows.Item(0).Cells(1).Value = gio & ":" & phut & ":" & giay
                        Datagrid.Rows.Item(0).Cells(2).Value = serial
                        Labelresult.BackColor = Color.Red
                        Labelresult.ForeColor = Color.White
                        Labelresult.Text = "NG"
                        nfail = nfail + 1
                        Thread.Sleep(1000)
                        SendKeys.SendWait("{TAB}") '1
                        SendKeys.SendWait("{TAB}")
                        SendKeys.SendWait("{TAB}")
                        SendKeys.SendWait("{TAB}") '4
                        SendKeys.SendWait("{TAB}")
                        SendKeys.SendWait("{TAB}")
                        SendKeys.SendWait("{TAB}")
                        SendKeys.SendWait("{TAB}") '8
                        SendKeys.SendWait("{ENTER}")
                        Datagrid.Rows.Item(0).Cells(3).Value = "NG"
                        SerialPort2.WriteLine("N")
                        filenamelog = "C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\NG\" & stringtime & "_ " & serial & "_Report.csv"
                        Dim textreportforpcb As New System.IO.StreamWriter(filenamelog, True)
                        textreportforpcb.WriteLine(linereport) ' ghi thong tin vao file log
                        textreportforpcb.Close()
                        retry = 0
                        countread = 0
                        start = False

                        AppActivate(setprocess)
                        Timer3.Enabled = True
                        Timer2.Enabled = False
                        textserial.Enabled = True
                        textserial.Focus()
                        TextReceive.Text = ""
                        textserial.Text = ""
                        SerialPort1.DiscardInBuffer()
                        SerialPort1.DiscardOutBuffer()
                        SerialPort2.DiscardInBuffer()
                        SerialPort2.DiscardOutBuffer()
                    Else
                        Datagrid.Rows.Item(0).Cells(0).Value = ngay & "/" & thang & "/" & Now.Year
                        Datagrid.Rows.Item(0).Cells(1).Value = gio & ":" & phut & ":" & giay
                        Datagrid.Rows.Item(0).Cells(2).Value = serial
                        SerialPort2.WriteLine("O")
                        Labelresult.BackColor = Color.Green
                        Labelresult.ForeColor = Color.White
                        Labelresult.Text = "OK"
                        npass = npass + 1
                        Datagrid.Rows.Item(0).Cells(3).Value = "OK"
                        filenamelog = "C:\setting support fct\Log_Report\" & Now.Year & "\" & thang & "\OK\" & stringtime & "_ " & serial & "_Report.csv"
                        Dim textreportforpcb As New System.IO.StreamWriter(filenamelog, True)
                        textreportforpcb.WriteLine(linereport) ' ghi thong tin vao file log
                        textreportforpcb.Close()
                        retry = 0
                        countread = 0
                        start = False

                        AppActivate(setprocess)
                        Timer3.Enabled = True
                        Timer2.Enabled = False
                        textserial.Enabled = True
                        textserial.Focus()
                        TextReceive.Text = ""
                        textserial.Text = ""
                        SerialPort1.DiscardInBuffer()
                        SerialPort1.DiscardOutBuffer()
                        SerialPort2.DiscardInBuffer()
                        SerialPort2.DiscardOutBuffer()
                    End If
                End If
                If System.IO.File.Exists("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt") = True Then ' kiem tra xem file text cua lot no nay da duoc tao hay chua
                    System.IO.File.Delete("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt")
                End If
                filename_passrate = "C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt"
                Dim text_passrate As New System.IO.StreamWriter(filename_passrate)
                text_passrate.WriteLine("# Total")
                text_passrate.WriteLine(ntotal)
                text_passrate.WriteLine("# Pass")
                text_passrate.WriteLine(npass)
                text_passrate.WriteLine("# NG")
                text_passrate.WriteLine(nfail)
                text_passrate.WriteLine("# Passrate")
                text_passrate.WriteLine(Passrate.Text)
                text_passrate.Close()
                Timer3.Enabled = True
                TextReceive.Text = ""
            End If
        End If
    End Sub
    Private Sub Sent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sent.Click
        If SerialPort2.IsOpen = False Then ' neu cong com dang dong thi mo lai
            SerialPort2.Open()
            SerialPort2.Writeline(TextSent.Text)
        Else
            SerialPort2.Writeline(TextSent.Text)
        End If
    End Sub

    'Private Sub Receive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Receive.Click
    '    If SerialPort2.IsOpen = False Then ' neu cong com dang dong thi mo lai
    '        CommPortSetup()
    '        SerialPort2.DiscardInBuffer()
    '        TextReceive.Text = SerialPort2.ReadExisting
    '    Else
    '        TextReceive.Text = SerialPort2.ReadExisting
    '    End If
    'End Sub
    Private Sub Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset.Click
        setup_display()
        Timer2.Enabled = False
        Timer3.Enabled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Do you want to exit?", "Support FCT ", _
MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
= DialogResult.Yes Then
            If Serialport2.IsOpen = True Then
                SerialPort2.Writeline("F")
                Serialport2.Close()
            Else
                Serialport2.Open()
                SerialPort2.Writeline("F")
                Serialport2.Close()
            End If
            End
        End If
    End Sub

    Private Sub Clear_report_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear_report.Click
        npass = 0
        nfail = 0
        ntotal = 0
        If System.IO.File.Exists("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt") = True Then ' kiem tra xem file text cua lot no nay da duoc tao hay chua
            System.IO.File.Delete("C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt")
        End If
        filename_passrate = "C:\setting support fct\Passrate\" & Now.Year & thang & ngay & "_Passrate.txt"
        Dim text_passrate As New System.IO.StreamWriter(filename_passrate)
        text_passrate.WriteLine("# Total")
        text_passrate.WriteLine(ntotal)
        text_passrate.WriteLine("# Pass")
        text_passrate.WriteLine(npass)
        text_passrate.WriteLine("# NG")
        text_passrate.WriteLine(nfail)
        text_passrate.WriteLine("# Passrate")
        text_passrate.WriteLine(Passrate.Text)
        text_passrate.Close()
        passpercent()
        textserial.Focus()
    End Sub




    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If textserial.Text = "" Then
            textserial.Text = SerialPort1.ReadExisting()
            If textserial.Text <> "" Then
                textserial.Focus()
                SendKeys.SendWait("{ENTER}")
                SerialPort1.WriteLine(Chr(&H16) & "U" & Chr(&HD))
                SerialPort1.DiscardInBuffer()
                SerialPort1.DiscardOutBuffer()
                Timer3.Enabled = False
            Else
                Exit Sub
            End If
        End If
        'Timer3.Enabled = False
    End Sub


    Private Sub SerialPort2_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort2.DataReceived

        Receivedtext_rs232(SerialPort2.ReadExisting())
        If TextReceive.Text = "S" Then
            SerialPort1.WriteLine(Chr(&H16) & "T" & Chr(&HD))
            SerialPort2.WriteLine("S")
        End If

    End Sub
    Private Sub Receivedtext_rs232(ByVal [text] As String) 'input from ReadExisting
        If TextReceive.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf Receivedtext_rs232)
            Me.Invoke(x, New Object() {(text)})
        Else
            TextReceive.Text &= [text] 'append text
        End If
    End Sub

End Class
