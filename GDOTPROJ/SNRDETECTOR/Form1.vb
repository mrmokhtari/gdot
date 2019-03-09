Imports System.Net.Sockets
'Imports System.Text
Imports System.Text
Imports System.IO.StreamReader
Imports System.IO.StreamWriter
Imports System.IO.File
Public Class Form1
    Dim WD As Boolean
    Dim k As System.IAsyncResult
    ' Private oTCPStream As Net.Sockets.NetworkStream
    ' Private oTCP As New Net.Sockets.TcpClient()
    Private bytWriting As Byte()
    Private bytReading As Byte()
    Dim rcb As System.AsyncCallback
    Dim st As Object
    Dim SEEKS As Integer
    Dim SHR As String
    Private Function isready(ByVal txt As String) As Boolean
        Dim l As Integer
        Dim s As String
        l = Len(txt)
        s = Mid(txt, l - 1)
        If s = ">" Or s = "#" Then
            isready = True
        Else
            isready = False
        End If

    End Function
    Private Function RSTR(ByVal OTCPStream As System.Net.Sockets.NetworkStream, ByVal OTCP As System.Net.Sockets.TcpClient) As String
        RSTR = ReadData(OTCPStream, OTCP) & vbCrLf

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'ZTEMAIN()
        'HUAWEISNR("172.27.129.135")
    End Sub

    Private Function ReadData(ByVal OTCPStream As System.Net.Sockets.NetworkStream, ByVal OTCP As System.Net.Sockets.TcpClient) As String
        On Error Resume Next
        Dim sData As String
        ReDim bytReading(oTCP.ReceiveBufferSize)
        WD = True
        OTCPStream.Read(bytReading, 0, oTCP.ReceiveBufferSize)
        sData = Trim(System.Text.Encoding.Default.GetString(bytReading))
        WD = False
        ReadData = sData
    End Function

    Private Sub WriteData(ByVal sData As String, ByVal OTCPStream As System.Net.Sockets.NetworkStream)
        On Error Resume Next
        bytWriting = System.Text.Encoding.Default.GetBytes(sData)
        oTCPStream.Write(bytWriting, 0, bytWriting.Length)
    End Sub
    Public Sub HUAWEISNR(ByVal IP As String, ByVal HNAM As String)
        'On Error Resume Next
        ' If IP = "172.27.7.51" Then Exit Sub
        Dim SLT As Integer
        Dim DES, LST, STAT, PVC, FDT, TM As String
        Dim VLAN, VPI, VCI As Integer
        Dim ens, strtmp, TSTR, TL, MRK As String
        Dim USNR, DSNR, TXST As Integer
        Dim URATE, DRATE, UATT, DATT As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        Dim SL, PRT, STL As Integer
        Dim DT As Date
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim a() As String
        Dim PTM As Boolean = False

        cnn.Open()
        OTCP = New System.Net.Sockets.TcpClient
        '        oTCPStream = New System.Net.Sockets.NetworkStream
        ens = IP
        ens = ens.Replace(".", "_")
        ens = "HOA" & ens & Now.Year.ToString & "_" & Now.Month.ToString & "_" & Now.Day.ToString
        'IP = "192.168.2.1"
        'Dim sw As New IO.StreamWriter("G:\TXTS\" & ens & ".TXT")

        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        'oTCP.EndConnect()
        On Error Resume Next
        '   MRK = MRKAZ(IP)
        OTCP.Connect(IP, "23")
        'k = oTCP.BeginConnect(IP, 23, rcb, st)
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        If Trim(TextBox1.Text) = "" Then
            DSACC(IP, False)
            Exit Sub
        Else
            DSACC(IP, True)
        End If
        System.Threading.Thread.Sleep(1000)
        WriteData("root" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        Dim LL As Integer
        LL = 0
        LL = InStr(strtmp, "invalid")
        If LL > 0 Then DSACC(IP, False) : Exit Sub


        'ListBox1.Items.Add(strtmp)
        'TextBox1.Text = strtmp
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT * FROM GDOTSLOTS WHERE IP ='" & IP & "'", cnn).ExecuteReader
        Do While RDATA.Read
            SLT = CInt(RDATA!SLOT)
            'NOW START TO SCAN PORTS
            For I = 0 To 31
                VLAN = 0 : DES = "" : STAT = "" : PVC = "" : FDT = "" : TM = ""
                System.Threading.Thread.Sleep(100)
                TextBox1.Text = ""
                WriteData("display interface shdsl 0/" & SLT & "/" & I & vbCrLf, OTCPStream) 'Password
                RichTextBox1.Clear()
                System.Threading.Thread.Sleep(250)
                WriteData(" ", OTCPStream) 'Password
                System.Threading.Thread.Sleep(250)
                strtmp = RSTR(OTCPStream, OTCP)
                RichTextBox1.AppendText(strtmp)
                strtmp = ""
                While InStr(strtmp, HNAM & "#") = 0
                    WriteData(" ", OTCPStream) 'Password
                    System.Threading.Thread.Sleep(200)
                    strtmp = RSTR(OTCPStream, OTCP)
                    RichTextBox1.AppendText(strtmp)

                End While
                RichTextBox1.AppendText(strtmp)
                For Each LIN In RichTextBox1.Lines
                    TSTR = LIN
                    LL = 0
                    LL = InStr(TSTR, "Description")
                    If LL > 0 Then
                        TSTR = Mid(TSTR, LL + 1)
                        LL = 0
                        LL = InStr(TSTR, ":")
                        If LL > 0 Then
                            TSTR = Trim(Mid(TSTR, LL + 1))
                            DES = TSTR
                        End If
                    End If
                    LL = 0
                    ' Shdsl  0/ 0/ 0 is down
                    LL = InStr(TSTR, "Shdsl")
                    If LL > 0 Then
                        TSTR = Mid(TSTR, LL + 12)
                        PRT = Val(TSTR)

                        LL = 0
                        LL = InStr(TSTR, "is")
                        If LL > 0 Then
                            TSTR = Trim(Mid(TSTR, LL + 3))
                            STAT = TSTR
                        End If
                    End If
                    LL = 0
                    LL = InStr(TSTR, "Last up time")
                    If LL > 0 Then
                        TSTR = Mid(TSTR, LL + 1)
                        LL = 0
                        LL = InStr(TSTR, ":")
                        If LL > 0 Then
                            TSTR = Trim(Mid(TSTR, LL + 1))
                            LST = TSTR
                            If Len(LST) > 5 Then
                                DT = CDate(Microsoft.VisualBasic.Left(LST, 10))
                                TM = Microsoft.VisualBasic.Left(Trim(Mid(TSTR, 13)), 8)
                                FDT = FDATE(DT)
                            Else
                                FDT = "NEVER"
                                TM = "NEVER"
                            End If

                        End If
                    End If
                    LL = 0
                    LL = InStr(TSTR, "PTM")
                    If LL > 0 Then
                        PTM = True
                        PVC = "PTM"
                    End If
                    LL = 0
                    LL = InStr(TSTR, "PVC  ")
                    If LL > 0 Then
                        TSTR = Trim(Mid(TSTR, LL + 4))
                        PVC = TSTR
                    End If
                    LL = 0
                    LL = InStr(TSTR, "VLAN")
                    If LL > 0 Then
                        If PTM Or PVC = "0/33" Or PVC = "8/35" Or PVC = "6/35" Or PVC = "0/34" Then
                            TSTR = Mid(TSTR, LL + 1)
                            LL = 0
                            LL = InStr(TSTR, ":")
                            If LL > 0 Then
                                TSTR = Trim(Mid(TSTR, LL + 1))
                                VLAN = Val(TSTR)
                                Exit For
                            End If
                        End If
                    End If

                Next
                If VLAN > 0 Then
                    ens = "INSERT INTO GDOT(STAT, DES, PVC, VLAN, LST, IP, SLT, PRT,FDT,FTM) VALUES (N'" & STAT & "', N'" & DES & "', '" & PVC & "', " & VLAN & ", N'" & LST & "', N'" & IP & "', " & SLT & ", " & PRT & ",'" & FDT & "','" & TM & "')"
                    SQLEXEC(ens)
                End If
                PTM = False
                VLAN = 0

NOFILE:
            Next

        Loop
        RDATA.Close()


        OTCPStream.Close()
        OTCP.Close()


    End Sub
    Private Function GTSTL(ByVal IP As String, ByVal SLT As Integer) As Integer
        Dim STL As Integer
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        '  Dim INSTREE As System.Data.SqlClient.SqlDataReader
        STL = 1950
        cnn.Open()
        Dim RDA As System.Data.SqlClient.SqlDataReader
        RDA = New System.Data.SqlClient.SqlCommand("SELECT STL FROM DSLSLOT WHERE (IP='" & IP & "' AND SLOT=" & SLT & ") ", cnn).ExecuteReader
        Do While RDA.Read
            STL = CInt(RDA!STL)
        Loop
        RDA.Close()
        cnn.Close()
        GTSTL = STL
    End Function
    Public Sub FIBERSNR(ByVal IP As String, ByVal HST As String)
        Dim ens, strtmp, TL, MRK As String
        'Dim IP As String = "172.27.178.140"
        Dim USNR, DSNR, TXST As Integer
        Dim URATE, DRATE, UATT, DATT As Integer
        Dim STL, SL, PRT, L, LL, SLT As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        Dim a() As String
        Dim INSTREE As System.Data.SqlClient.SqlCommand
        ens = IP
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        On Error Resume Next
        OTCP.SendTimeout = 3000
1:

        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)

        System.Threading.Thread.Sleep(1000)
        '        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        '        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(strtmp)
        TextBox1.Text = strtmp
        WriteData("show running-config" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(500)
        strtmp = RSTR(OTCPStream, OTCP)
        RichTextBox1.AppendText(strtmp)
        strtmp = ""
        While InStr(strtmp, HST & "#") = 0
            WriteData(" ", OTCPStream) 'Password
            System.Threading.Thread.Sleep(200)
            strtmp = RSTR(OTCPStream, OTCP)
            RichTextBox1.AppendText(strtmp)

        End While
        RichTextBox1.AppendText(strtmp)
        'استخراج اطلاعات 
        'MsgBox("finished")
        Dim VLANS As New Dictionary(Of Integer, String)
        Dim TSTR As String
        Dim VLAN, VMIN, VMAX As Integer
        Dim VLANDES As String
        For Each LIN In RichTextBox1.Lines
            TSTR = LIN
            If InStr(TSTR, "add uplink vlan") > 0 Then
                TSTR = Trim(Mid(TSTR, 19))
                LL = InStr(TSTR, " ")
                TSTR = Trim(Mid(TSTR, LL))
                LL = InStr(TSTR, " start")
                VLANDES = Trim(Mid(TSTR, 1, LL))
                If Microsoft.VisualBasic.Right(VLANDES, 1) = "2" Then
                    VLANDES = Mid(VLANDES, 1, Len(VLANDES) - 1)
                End If
                TSTR = Trim(Mid(TSTR, LL + 6))
                LL = InStr(TSTR, " end")

                VMIN = Val(TSTR)
                TSTR = Trim(Mid(TSTR, LL + 4))
                VMAX = Val(TSTR)
                For i = VMIN To VMAX
                    VLANS(i) = VLANDES
                Next
            End If
        Next
        Dim VAZ As String
        Dim SLTT, PRTT As Integer
        For Each LIN In RichTextBox1.Lines
            TSTR = LIN
            If InStr(TSTR, "interface ") > 0 Then
                LL = InStr(TSTR, "interface ")
                TSTR = Trim(Mid(TSTR, LL + 10, 5))
                LL = InStr(TSTR, "/")
                SLTT = Val(TSTR)
                PRTT = Val(Mid(TSTR, LL + 1))
            End If
            If InStr(TSTR, "add unicast tag cvlan") > 0 Then
                TSTR = Trim(Mid(TSTR, 19))
                LL = InStr(TSTR, " ")
                TSTR = Trim(Mid(TSTR, LL))
                LL = InStr(TSTR, " vid ")
                VMIN = Val(Trim(Mid(TSTR, LL + 4)))
                If VLANS.ContainsKey(VMIN) Then
                    VAZ = "IN USE"
                    '***********************************************************
                    '  ens = "INSERT INTO GDOT(STAT, DES, PVC, VLAN, LST, IP, SLT, PRT,FDT,FTM) VALUES (N'" & VAZ & "', N'" & VLANS(VMIN) & "', '" & PVC & "', " & VLAN & ", N'" & LST & "', N'" & IP & "', " & SLT & ", " & PRT & ",'" & FDT & "','" & TM & "')"
                    ' SQLEXEC(ens)

                Else
                    VAZ = "EMPTY"
                End If

            End If
        Next

        OTCPStream.Close()
        OTCP.Close()





    End Sub
    Public Sub FIBERT20(ByVal IP As String, ByVal hua As Boolean)
        Dim ens, strtmp, TL, MRK As String
        Dim USNR, DSNR, TXST As Integer
        Dim URATE, DRATE, UATT, DATT As Integer
        Dim STL, SL, PRT, L As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        Dim INSTREE As System.Data.Sqlclient.SqlCommand
        ens = IP
        ens = IP
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        OTCP.SendTimeout = 3000
1:
        MRK = MRKAZ(IP)
        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        System.Threading.Thread.Sleep(1000)
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        '        TextBox1.Text = RSTR()
        'ListBox1.Items.Add(RSTR())
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("GEPON" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(strtmp)
        TextBox1.Text = strtmp
        WriteData("CD DSL" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        'استخراج اطلاعات 
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        cnn.Open()
        Dim RDATA As System.Data.Sqlclient.SqlDataReader
        Dim RDA As System.Data.Sqlclient.SqlDataReader
        RDATA = New System.Data.Sqlclient.SqlCommand("SELECT COUNT(SLOT)AS CNT FROM DSLSLOT WHERE IP='" & IP & "'", cnn).ExecuteReader
        Do While RDATA.Read
            MXPSLOT = RDATA!CNT
        Loop
        ProgressBar2.Maximum = MXPSLOT
        ProgressBar2.Minimum = 1
        RDATA.Close()
        RDATA = New System.Data.Sqlclient.SqlCommand("SELECT SLOT,STL FROM DSLSLOT WHERE IP='" & IP & "'", cnn).ExecuteReader
        Do While RDATA.Read
            SL = RDATA!SLOT
            STL = RDATA!STL
            For I = 1 To 64
                'if tel exist and
                PRT = I
                TL = TELNO(STL, PRT, IP)
                TXST = TLEXIST(TL)

                If TL <> "Unknown" And TXST = 0 Then
                    'یعنی باید استخراج شود
                    strtmp = ""
                    WriteData("SHOW DSL port status slot " & SL & " port " & I & vbCrLf, OTCPStream) 'Options
                    System.Threading.Thread.Sleep(200)
                    strtmp = RSTR(OTCPStream, OTCP)
                    L = 0
                    L = InStr(strtmp, "closed")
                    If L > 4 Then
                        WriteData(vbCrLf, OTCPStream) 'Options
                        System.Threading.Thread.Sleep(200)
                        GoTo 1
                    End If
                    L = Microsoft.VisualBasic.InStr(strtmp, "state")
                    '   L = Microsoft.VisualBasic.InStr("state", TextBox1.Text)
                    L += 22
                    Dim sss As String
                    sss = Mid(strtmp, L, 8)
                    If sss = "ShowTime" Then
                        'خواندن اطلاعات
                        L = InStr(strtmp, "USsnrMargin")
                        L += 22
                        USNR = Val(Mid(strtmp, L, 4))
                        L = InStr(strtmp, "DSsnrMargin")
                        L += 22
                        DSNR = Val(Mid(strtmp, L, 4))
                        L = InStr(strtmp, "USLineBitRate")
                        L += 22
                        URATE = Val(Mid(strtmp, L, 5))
                        L = InStr(strtmp, "DSLineBitRate")
                        L += 22
                        DRATE = Val(Mid(strtmp, L, 5))
                        L = InStr(strtmp, "AggLnAttenDs")
                        L += 22
                        DATT = 10 * Val(Mid(strtmp, L, 2))
                        L = InStr(strtmp, "AggLnAttenUs")
                        L += 22
                        UATT = 10 * Val(Mid(strtmp, L, 2))
                        PRT = I
                        'دراینجا باید اطلاعات ذخیره شوند
                        'Dim TL As String
                        TL = TELNO(STL, I, IP)
                        ens = "INSERT INTO TLOG (TEL, STL, PRT, URATE, DRATE, USNR, DSNR, UATT, DATT, FDT, HR, MARKAZ, SHAHR) VALUES (N'" & TL & "'," & STL & "," & I & "," & URATE & "," & DRATE & "," & USNR & "," & DSNR & "," & UATT & "," & DATT & ",'" & fnow() & "','" & HRNOW() & "','" & MRKAZ(IP) & "', '" & SHR & "')"
                        L = SQLEXEC(ens)
                        ens = "UPDATE TELS SET MDT ='" & fnow() & "' WHERE TEL ='" & TL & "'"
                        SQLEXEC(ens)


                    End If
                Else
                    WriteData(vbCrLf, OTCPStream) 'Options
                    'System.Threading.Thread.Sleep(100)
                    strtmp = RSTR(OTCPStream, OTCP)
                    L = L
                End If

                If ProgressBar3.Value < ProgressBar3.Maximum Then ProgressBar3.Value = I
            Next I
        Loop
        cnn.Close()
        '        My.Computer.FileSystem.CopyFile(My.Computer.FileSystem.CurrentDirectory & "\Data.sdf", My.Computer.FileSystem.CurrentDirectory & "\Report.sdf", True)
        OTCPStream.Close()
        OTCP.Close()

    End Sub

    Private Sub tmp()
        Dim FD As String = "9"
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        'Timer1.Enabled = False
        On Error Resume Next
        DELMONTH()
        cnn.Open()
        '**************************************
        Dim ACTV As Boolean
        Dim MDL, IP, COD, MDT, SHR As String
        MDL = ""
        MDT = "0"
        SHR = "0"
        IP = ""
        COD = ""
        'Timer1.Stop()
        'ACTV = True
        Dim RDATA As System.Data.Sqlclient.SqlDataReader
        RDATA = New System.Data.Sqlclient.SqlCommand("SELECT IP,TYPE,CODE,HR FROM CONFIG", cnn).ExecuteReader
        Do While RDATA.Read
            MDL = RDATA!TYPE
            IP = CStr(RDATA!IP)
            COD = RDATA!CODE
            SHR = RDATA!HR
        Loop
        '     '*****ACTIVATION CHECK
        ACTV = CHKACTV(COD)
        If ACTV = True Then
            RDATA = New System.Data.Sqlclient.SqlCommand("SELECT MAX(FDT) AS MDAT FROM TLOG", cnn).ExecuteReader
            Do While RDATA.Read
                MDT = RDATA!MDAT
            Loop
            If MDT < fnow() Then
                If SHR <= HRNOW() Then
                    'Timer1.Stop()
                    If IP <> "" Then
                        Select Case MDL
                            Case "ZTE"
                                ZTEMAIN()
                            Case "HUAWEI"
                                '                               HUAWEIMAIN(IP)
                        End Select
                        'Timer1.Start()
                        Shell(My.Computer.FileSystem.CurrentDirectory & "\SNRAssist.exe")
                        System.Threading.Thread.Sleep(5000)
                        Me.Close()
                    Else
                        MsgBox("لطفاً با مراجعه به قسمت تنظیمات اطلاعات مربوطه را به درستی وارد کنید", MsgBoxStyle.Critical, "پیام نرم افزار")
                    End If
                End If
            End If
        End If
        cnn.Close()
        'Here to alcholing
        'Timer1.Start()
        'Timer1.Enabled = True
        ' Else
        'If Now.Year = 2016 Then
        '        MsgBox("نرم افزار نیاز به فعال سازی دارد. لطفاً مراتب از طریق معاونت بازاریابی و فروش پیگیری گردد", , "پیام نرم افزار")
        '       Else
        '      MsgBox("Unknown error has been occured! Windows critically closed the applicatin", MsgBoxStyle.Critical, "Unknown error")
        '     End If
        '    End If
        '
        '**************************************
        '        RDATA = New System.Data.Sqlclient.SqlCommand("SELECT MAX(FDT) AS MDAT FROM IPLOG", cnn).ExecuteReader
        '       Do While RDATA.Read
        'FD = RDATA!MDAT
        '        Loop
        '
        '
        '        If FD < fnow() Then
        'Timer1.Enabled = False
        '        Call Button2_Click(Me, e)
        '       Timer1.Enabled = True
        '      End If
        'ZTEMAIN()

    End Sub
    Private Sub BROOZ()
        Dim LDT, ENS, FN As String
        Dim INSTREE As System.Data.SqlClient.SqlCommand
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        cnn.Open()
        '        Dim RDATA As System.Data.Sqlclient.SqlDataReader
        '       RDATA = New System.Data.Sqlclient.SqlCommand("SELECT COUNT(SLOT)AS CNT FROM DSLSLOT WHERE IP='" & IP & "'", cnn).ExecuteReader
        'My.Computer.FileSystem.CopyFile(My.Computer.FileSystem.CurrentDirectory & "\Data.sdf", My.Computer.FileSystem.CurrentDirectory & "\Report.sdf", True)

        ' Panel1.Visible = False
        '********************************
        On Error Resume Next
        Dim IPS, TYPS, MMDT As String
        Dim RT As System.Data.SqlClient.SqlDataReader
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        MMDT = " "
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT MAX(FDT) AS MDT FROM PROFSHR", cnn).ExecuteReader
        Do While RDATA.Read
            MMDT = CStr(RDATA!MDT)
        Loop
        RDATA.Close()
        ' FN = fnow()

        If MMDT < DRZZ Then
            'PROFILE 
            ENS = "DELETE FROM PROFSHR"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()
            ENS = "INSERT INTO PROFSHR (CNT, DSNR, SHAHR, DDR, FDT) SELECT COUNT(TEL) AS CNT, AVG(DSNR) AS DSNR, SHAHR, DDR, '" & DRZZ & "' AS FDT FROM (SELECT        TEL, AVG(USNR) AS USNR, AVG(DSNR) AS DSNR, AVG(UATT) AS UATT, AVG(DATT) AS DATT, AVG(URATE) AS URATE, AVG(DRATE) AS DRATE, MARKAZ, SHAHR, AVG(DRATE) / 1000 AS DDR, MAX(FDT) AS FDT FROM TLOG " & _
" WHERE (DSNR > 119) AND (FDT = '" & DRZZ & "') GROUP BY TEL, MARKAZ, SHAHR) AS DSLG GROUP BY SHAHR, DDR"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()
            '**************
            'PROFILE 
            ENS = "DELETE FROM PROFALL"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()
            ENS = "INSERT INTO PROFALL (CNT, DSNR, SHAHR, DDR, FDT) SELECT COUNT(TEL) AS CNT, AVG(DSNR) AS DSNR, SHAHR, DDR, '" & DRZZ & "' AS FDT FROM (SELECT        TEL, AVG(USNR) AS USNR, AVG(DSNR) AS DSNR, AVG(UATT) AS UATT, AVG(DATT) AS DATT, AVG(URATE) AS URATE, AVG(DRATE) AS DRATE, MARKAZ, SHAHR, AVG(DRATE) / 1000 AS DDR, MAX(FDT) AS FDT FROM TLOG " & _
" WHERE  (FDT = '" & DRZZ & "') GROUP BY TEL, MARKAZ, SHAHR) AS DSLG GROUP BY SHAHR, DDR"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()

            'SHAHR AVG
            ENS = "DELETE FROM SHRAVG"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()
            ENS = "INSERT INTO SHRAVG (USNR, DSNR, UATT, DATT, URATE, DRATE, SHAHR, FDT) SELECT AVG(USNR) AS USNR, AVG(DSNR) AS DSNR, AVG(UATT) AS UATT, AVG(DATT) AS DATT, AVG(URATE) AS URATE, AVG(DRATE) AS DRATE, SHAHR, '" & DRZZ & "' AS FDT " & _
" FROM TLOG WHERE (FDT = '" & DRZZ & "') GROUP BY SHAHR ORDER BY SHAHR"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()

            'MARKAZ AVG
            ENS = "DELETE FROM MRAVG"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()
            ENS = "INSERT INTO MRAVG (USNR, DSNR, UATT, DATT, URATE, DRATE, SHAHR, MARKAZ, FDT, CNT) SELECT AVG(USNR) AS USNR, AVG(DSNR) AS DSNR, AVG(UATT) AS UATT, AVG(DATT) AS DATT, AVG(URATE) AS URATE, AVG(DRATE) AS DRATE, SHAHR, MARKAZ, '" & DRZZ & "' AS FDT, COUNT(TEL) AS CNT " & _
" FROM TLOG WHERE (FDT = '" & DRZZ & "') GROUP BY SHAHR, MARKAZ ORDER BY SHAHR, MARKAZ"
            INSTREE = New System.Data.SqlClient.SqlCommand(ENS, cnn)
            INSTREE.ExecuteNonQuery()

        End If
        cnn.Close()
    End Sub
    Public Sub ZTEMAIN()
        Dim ens, strtmp, TL, MRK, MDT, SHR, HNAM As String
        Dim STL, SL, L, UATT, DATT, URATE, DRATE As Integer
        Dim USNR, DSNR As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        'Timer1.Stop()
        MDT = "0"
        ' Dim sw1 As New IO.StreamWriter(My.Computer.FileSystem.SpecialDirectories.Temp & "\HOA" & ZZ & ".TXT", False)
        SHR = "0"
        '********************************************************************************************************************************
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        cnn.Open()
        Dim IPS, TYPS, MMDT As String
        Dim RT As System.Data.SqlClient.SqlDataReader
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT MAX(MDT) AS MDT FROM TELS", cnn).ExecuteReader
        Do While RDATA.Read
            MMDT = CStr(RDATA!MDT)
        Loop
        RDATA.Close()
        RT = New System.Data.SqlClient.SqlCommand("SELECT COUNT(IP) AS CNT FROM DSLAMS", cnn).ExecuteReader
        Do While RT.Read
            MXPDSLAM = RT!CNT
        Loop
        RT.Close()
        ProgressBar1.Maximum = MXPDSLAM
        ProgressBar1.Value = 1
        Do
            RT = New System.Data.SqlClient.SqlCommand("SELECT IP, HNAM, DEVT FROM GDOTDEVS where DEVT ='ZXAN'", cnn).ExecuteReader
            Do While RT.Read
                IPS = Trim(RT!IP)
                HNAM = Trim(RT!HNAM)
                TYPS = Trim(RT!DEVT)
                If IPS <> "" Then
                    Select Case TYPS
                        Case "MA-5600"
                            HUAWEISNR(IPS, HNAM)
                        Case "MA-5600T"
                            HUAWEISNR(IPS, HNAM)
                        Case "FSAP-9800"
                            '    ZTESNR(IPS, HNAM)
                        Case "AN5006-20"
                            FIBERSNR(IPS, HNAM)
                        Case "AN5006-30"
                            FIBERSNR(IPS, HNAM)
                        Case "MA-5616"
                            '   HUA5616(IPS, HNAM)
                        Case "FSAP-9806"
                            Z9806(IPS, HNAM)
                        Case "ZXAN"
                            '  ZXAN(IPS, HNAM)
                            'ZTE
                    End Select
                Else
                    MsgBox("لطفاً اطلاعات مربوط دی اس لم ها را به درستی وارد کنید", MsgBoxStyle.Critical, "نرم افزار سپهر")
                End If
2:
                '  If ListBox1.Items.Count > 150 Then Exit Do
            Loop
            RT.Close()
            'چـــــــــــــــــــــــــــــــــــــراااااااااااااا
            'My.Computer.FileSystem.CopyFile(My.Computer.FileSystem.CurrentDirectory & "\Data.sdf", My.Computer.FileSystem.CurrentDirectory & "\Report.sdf", True)
        Loop
        cnn.Close()


        'Save to file

        'Timer1.Start()
        'Catch Err As Exception
        ' MsgBox(Err.ToString)
        ' End Try


    End Sub

    Private Sub Z9806(ByVal IP As String, ByVal HST As String)
        Dim ens, strtmp, TL, MRK As String
        'Dim IP As String = "172.27.178.140"
        Dim USNR, DSNR, TXST As Integer
        Dim URATE, DRATE, UATT, DATT As Integer
        Dim STL, SL, PRT, L, LL, SLT As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        Dim a() As String
        Dim INSTREE As System.Data.SqlClient.SqlCommand
        ens = IP
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        On Error Resume Next
        OTCP.SendTimeout = 3000
1:

        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)

        System.Threading.Thread.Sleep(500)
        '        strtmp = RSTR(OTCPStream, OTCP)
        WriteData(vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(200)
        WriteData(vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(200)
        strtmp = RSTR(OTCPStream, OTCP)
        '        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(strtmp)
        TextBox1.Text = strtmp
        WriteData("show card" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(500)
        strtmp = RSTR(OTCPStream, OTCP)
        RichTextBox1.AppendText(strtmp)

        'استخراج اطلاعات 
        'MsgBox("finished")
        For Each lin In RichTextBox1.Lines
            If InStr(lin, "SSTEB") > 0 Then
                SLT = Val(Mid(lin, 1, 5))
                If Not EXIST(IP, "SSTEB", SLT) Then
                    ens = "INSERT INTO GDOTSLOTS(IP, TYP, SLOT) VALUES('" & IP & "','SSTEB'," & SLT & ")"
                    SQLEXEC(ens)
                End If
            End If
        Next
        OTCPStream.Close()
        OTCP.Close()



    End Sub
    Private Function EXIST(ByVal IP As String, ByVal TYP As String, ByVal SLT As Integer) As Boolean
        Dim C As Integer
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        cnn.Open()
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT COUNT(IP) AS CNT FROM GDOTSLOTS WHERE IP='" & IP & "' AND TYP='" & TYP & "' AND SLOT = " & SLT, cnn).ExecuteReader
        Do While RDATA.Read
            C = CInt(RDATA!CNT)
        Loop
        RDATA.Close()
        cnn.Close()
        If C > 0 Then EXIST = True Else EXIST = False
    End Function

 
    Public Sub ZTEMINI(ByVal IP As String)
        Dim ens, strtmp, TL, MRK As String
        Dim SL, STL, L, UATT, DATT, URATE, DRATE As Integer
        Dim USNR, DSNR, PRT, TXST As Integer
        Dim MXP1, MNP1, MXP2, MNP2 As Integer

        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient
        ' Dim sw As New IO.StreamWriter("G:\TXTS\" & ens & ".TXT")
        On Error Resume Next
        'Try
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        OTCP.SendTimeout = 3000
        MRK = MRKAZ(IP)
        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        If Trim(TextBox1.Text) = "" Then
            DSACC(IP, False)
            Exit Sub
        Else
            DSACC(IP, True)
        End If
        System.Threading.Thread.Sleep(2000)
        TextBox1.Text = ""
        'OTCPStream = OTCP.GetStream
        ' TextBox1.Text = RSTR(OTCPStream, OTCP)
        WriteData(vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(0)

        System.Threading.Thread.Sleep(1000)
        WriteData("admin" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("en" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(strtmp)
        WriteData("show card" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        '********************************
        Dim SLT As Integer
        Dim TYP, TSTR As String
        Dim sw1 As New System.IO.StreamWriter(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")
        TextBox1.Text = strtmp
        sw1.AutoFlush = True
        sw1.Write("EXECUTE NEW LINE " & TextBox1.Text)
        System.Threading.Thread.Sleep(2000)
        sw1.Close()
        Dim sw2 As New IO.StreamReader(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")
        Dim CNTR As Integer
        CNTR = 0
        While Not sw2.EndOfStream
            TSTR = sw2.ReadLine()
            If TSTR Is Nothing Then CNTR += 1
            If CNTR > 200 Then Exit Sub
            If Val(Mid(TSTR, 1, 5)) > 0 Then
                SLT = Val(Mid(TSTR, 1, 5))
                TYP = Trim(Mid(TSTR, 13, 8))
                STL = GTSTL(IP, SLT)
                If STL < 1900 Then
                    ens = "INSERT INTO STLTYP (MARKAZ, DSTYP, STL, SLOT, SLTYP, IP, SHAHR) VALUES ('" & MRK & "','ZTE9806'," & STL & "," & SLT & ",'" & TYP & "','" & IP & "','" & SHR & "')"
                    SQLEXEC(ens)
                End If
            End If
        End While
        sw2.Close()
        My.Computer.FileSystem.DeleteFile(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")

        OTCPStream.Close()
        OTCP.Close()
    End Sub
    Public Sub HUA5616(ByVal IP As String)
        'On Error Resume Next
        Dim ens, strtmp, TSTR, TL, MRK As String
        Dim SL, STL, L, UATT, DATT, URATE, DRATE, PRT As Integer
        Dim USNR, DSNR, SLTCNT, TXST As Integer
        Dim MXPDSLAM, MXPSLOT, MXPPORT As Integer
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient

        'IP = "136.1.1.100"
        System.Threading.Thread.Sleep(1000)
        On Error Resume Next
        TextBox1.Text = ""
        OTCP.SendTimeout = 3000
        MRK = MRKAZ(IP)
        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        If Trim(TextBox1.Text) = "" Then
            DSACC(IP, False)
            Exit Sub
        Else
            DSACC(IP, True)
        End If
        System.Threading.Thread.Sleep(1000)
        WriteData("root" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("mduadmin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(2000)
        strtmp = RSTR(OTCPStream, OTCP)
        Dim LL As Integer
        LL = 0
        LL = InStr(strtmp, "invalid")
        If LL > 0 Then DSACC(IP, False) : Exit Sub

        'ListBox1.Items.Add(strtmp)
        TextBox1.Text = strtmp
        WriteData("en" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        WriteData("display board 0" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)

        Dim SLT As Integer
        Dim TYP As String
        Dim sw1 As New System.IO.StreamWriter(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")
        TextBox1.Text = strtmp
        sw1.AutoFlush = True
        sw1.Write("EXECUTE NEW LINE " & TextBox1.Text)
        System.Threading.Thread.Sleep(2000)
        sw1.Close()
        Dim sw2 As New IO.StreamReader(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")
        Dim CNTR As Integer
        CNTR = 0
        While Not sw2.EndOfStream
            TSTR = sw2.ReadLine()
            If TSTR Is Nothing Then CNTR += 1
            If CNTR > 200 Then Exit Sub
            If Val(Mid(TSTR, 1, 5)) > 0 Then
                SLT = Val(Mid(TSTR, 1, 5))
                TYP = Trim(Mid(TSTR, 10, 10))
                STL = GTSTL(IP, SLT)
                If STL < 1900 Then
                    ens = "INSERT INTO STLTYP (MARKAZ, DSTYP, STL, SLOT, SLTYP, IP, SHAHR) VALUES ('" & MRK & "','Huawei5616'," & STL & "," & SLT & ",'" & TYP & "','" & IP & "','" & SHR & "')"
                    SQLEXEC(ens)
                End If
            End If
        End While
        sw2.Close()
        My.Computer.FileSystem.DeleteFile(My.Computer.FileSystem.CurrentDirectory & "\CRD" & ZZ & ".TXT")

        OTCPStream.Close()
        OTCP.Close()

    End Sub
    Public Sub ZTESNR(ByVal IP As String)
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim SLT, PRT As Integer
        Dim DES, LST, STAT, PVC, FDT, TM As String
        Dim VLAN, VPI, VCI As Integer
        Dim ens, strtmp, TL, MRK, TSTR As String
        Dim STL, SL, L, UATT, DATT, URATE, DRATE, TXST As Integer
        Dim USNR, DSNR As Integer
        Dim MXP1, MNP1, MXP2, MNP2 As Integer
        Dim LL As Integer
        Dim a() As String
        Dim DT As Date
        Dim PTM As Boolean = False
        'On Error Resume Next
        Dim OTCP As System.Net.Sockets.TcpClient
        Dim OTCPStream As System.Net.Sockets.NetworkStream
        OTCP = New System.Net.Sockets.TcpClient
        ens = IP
        ens = ens.Replace(".", "_")
        ens = "ZTE" & ens & Now.Year.ToString & "_" & Now.Month.ToString & "_" & Now.Day.ToString
        On Error Resume Next
        'Try
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = ""
        OTCP.SendTimeout = 3000
        MRK = MRKAZ(IP)
        OTCP.Connect(IP, "23")
        OTCPStream = OTCP.GetStream
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        If Trim(TextBox1.Text) = "" Then
            DSACC(IP, False)
            Exit Sub
        Else
            DSACC(IP, True)
        End If
        System.Threading.Thread.Sleep(1000)
        WriteData("admin" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        WriteData("en" & vbCrLf, OTCPStream) 'Account
        System.Threading.Thread.Sleep(1000)
        TextBox1.Text = RSTR(OTCPStream, OTCP)
        'ListBox1.Items.Add(RSTR(OTCPStream, OTCP))
        '            TextBox1.Text = TextBox1.Text & ReadData() & vbCrLf, OTCPStream
        WriteData("admin" & vbCrLf, OTCPStream) 'Password
        System.Threading.Thread.Sleep(1000)
        strtmp = RSTR(OTCPStream, OTCP)
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT * FROM DSCARDS WHERE IP ='" & IP & "' AND [CARD] LIKE '%SH%'", cnn).ExecuteReader
        Dim sw1 As New System.IO.StreamWriter(My.Computer.FileSystem.CurrentDirectory & "\HOAcard.TXT")
        Dim sw2 As New IO.StreamReader(My.Computer.FileSystem.CurrentDirectory & "\HOAcard.TXT")
        Do While RDATA.Read
            SLT = CInt(RDATA!SLT)
            'NOW START TO SCAN PORTS
            For I = 1 To 32
                VLAN = 0 : DES = "" : STAT = "" : PVC = "" : FDT = "" : TM = ""
                System.Threading.Thread.Sleep(500)
                TextBox1.Text = ""
                WriteData("show interface shdsl_0/" & SLT & "/" & I & vbCrLf, OTCPStream) 'Password
                System.Threading.Thread.Sleep(50)
                WriteData(" ", OTCPStream) 'Password
                System.Threading.Thread.Sleep(50)
                WriteData(" ", OTCPStream) 'Password
                System.Threading.Thread.Sleep(50)
                strtmp = RSTR(OTCPStream, OTCP)
                TextBox1.Text = TextBox1.Text & strtmp
                If Trim(TextBox1.Text) = "" Then GoTo NOFILE
                RichTextBox1.Clear()
                RichTextBox1.AppendText(TextBox1.Text)
                a = RichTextBox1.Lines.ToArray()
                'RichTextBox1.Lines.TO()
                Dim N As Integer
                N = RichTextBox1.Lines.Length
                For j = 0 To N
                    TSTR = a(j)
                    LL = 0
                    LL = InStr(TSTR, "Description")
                    If LL > 0 Then
                        TSTR = Mid(TSTR, LL + 1)
                        LL = 0
                        LL = InStr(TSTR, "is")
                        If LL > 0 Then
                            TSTR = Trim(Mid(TSTR, LL + 3))
                            DES = TSTR
                        End If
                    End If
                    LL = 0
                    ' Shdsl  0/ 0/ 0 is down
                    LL = InStr(TSTR, "line protocol is")
                    If LL > 0 Then
                        ' TSTR = Mid(TSTR, LL + 17)
                        ' PRT = Val(TSTR)

                        'LL = 0
                        'LL = InStr(TSTR, "is")
                        'If LL > 0 Then
                        TSTR = Trim(Mid(TSTR, LL + 17))
                        STAT = TSTR
                        'End If
                    End If
                    LL = 0
                    'Last up time is 2019-2-18 12:59:17.
                    LL = InStr(TSTR, "Last up time")
                    If LL > 0 Then
                        TSTR = Trim(Mid(TSTR, LL + 16))
                        LST = TSTR
                        If Len(LST) > 5 Then
                            DT = CDate(Microsoft.VisualBasic.Left(LST, 10))
                            TM = Microsoft.VisualBasic.Left(Trim(Mid(TSTR, 11)), 8)
                            FDT = FDATE(DT)
                        Else
                            FDT = "NEVER"
                            TM = "NEVER"
                        End If

                    End If
                    LL = 0
                    LL = InStr(TSTR, "PTM")
                    If LL > 0 Then
                        PTM = True
                        PVC = "PTM"
                    End If
                    LL = 0
                    LL = InStr(TSTR, "PVC  ")
                    If LL > 0 Then
                        TSTR = Trim(Mid(TSTR, LL + 4))
                        PVC = TSTR
                    End If
                    LL = 0
                    LL = InStr(TSTR, "VLAN")
                    If LL > 0 Then
                        If PTM Or PVC = "0/33" Or PVC = "8/35" Or PVC = "6/35" Or PVC = "0/34" Then
                            TSTR = Mid(TSTR, LL + 1)
                            LL = 0
                            LL = InStr(TSTR, ":")
                            If LL > 0 Then
                                TSTR = Trim(Mid(TSTR, LL + 1))
                                VLAN = Val(TSTR)
                                Exit For
                            End If
                        End If
                    End If

                Next j
                If VLAN > 0 Then
                    ens = "INSERT INTO GDOT(STAT, DES, PVC, VLAN, LST, IP, SLT, PRT,FDT,FTM) VALUES (N'" & STAT & "', N'" & DES & "', '" & PVC & "', " & VLAN & ", N'" & LST & "', N'" & IP & "', " & SLT & ", " & PRT & ",'" & FDT & "','" & TM & "')"
                    SQLEXEC(ens)
                End If
                PTM = False
                VLAN = 0

NOFILE:
            Next

        Loop
        RDATA.Close()

        OTCPStream.Close()
        OTCP.Close()

    End Sub
    Private Function SQLEXEC(ByVal cmd As String) As Integer
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim INSTREE As System.Data.SqlClient.SqlCommand
        cnn.Open()
        INSTREE = New System.Data.SqlClient.SqlCommand(cmd, cnn)
        SQLEXEC = INSTREE.ExecuteNonQuery()
        cnn.Close()
    End Function

    Private Function TLEXIST(ByVal TLL As String) As Integer
        On Error Resume Next
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim MDT As String = ""
        cnn.Open()
        Dim RDA As System.Data.SqlClient.SqlDataReader
        RDA = New System.Data.SqlClient.SqlCommand("SELECT MDT FROM TELS WHERE TEL='" & TLL & "'", cnn).ExecuteReader
        Do While RDA.Read
            MDT = CStr(RDA!MDT)
        Loop
        If MDT < fnow() Then
            TLEXIST = 0
        Else
            TLEXIST = 1
        End If
        RDA.Close()
        cnn.Close()
    End Function
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SEEKS = 0
        'oTCP.SendTimeout = 3000
        'TODO: This line of code loads data into the 'DataDataSet.TELLOG' table. You can move, or remove it, as needed.
        '        My.Computer.FileSystem.CopyFile(My.Computer.FileSystem.CurrentDirectory & "\Data.sdf", My.Computer.FileSystem.CurrentDirectory & "\Report.sdf", True)
        '   ZTEMAIN("136.1.1.100")
        'Application.Run(STRY)
        ' Call Timer1_Tick(Me, e)
        'ZTEMAIN()
        FIBERSNR("172.27.16.120", "FH30-Najaf-Emam-DS1-27-16-120")
    End Sub







    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ZTEMAIN()
    End Sub
    Private Function CHKACTV(ByVal CD As String) As Boolean
        'If CHECKSERIAL(CD) > 0 Then
        'CHKACTV = True
        'Else
        'CHKACTV = False
        'End If
        CHKACTV = True
    End Function
    Private Function TELNO(ByVal STL As Integer, ByVal PRT As Integer, ByVal IP As String) As String
        ' On Error Resume Next
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        cnn.Open()
        '**************************************
        Dim TL As String
        Dim MRKZ, SSTR As String
        TL = "Unknown"
        MRKZ = MRKAZ(IP)
        SSTR = "SELECT TEL FROM TELS WHERE ((STL=" & STL & " AND PRT=" & PRT & ") AND (MARKAZ = '" & MRKZ & "'))"
        Dim RDA As System.Data.SqlClient.SqlDataReader
        RDA = New System.Data.SqlClient.SqlCommand("SELECT TEL FROM TELS WHERE ((STL=" & STL & " AND PRT=" & PRT & ") AND (MARKAZ = '" & MRKZ & "'))", cnn).ExecuteReader
        Do While RDA.Read
            TL = CStr(RDA!TEL)

        Loop
        RDA.Close()
        cnn.Close()
        TELNO = TL
    End Function
    Private Function MRKAZ(ByVal IP As String) As String
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR

        cnn.Open()
        '**************************************
        Dim MRKZ As String
        MRKZ = "Unknown"
        Dim RDATA As System.Data.Sqlclient.SqlDataReader
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT MARKAZ,SHAHR FROM DSLAMS WHERE IP ='" & IP & "'", cnn).ExecuteReader
        Do While RDATA.Read
            MRKZ = RDATA!MARKAZ
            SHR = RDATA!SHAHR
        Loop
        cnn.Close()
        MRKAZ = MRKZ
    End Function
    Private Sub DELMONTH()
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim DTT As String
        DTT = fnow()
        cnn.Open()
        '**************************************
        Dim MRKZ As String
        MRKZ = "Unknown"
        Dim RDATA As System.Data.Sqlclient.SqlCommand
        RDATA = New System.Data.Sqlclient.SqlCommand("DELETE FROM TLOG WHERE FDT < '" & DTT & "'", cnn)
        RDATA.ExecuteNonQuery()
        cnn.Close()

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If WD = True Then
            Shell(My.Computer.FileSystem.CurrentDirectory & "\SNRAssist.exe")
            System.Threading.Thread.Sleep(5000)
            Me.Close()
        End If
    End Sub


End Class