Imports System.Environment
Imports System
Imports System.Net
Imports System.Net.Mail
Imports System.Management
Imports System.ComponentModel
Imports System.Net.Mail.SmtpClient
Module Module1
    Dim CAL As New System.Globalization.PersianCalendar
    Public SNDING As Boolean = False
    Public ALRM As Boolean = False
    Public TR As Boolean
    Public BCK As Boolean
    Public MIL As String
    Public usrnam As String
    Public DYS As Integer
    Public untcod As String
    Public unt As String
    Public ARM As Byte
    Public USRCOD As Long
    Public FDTSTR As String
    Public DRZZ As String
    Const Z1 As String = " IP < '172.27.1.3'"
    Const Z2 As String = "IP > '172.27.1.255' AND IP < '172.27.147.11'"
    Const Z3 As String = "IP > '172.27.147.109' AND IP < '172.27.18.120  '"
    Const Z4 As String = "IP > '172.27.18.119' AND IP < '172.27.2.46'"
    Const Z5 As String = "IP > '172.27.2.45' AND IP < '172.27.22.4'"
    Const Z6 As String = "IP > '172.27.22.3' AND IP < '172.27.25.82'"
    Const Z7 As String = "IP > '172.27.25.81' AND IP < '172.27.3.108'"
    Const Z8 As String = "IP > '172.27.3.107' AND IP < '172.27.30.147'"
    Const Z9 As String = "IP > '172.27.30.146' AND IP < '172.27.5.72'"
    Const Z10 As String = "IP > '172.27.5.71'" ' AND IP < '172.27.147.12'"
    Public CONDIT As String = Z10
    Public ZZ As String = "Z10"
    Public CNNSTR As String = "Data Source=WIN-LNVNRPVDCM1;Initial Catalog=SNRDATA;Integrated Security=True"
    Public RPTSTR As String = "Data Source=WIN-LNVNRPVDCM1;Initial Catalog=SNRDATA;Integrated Security=True"
    Dim TMR As New Timer
    Public Sub trr()
        'On Error Resume Next
        '       frmtray.Show()
        '       frmtray.Hide()
    End Sub
    Public Function fnow() As String
        Dim y, m, d As Integer
        Dim sy, sm, sd As String
        y = CAL.GetYear(Now)
        m = CAL.GetMonth(Now)
        d = CAL.GetDayOfMonth(Now)
        sy = Str(y)
        If m < 10 Then sm = "0" & Trim(Str(m)) Else sm = Trim(Str(m))
        If d < 10 Then sd = "0" & Trim(Str(d)) Else sd = Trim(Str(d))
        fnow = Trim(sy & "/" & sm & "/" & sd)

    End Function
    Public Function FDATE(ByVal DT As Date) As String
        Dim y, m, d As Integer
        Dim sy, sm, sd As String
        y = CAL.GetYear(DT)
        m = CAL.GetMonth(DT)
        d = CAL.GetDayOfMonth(DT)

        '   If m > 1 Then m -= 1 Else m = 12 : y -= 1
        sy = Str(y)
        If m < 10 Then sm = "0" & Trim(Str(m)) Else sm = Trim(Str(m))
        If d < 10 Then sd = "0" & Trim(Str(d)) Else sd = Trim(Str(d))
        FDATE = Trim(sy & "/" & sm & "/" & sd)

    End Function

    Public Function HRNOW() As String
        Dim H, M As Integer
        Dim SH, SM As String
        H = Now.Hour
        M = Now.Minute
        If H < 10 Then SH = "0" & H.ToString Else SH = H.ToString
        If M < 10 Then SM = "0" & M.ToString Else SM = M.ToString
        HRNOW = SH & ":" & SM
    End Function
    Public Sub SETFARS(ByVal FARS As Boolean)
        If FARS = True Then
            If Microsoft.VisualBasic.Left(System.Windows.Forms.Application.CurrentInputLanguage.Culture.EnglishName, 1) = "E" Then
                My.Computer.Keyboard.SendKeys("%+")
            End If
        Else
            If Microsoft.VisualBasic.Left(System.Windows.Forms.Application.CurrentInputLanguage.Culture.EnglishName, 1) <> "E" Then
                My.Computer.Keyboard.SendKeys("%+")
            End If
        End If
    End Sub
    Public Sub CAPS(ByVal CPSLOCK As Boolean)
        If My.Computer.Keyboard.CapsLock <> CPSLOCK Then
            My.Computer.Keyboard.SendKeys("{CAPSLOCK}")
        End If
    End Sub
    Public Sub TASKUPDATE()
        On Error Resume Next
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim CMD As String
        CMD = "INSERT INTO TODO (TEL, PUSNR, PDSNR, PUATT, PDATT, PURATE, PDRATE, DTP, SHAHR)" & _
"SELECT TEL, USNR, DSNR, UATT, DATT, URATE, DRATE, FDT, SHAHR FROM RLOG" & _
" WHERE (TEL NOT IN (SELECT TEL FROM TODO AS TODO_2)) AND (DSNR < 120) OR" & _
" (TEL NOT IN (SELECT TEL FROM TODO AS TODO_1)) AND (DRATE < 2000)"
        '**************************************
        cnn.Open()
        Dim RDATA As System.Data.SqlClient.SqlCommand
        RDATA = New System.Data.SqlClient.SqlCommand(CMD, cnn)
        RDATA.ExecuteNonQuery()
        cnn.Close()

    End Sub
    Public Sub DONEUPDATE()
        On Error Resume Next
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim CMD As String
        CMD = "UPDATE TODO SET LUSNR = RLOG.USNR, LDSNR = RLOG.DSNR, LUATT = RLOG.UATT, LDATT = RLOG.DATT, LURATE = RLOG.URATE, LDRATE = RLOG.DRATE, VAZ = N'انجام شده', DTD = RLOG.FDT " & _
" FROM TODO INNER JOIN RLOG ON TODO.TEL = RLOG.TEL WHERE (RLOG.DSNR > 119 AND RLOG.DRATE > 2000)"
        cnn.Open()
        Dim RDATA As System.Data.SqlClient.SqlCommand
        RDATA = New System.Data.SqlClient.SqlCommand(CMD, cnn)
        RDATA.ExecuteNonQuery()
        CMD = "DELETE FROM TODO WHERE VAZ='انجام شده' AND WRK IS NULL"
        RDATA = New System.Data.SqlClient.SqlCommand(CMD, cnn)
        RDATA.ExecuteNonQuery()
        cnn.Close()
    End Sub
    Public Function DRZ() As String
        Dim CNT As Integer
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim RDATA As System.Data.SqlClient.SqlDataReader
        CNT = 0
        cnn.Open()
        RDATA = New System.Data.SqlClient.SqlCommand("SELECT FDT FROM TLOG GROUP BY FDT ORDER BY FDT DESC", cnn).ExecuteReader
        Do While RDATA.Read
            CNT += 1
            If CNT = 2 Then DRZ = CStr(RDATA!FDT) : Exit Do
        Loop
        RDATA.Close()
        cnn.Close()
        ' DRZ = "1395/04/21"
    End Function

    Public Sub DSACC(ByVal IP As String, ByVal ACC As Boolean)
        Dim cnn As New System.Data.SqlClient.SqlConnection : cnn.ConnectionString = CNNSTR
        Dim CMD As String
        Dim AC As Integer
        If ACC = True Then AC = 1 Else AC = 0
        CMD = "UPDATE DSLAMS SET ACC = " & AC & " WHERE (IP = '" & Trim(IP) & "')"
        cnn.Open()
        Dim RDATA As System.Data.SqlClient.SqlCommand
        RDATA = New System.Data.SqlClient.SqlCommand(CMD, cnn)
        AC = RDATA.ExecuteNonQuery()
        'MsgBox(AC)
        cnn.Close()

    End Sub
    '****************************

    Public Function INSCERT(ByVal CRT As String) As Boolean
        Dim CNN As New System.Data.SqlClient.SqlConnection : CNN.ConnectionString = CNNSTR
        CNN.Open()
        Dim deltree As System.Data.SqlClient.SqlCommand
        deltree = New System.Data.SqlClient.SqlCommand("UPDATE CONFIG(CODE) VALUES ( '" & CRT & "')", CNN)
        If deltree.ExecuteNonQuery() = 1 Then
            INSCERT = True
        Else
            INSCERT = False
        End If
    End Function

End Module
