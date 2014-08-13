Imports System.IO
Imports System.Net.Mail
Imports System.Text


Public Class Main_Screen

    Dim progresslabel As String = ""

    Dim shownminimizetip As Boolean = False


    Dim busyworking As Boolean = False
    Dim TempFolder As String = ""
    Const sweeptimerMAX As Long = 900
    Dim sweeptimer As Long = sweeptimerMAX
    Dim inputfoldername As String = ""

    Private mailserver1 As String = ""
    Private mailserver1port As String = ""
    Private mailserver2 As String = ""
    Private mailserver2port As String = ""
    Private webmasteraddress As String = ""
    Private webmasterdisplay As String = ""
    Private webroot As String = ""
    Private webroottranslate As String = ""

    Private AutoUpdate As Boolean = False

    Private LastReport As Date



    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
                Label2.Text = "Error encountered in last action"
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Handler(ByVal message As String)
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.WriteLine("")
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Activity Handler")
        End Try
    End Sub


    Private Sub RunWorker(ByVal CompleteFileName As String)
        Try
            If busyworking = False Then
                busyworking = True
                Control_Enabler(False)
                Label2.Text = "Preparing to deal with submission..."
                progresslabel = ""
                BackgroundWorker1.RunWorkerAsync(CompleteFileName)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LastReport = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0, 0, DateTimeKind.Local)
            shownminimizetip = False
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            WebBrowser1.Navigate((Application.StartupPath & "\Images\Monitoring-Still.htm").Replace("\\", "\"))
            TempFolder = (Application.StartupPath & "\Temp").Replace("\\", "\")
            If My.Computer.FileSystem.DirectoryExists(TempFolder) = False Then
                My.Computer.FileSystem.CreateDirectory(TempFolder)
            End If
            loadSettings()
            AboutToolStripMenuItem1.Text = "About " & My.Application.Info.ProductName
            Label2.Text = "Application loaded"
            Sweep()
            SendNotificationEmail("Startup")
        Catch ex As Exception
            Error_Handler(ex, "Form Load")
        End Try
    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Label2.Text = "Submission Detected"
            Dim finfo As FileInfo = New FileInfo(e.Argument.ToString)
            Dim basefilename As String = ""
            basefilename = finfo.Name.Replace("__Complete__", "")
            Dim basefilepath As String = ""
            basefilepath = finfo.DirectoryName
            Dim fileextension As String
            fileextension = finfo.Extension
            fileextension = fileextension.Replace("__Complete__", "")
            finfo = Nothing
            Dim prefix As String = Format(Now, "yyyyMMddHHmmss")
            My.Computer.FileSystem.MoveFile((basefilepath & "\" & basefilename).Replace("\\", "\"), (TempFolder & "\" & prefix & "_" & basefilename).Replace("\\", "\"))
            My.Computer.FileSystem.MoveFile((basefilepath & "\" & basefilename & "__Email__").Replace("\\", "\"), (TempFolder & "\" & prefix & "_" & basefilename & "__Email__").Replace("\\", "\"))
            My.Computer.FileSystem.DeleteFile(e.Argument.ToString)
            Dim emailaddress As String = ""
            emailaddress = My.Computer.FileSystem.ReadAllText((TempFolder & "\" & prefix & "_" & basefilename & "__Email__").Replace("\\", "\"))
            My.Computer.FileSystem.DeleteFile((TempFolder & "\" & prefix & "_" & basefilename & "__Email__").Replace("\\", "\"))
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\PDF Generator DOC2PDF\PDF Generator DOC2PDF.exe").Replace("\\", "\")) = True Then
                Dim procInfo As ProcessStartInfo = New ProcessStartInfo
                If My.Computer.FileSystem.DirectoryExists(outputfolder.Text) = False Then
                    My.Computer.FileSystem.CreateDirectory(outputfolder.Text)
                End If
                procInfo.Arguments = """" & (TempFolder & "\" & prefix & "_" & basefilename).Replace("\\", "\") & """ """ & (outputfolder.Text & "\" & prefix & "_" & basefilename.Replace(fileextension, ".pdf")).Replace("\\", "\") & """ " & emailaddress
                Label2.Text = "Launching Converter"
                If Not fileextension.ToLower = ".exe" Then
                    Select Case fileextension.ToLower
                        Case ".doc"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator DOC2PDF\PDF Generator DOC2PDF.exe").Replace("\\", "\")
                        Case ".docx"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator DOC2PDF\PDF Generator DOC2PDF.exe").Replace("\\", "\")
                        Case ".xls"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator XLS2PDF\PDF Generator XLS2PDF.exe").Replace("\\", "\")
                        Case ".xlsx"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator XLS2PDF\PDF Generator XLS2PDF.exe").Replace("\\", "\")
                        Case ".ppt"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator PPT2PDF\PDF Generator PPT2PDF.exe").Replace("\\", "\")
                        Case ".pptx"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator PPT2PDF\PDF Generator PPT2PDF.exe").Replace("\\", "\")
                        Case ".jpg"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator IMG2PDF\PDF Generator IMG2PDF.exe").Replace("\\", "\")
                        Case ".jpeg"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator IMG2PDF\PDF Generator IMG2PDF.exe").Replace("\\", "\")
                        Case ".bmp"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator IMG2PDF\PDF Generator IMG2PDF.exe").Replace("\\", "\")
                        Case ".gif"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator IMG2PDF\PDF Generator IMG2PDF.exe").Replace("\\", "\")
                        Case ".png"
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator IMG2PDF\PDF Generator IMG2PDF.exe").Replace("\\", "\")
                        Case Else
                            procInfo.FileName = (Application.StartupPath & "\PDF Generator DOC2PDF\PDF Generator DOC2PDF.exe").Replace("\\", "\")
                    End Select

                    Process.Start(procInfo)
                    Activity_Handler("The following job was submitted for conversion:" & vbCrLf & """" & (TempFolder & "\" & prefix & "_" & basefilename).Replace("\\", "\") & """" & vbCrLf & """" & (outputfolder.Text & "\" & prefix & "_" & basefilename.Replace(fileextension, ".pdf")).Replace("\\", "\") & """" & vbCrLf & "" & Uri.EscapeUriString(((outputfolder.Text & "\" & prefix & "_" & basefilename.Replace(fileextension, ".pdf")).Replace("\\", "\")).Replace(webroot, webroottranslate).Replace("\", "/")) & vbCrLf & "" & emailaddress)
                    Label11.Text = Integer.Parse(Label11.Text) + 1
                    Label7.Text = Integer.Parse(Label7.Text) + 1
                    Label9.Text = Integer.Parse(Label9.Text) + 1
                Else
                    SendErrorMail(emailaddress, basefilename)
                    Activity_Handler("The following job was ignored:" & vbCrLf & """" & (TempFolder & "\" & prefix & "_" & basefilename).Replace("\\", "\") & """" & vbCrLf & """" & (outputfolder.Text & "\" & prefix & "_" & basefilename.Replace(fileextension, ".pdf")).Replace("\\", "\") & """" & vbCrLf & "" & emailaddress)
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Background Worker: " & progresslabel)
            progresslabel = "Operation Failed: Error reported (" & progresslabel & ")"
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            Label2.Text = progresslabel
        Catch ex As Exception
            Error_Handler(ex, "Worker Progress Changed")
        End Try
    End Sub

    Private Sub Control_Enabler(ByVal IsEnabled As Boolean)
        Try
            Select Case IsEnabled
                Case True
                    inputfolder.Enabled = True
                    Button1.Enabled = True
                    outputfolder.Enabled = True
                    Button2.Enabled = True
                    MenuStrip1.Enabled = True
                    Me.ControlBox = True
                Case False
                    inputfolder.Enabled = False
                    Button1.Enabled = False
                    outputfolder.Enabled = False
                    Button2.Enabled = False
                    MenuStrip1.Enabled = False
                    Me.ControlBox = False
            End Select
        Catch ex As Exception
            Error_Handler(ex, "Control Enabler")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
        
            Control_Enabler(True)
            If e.Cancelled = True Then
                Label2.Text = "Unable to convert submitted file to PDF"
            Else
                Label2.Text = "File successfully converted to PDF"
            End If
        Catch ex As Exception
            Error_Handler(ex, "Run Worker Completed")
        End Try
        busyworking = False
    End Sub

    

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Label2.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub



    Private Sub loadSettings()
        Try
            Label2.Text = "Loading application settings..."
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then
                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)
                        If lineread.StartsWith("inputfolder=") Then
                            If My.Computer.FileSystem.DirectoryExists(variablevalue) = True Then
                                inputfolder.Text = variablevalue
                                FolderBrowserDialog1.SelectedPath = variablevalue
                            End If
                        End If
                        If lineread.StartsWith("outputfolder=") Then
                            If My.Computer.FileSystem.DirectoryExists(variablevalue) = True Then
                                outputfolder.Text = variablevalue
                                FolderBrowserDialog1.SelectedPath = variablevalue
                            End If
                        End If
                        If lineread.StartsWith("totalsubmissions=") Then
                            Label9.Text = variablevalue
                        End If

                        If lineread.StartsWith("mailserver1=") Then
                            mailserver1 = variablevalue
                            If mailserver1.Length < 1 Then
                                mailserver1 = "mail.uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("mailserver1port=") Then
                            mailserver1port = variablevalue
                            If mailserver1port.Length < 1 Then
                                mailserver1port = "25"
                            End If
                        End If
                        If lineread.StartsWith("mailserver2=") Then
                            mailserver2 = variablevalue
                            If mailserver2.Length < 1 Then
                                mailserver2 = "obe1.com.uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("mailserver2port=") Then
                            mailserver2port = variablevalue
                            If mailserver2port.Length < 1 Then
                                mailserver2port = "25"
                            End If
                        End If
                        If lineread.StartsWith("webmasteraddress=") Then
                            webmasteraddress = variablevalue
                            If webmasteraddress.Length < 1 Then
                                webmasteraddress = "com-webmaster@uct.ac.za"
                            End If
                        End If
                        If lineread.StartsWith("webmasterdisplay=") Then
                            webmasterdisplay = variablevalue
                            If webmasterdisplay.Length < 1 Then
                                webmasterdisplay = "Commerce Webmaster"
                            End If
                        End If
                        If lineread.StartsWith("webroot=") Then
                            webroot = variablevalue
                            If webroot.Length < 1 Then
                                webroot = "C:\Inetpub\wwwroot"
                            End If
                        End If
                        If lineread.StartsWith("webroottranslate=") Then
                            webroottranslate = variablevalue
                            If webroottranslate.Length < 1 Then
                                webroottranslate = "http://www.commerce.uct.ac.za"
                            End If
                        End If
                    End If
                End While
                reader.Close()
                reader = Nothing
            End If
            'default values
            If inputfolder.Text.Length < 1 Then
                inputfolder.Text = "C:\inetpub\wwwroot\services\PDF Generator\Input"
            End If

            If outputfolder.Text.Length < 1 Then
                outputfolder.Text = "C:\inetpub\wwwroot\services\PDF Generator\Output"
            End If
            Label2.Text = "Application Settings successfully loaded"
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub


    Private Sub SaveSettings()
        Try
            Label2.Text = "Saving application settings..."
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            If inputfolder.Text.Length > 0 Then
                writer.WriteLine("inputfolder=" & inputfolder.Text)
            End If
            If outputfolder.Text.Length > 0 Then
                writer.WriteLine("outputfolder=" & outputfolder.Text)
            End If

            writer.WriteLine("totalsubmissions=" & Label9.Text)
            writer.WriteLine("mailserver1=" & mailserver1)
            writer.WriteLine("mailserver1port=" & mailserver1port)
            writer.WriteLine("mailserver2=" & mailserver2)
            writer.WriteLine("mailserver2port=" & mailserver2port)
            writer.WriteLine("webmasteraddress=" & webmasteraddress)
            writer.WriteLine("webmasterdisplay=" & webmasterdisplay)
            writer.WriteLine("webroot=" & webroot)
            writer.WriteLine("webroottranslate=" & webroottranslate)
            writer.Flush()
            writer.Close()
            writer = Nothing

            Label2.Text = "Application Settings successfully saved"

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub


    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            SendNotificationEmail("Shutdown")
            SaveSettings()
            If AutoUpdate = True Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")) = True Then
                    Dim startinfo As ProcessStartInfo = New ProcessStartInfo
                    startinfo.FileName = (Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")
                    startinfo.Arguments = "force"
                    startinfo.CreateNoWindow = False
                    Process.Start(startinfo)
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            sweeptimer = sweeptimer - 1
            Dim tm As TimeSpan = New TimeSpan(0, 0, sweeptimer)
            Label8.Text = tm.ToString
            If sweeptimer = 0 Then
                If busyworking = False Then
                    Sweep()
                End If
                sweeptimer = sweeptimerMAX
            End If
            tm = New TimeSpan(0, 0, sweeptimer)
            Label8.Text = tm.ToString
            tm = Nothing
            If busyworking = False Then
                If FileSystemWatcher1.EnableRaisingEvents = True Then
                    Label2.Text = "Monitoring Input Folder """ & inputfoldername & """"
                    If WebBrowser1.Url.ToString.EndsWith("Still.htm") = True Then
                        WebBrowser1.Navigate((Application.StartupPath & "\Images\Monitoring-Animation.htm").Replace("\\", "\"))
                    End If
                Else
                    Label2.Text = "No folder is currently being monitored"
                    If WebBrowser1.Url.ToString.EndsWith("Animation.htm") = True Then
                        WebBrowser1.Navigate((Application.StartupPath & "\Images\Monitoring-Still.htm").Replace("\\", "\"))
                    End If
                End If

            End If
            Dim dt As Date = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0, 0, DateTimeKind.Local)
            If dt > LastReport Then
                Send_Report(Now, Format(LastReport, "yyyyMMdd"))
                LastReport = dt
            End If

        Catch ex As Exception
            Error_Handler(ex, "Timer Ticking")
        End Try
    End Sub

    Private Sub SendNotificationEmail(ByVal StartOrClose As String)
        Try
            Dim obj As SmtpClient
            If mailserver1port.Length > 0 Then
                obj = New SmtpClient(mailserver1, mailserver1port)
            Else
                obj = New SmtpClient(mailserver1)
            End If

            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

            If StartOrClose = "Startup" Then
                msg.Subject = My.Application.Info.ProductName & ": Application Startup"
                Label2.Text = "Sending Startup Notification"
            Else
                msg.Subject = My.Application.Info.ProductName & ": Application Shutdown"
                Label2.Text = "Sending Shutdown Notification"
            End If

            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
            msg.From = fromaddress
            msg.ReplyTo = fromaddress
            msg.To.Add(fromaddress)

            msg.IsBodyHtml = False

            Dim body As String
            If StartOrClose = "Startup" Then
                body = "This is just a notification message to inform you that " & My.Application.Info.ProductName & " has been successfully started up."
            Else
                body = "This is just a notification message to inform you that " & My.Application.Info.ProductName & " has been shutdown."
            End If

            body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf & "This is an auto-generated email submitted from " & My.Application.Info.ProductName & " at " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & ", running on:"
            body = body & vbCrLf & vbCrLf & "Machine Name: " + Environment.MachineName
            body = body & vbCrLf & "OS Version: " & Environment.OSVersion.ToString()
            body = body & vbCrLf & "User Name: " + Environment.UserName
            msg.Body = body

            obj.DeliveryMethod = SmtpDeliveryMethod.Network
            obj.EnableSsl = False
            obj.UseDefaultCredentials = True

           
            obj.Send(msg)
            obj = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Send Startup/Shutdown Email")
        End Try
    End Sub

    Private Sub Send_Report(ByVal dt As Date, ByVal FileNamePrefix As String)
        '*********************
        'Send Mail Out - activity report
        Try
            Dim obj As SmtpClient
            If mailserver1port.Length > 0 Then
                obj = New SmtpClient(mailserver1, mailserver1port)
            Else
                obj = New SmtpClient(mailserver1)
            End If

            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

            msg.Subject = "PDF Generator Daily Report"
            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
            msg.From = fromaddress
            msg.ReplyTo = fromaddress
            msg.To.Add(fromaddress)

            msg.IsBodyHtml = False

            obj.DeliveryMethod = SmtpDeliveryMethod.Network
            obj.EnableSsl = False
            obj.UseDefaultCredentials = True

            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt") = True Then
                msg.Body = "This daily activity report from PDF Generator running on " & webroottranslate & " was generated at " & Format(dt, "dd/MM/yyyy HH:mm:ss") & "." & vbCrLf & vbCrLf & "The activity report for today is attached. Currently the total number of files generated stands at " & Label9.Text & ", while the number of files generated in this session startup is " & Label7.Text & ". The number of files submitted since the last report was sent is " & Label11.Text
                Dim att As Attachment = New Attachment((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt")
                msg.Attachments.Add(att)
            Else
                msg.Body = "This daily activity report from PDF Generator running on " & webroottranslate & " was generated at " & Format(dt, "dd/MM/yyyy HH:mm:ss") & "." & vbCrLf & vbCrLf & "It seems that no file conversions took place today. Currently the total number of files generated stands at " & Label9.Text & ", while the number of files generated in this session startup is " & Label7.Text & ". The number of files submitted since the last report was sent is " & Label11.Text
            End If
            Label11.Text = 0
            obj.Send(msg)
            obj = Nothing
        Catch ex As Exception
            Try
                Dim obj As SmtpClient
                If mailserver2port.Length > 0 Then
                    obj = New SmtpClient(mailserver2, mailserver2port)
                Else
                    obj = New SmtpClient(mailserver2)
                End If
                Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                msg.Subject = "PDF Generator Daily Report"
                Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                msg.From = fromaddress
                msg.ReplyTo = fromaddress
                msg.To.Add(fromaddress)

                msg.IsBodyHtml = False

                obj.DeliveryMethod = SmtpDeliveryMethod.Network
                obj.EnableSsl = False
                obj.UseDefaultCredentials = True

                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt") = True Then
                    msg.Body = "This daily activity report from PDF Generator running on " & webroottranslate & " was generated at " & Format(dt, "dd/MM/yyyy HH:mm:ss") & "." & vbCrLf & vbCrLf & "The activity report for today is attached. Currently the total number of files generated stands at " & Label9.Text & ", while the number of files generated in this session is " & Label7.Text & ". The number of files submitted since the last report was sent is " & Label11.Text
                    Dim att As Attachment = New Attachment((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & FileNamePrefix & "_Activity_Log.txt")
                    msg.Attachments.Add(att)
                Else
                    msg.Body = "This daily activity report from PDF Generator running on " & webroottranslate & " was generated at " & Format(dt, "dd/MM/yyyy HH:mm:ss") & "." & vbCrLf & vbCrLf & "It seems that no file conversions took place today. Currently the total number of files generated stands at " & Label9.Text & ", while the number of files generated in this session is " & Label7.Text & ". The number of files submitted since the last report was sent is " & Label11.Text
                End If

                obj.Send(msg)
                obj = Nothing
            Catch ex1 As Exception
                Error_Handler(ex, "Send Report")
            End Try
        End Try
    End Sub



    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub

    Private Sub Main_Screen_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
                If shownminimizetip = False Then
                    NotifyIcon1.ShowBalloonTip(1)
                    shownminimizetip = True
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Change Window State")
        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If My.Computer.FileSystem.DirectoryExists(inputfolder.Text) = True Then
                FolderBrowserDialog1.SelectedPath = inputfolder.Text
            End If
            FolderBrowserDialog1.Description = "Select the input folder which needs to be monitored."
            If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                inputfolder.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex, "Select Input Folder")
        End Try
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If My.Computer.FileSystem.DirectoryExists(outputfolder.Text) = True Then
                FolderBrowserDialog1.SelectedPath = outputfolder.Text
            End If
            FolderBrowserDialog1.Description = "Select the output folder in which to store the generated PDF files."
            If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                outputfolder.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex, "Select Output Folder")
        End Try
    End Sub

    Private Sub inputfolder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles inputfolder.TextChanged
        Try
            If My.Computer.FileSystem.DirectoryExists(inputfolder.Text) = True Then
                FileSystemWatcher1.Path = inputfolder.Text
                Dim dinfo As DirectoryInfo = New DirectoryInfo(inputfolder.Text)
                inputfoldername = dinfo.Name
                dinfo = Nothing
                FileSystemWatcher1.EnableRaisingEvents = True
            Else
                FileSystemWatcher1.EnableRaisingEvents = False
            End If
        Catch ex As Exception
            Error_Handler(ex, "Input Folder Changed")
        End Try
    End Sub

    Private Sub FileSystemWatcher1_Created(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Created
        Try
            If e.FullPath.EndsWith("__Complete__") Then
                RunWorker(e.FullPath)
            End If
        Catch ex As Exception
            Error_Handler(ex, "File Creation Detected")
        End Try
    End Sub

    Private Sub FileSystemWatcher1_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Changed
        Try
            If e.FullPath.EndsWith("__Complete__") Then
                RunWorker(e.FullPath)
            End If
        Catch ex As Exception
            Error_Handler(ex, "File Change Detected")
        End Try
    End Sub

    Private Sub FileSystemWatcher1_Renamed(ByVal sender As System.Object, ByVal e As System.IO.RenamedEventArgs) Handles FileSystemWatcher1.Renamed
        Try
            If e.FullPath.EndsWith("__Complete__") Then
                RunWorker(e.FullPath)
            End If
        Catch ex As Exception
            Error_Handler(ex, "File Rename Detected")
        End Try
    End Sub
    Private Sub Sweep()
        Try
            Dim dinfo As DirectoryInfo
            dinfo = New DirectoryInfo(TempFolder)
            If dinfo.Exists = True Then
                For Each finfo As FileInfo In dinfo.GetFiles
                    Try
                        If Now > finfo.CreationTime.AddHours(48) Then
                            finfo.Delete()
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Sweep")
                    End Try
                    finfo = Nothing
                Next
            End If
            dinfo = Nothing
            dinfo = New DirectoryInfo(outputfolder.Text)
            If dinfo.Exists = True Then
                For Each finfo As FileInfo In dinfo.GetFiles
                    Try
                        If Now > finfo.CreationTime.AddHours(48) Then
                            finfo.Delete()
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Sweep")
                    End Try
                    finfo = Nothing
                Next
            End If
            dinfo = Nothing
            dinfo = New DirectoryInfo(inputfolder.Text)
            If dinfo.Exists = True Then
                For Each finfo As FileInfo In dinfo.GetFiles
                    Try
                        If finfo.FullName.EndsWith("__Complete__") = True And Now <= finfo.CreationTime.AddHours(48) Then
                            If My.Computer.FileSystem.FileExists(finfo.FullName.Replace("__Complete__", "__Email__")) = True And My.Computer.FileSystem.FileExists(finfo.FullName.Replace("__Complete__", "")) = True Then
                                Dim fullpath As String = finfo.FullName
                                RunWorker(fullpath)
                                finfo = Nothing
                            Else
                                finfo.Delete()
                            End If
                        Else
                            If Now > finfo.CreationTime.AddHours(48) Then
                                finfo.Delete()
                            End If
                            finfo = Nothing
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Sweep")
                        finfo = Nothing
                    End Try
                Next
            End If
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Sweep")
        End Try
    End Sub

    Private Sub SendErrorMail(ByVal inputEmailAddress As String, ByVal Filename As String)
        '*********************
        'Send Mail Out - file wasn't converted
        Try
            Dim obj As SmtpClient
            If mailserver1port.Length > 0 Then
                obj = New SmtpClient(mailserver1, mailserver1port)
            Else
                obj = New SmtpClient(mailserver1)
            End If

            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

            msg.Subject = "PDF Generator Notification"
            Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
            msg.From = fromaddress
            msg.ReplyTo = fromaddress
            msg.To.Add(inputEmailAddress)

            msg.IsBodyHtml = False

            obj.DeliveryMethod = SmtpDeliveryMethod.Network
            obj.EnableSsl = False
            obj.UseDefaultCredentials = True
            msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that there was a problem in converting your submitted file (" & Filename & "). It is suggested that you try submitting a simple text file, just to ensure that this service is in fact operating correctly."
            obj.Send(msg)
            obj = Nothing
        Catch ex As Exception
            Try
                Dim obj As SmtpClient
                If mailserver2port.Length > 0 Then
                    obj = New SmtpClient(mailserver2, mailserver2port)
                Else
                    obj = New SmtpClient(mailserver2)
                End If
                Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage

                msg.Subject = "PDF Generator Notification"
                Dim fromaddress As MailAddress = New MailAddress(webmasteraddress, webmasterdisplay)
                msg.From = fromaddress
                msg.ReplyTo = fromaddress
                msg.To.Add(inputEmailAddress)

                msg.IsBodyHtml = False

                obj.DeliveryMethod = SmtpDeliveryMethod.Network
                obj.EnableSsl = False
                obj.UseDefaultCredentials = True
                msg.Body = "This is a notification message from PDF Generator running on " & webroottranslate & " to inform you that there was a problem in converting your submitted file (" & Filename & "). It is suggested that you try submitting a simple text file, just to ensure that this service is in fact operating correctly."

                obj.Send(msg)
                obj = Nothing
            Catch ex1 As Exception
                Error_Handler(ex, "Send Error Mail")
            End Try
        End Try
    End Sub
 
    Private Sub AutoUpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoUpdateToolStripMenuItem.Click
        Try
            AutoUpdate = True
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "AutoUpdate")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        Try
            Label2.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub
End Class
