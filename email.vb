Imports System.Web.Configuration.WebConfigurationManager
Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient
Imports Common.Validation
Imports Common.Enums

Namespace Email

    Public Module mEmail

        ''' <summary>
        ''' Splits a string of email addresses and adds each address to the passed in mail address collection
        ''' </summary>
        ''' <param name="objEmails">The MailAddressCollection to add the email addresses to</param>
        ''' <param name="strEmails">A comma or semi-colon separated list of email addresses</param>
        Public Sub SplitEmailAddresses(ByRef objEmails As MailAddressCollection, ByVal strEmails As String)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Splits a string of email addresses and adds each address
            '               to the passed in mail address collection
            ' *****************************************************************************
            Dim astrEmails As Array
            Dim intCount As Integer

            If strEmails.Contains(";") Then
                astrEmails = strEmails.Split(";")
                intCount = astrEmails.Length - 1
                While intCount >= 0
                    astrEmails(intCount).ToString.Replace(" ", "")
                    If ValidEmail(astrEmails(intCount).ToString, False) Then
                        objEmails.Add(astrEmails(intCount).ToString.Trim)
                    End If
                    intCount -= 1
                End While
            ElseIf strEmails.Contains(",") Then
                astrEmails = strEmails.Split(",")
                intCount = astrEmails.Length - 1
                While intCount >= 0
                    astrEmails(intCount).ToString.Replace(" ", "")
                    If ValidEmail(astrEmails(intCount).ToString, False) Then
                        objEmails.Add(astrEmails(intCount).ToString.Trim)
                    End If
                    intCount -= 1
                End While
            ElseIf ValidEmail(strEmails, False) Then
                objEmails.Add(strEmails)
            End If
        End Sub

        ''' <summary>
        ''' Adds a new email log entry to the database
        ''' </summary>
        ''' <param name="strName">A name used only as a reference to the feature or function that generated the email</param>
        ''' <param name="strStatus">A status string describing the outcome of the email attempt</param>
        ''' <param name="datStart">The time that the system started generating the email</param>
        ''' <param name="datEnd">The time that the system finished generating the email</param>
        ''' <param name="strException">Any error messages that were thrown while generating or sending the email</param>
        Public Sub AddEmailLogEntry(ByVal strName As String, ByVal strStatus As String, _
            ByVal datStart As Date, ByVal datEnd As Date, Optional ByVal strException As String = Nothing)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds a new email log entry to the database
            ' *****************************************************************************
            Dim objConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(My.Settings.ConnName).ToString)
            Dim cmdGet As New SqlCommand("spEmailerLog_AddEntry", objConn)
            Dim parEmailerName As New SqlParameter("@EmailerName", SqlDbType.VarChar, 50)
            Dim parStatus As New SqlParameter("@Status", SqlDbType.VarChar, 50)
            Dim parDuration As New SqlParameter("@Duration", SqlDbType.Float)
            Dim parException As New SqlParameter("@Exception", SqlDbType.VarChar)

            'set up the command object
            cmdGet.Connection = objConn
            cmdGet.CommandType = CommandType.StoredProcedure
            cmdGet.CommandTimeout = 30

            'add the parameters to the command object
            cmdGet.Parameters.Add(parEmailerName)
            cmdGet.Parameters.Add(parStatus)
            cmdGet.Parameters.Add(parDuration)
            cmdGet.Parameters.Add(parException)

            'assign values to the parameters
            parEmailerName.Value = strName
            parStatus.Value = strStatus
            parDuration.Value = TimeSpan.FromTicks(datEnd.Ticks - datStart.Ticks).TotalSeconds
            parException.Value = IIf(strException = "", DBNull.Value, strException)

            'open the connection
            objConn.Open()

            cmdGet.ExecuteNonQuery()

            objConn.Close()
        End Sub

        ''' <summary>
        ''' Sends an email through the SMTP server specified in the web.config
        ''' </summary>
        ''' <param name="strSendTo">A comma or semi-colon separated list of email addresses to send the email to</param>
        ''' <param name="strSendFrom">The email address that the email should be sent from</param>
        ''' <param name="strSubject">The subject of the email</param>
        ''' <param name="strBody">The body of the email</param>
        ''' <param name="blnHTML">Boolean value indicating if the body contains HTML</param>
        ''' <param name="strFromDisplay">The name that should be displayed instead of the from address</param>
        ''' <param name="strCC">A comma or semi-colon separated list of email addresses to CC the email to</param>
        ''' <param name="strBCC">A comma or semi-colon separated list of email addresses to BCC the email to</param>
        ''' <param name="strReplyTo">A comma or semi-colon separated list of email addresses to include as the reply to addresses</param>
        ''' <returns>A return status object indicating if the email was successful or not</returns>
        Public Function SendMail(ByVal strSendTo As String, ByVal strSendFrom As String, _
            ByVal strSubject As String, ByVal strBody As String, _
            Optional ByVal blnHTML As Boolean = True, Optional ByVal strFromDisplay As String = "", _
            Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "", _
            Optional ByVal strReplyTo As String = "") As RetStatus
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Sends an email through the SMTP server specified in the web.config
            ' Modified:     2012.10.30 by JR - Added reply to address parameter
            ' *****************************************************************************
            SendMail = RetStatus.Fail

            Dim mmMessage As New MailMessage
            Dim objSmtpClient As New SmtpClient
            Dim maFromAddr As New MailAddress(strSendFrom, strFromDisplay)

            If strSendTo.Trim = "" And strBCC.Trim = "" Then
                mmMessage.Dispose()
                Exit Function
            End If

            'Add email addresses
            SplitEmailAddresses(mmMessage.To, strSendTo)
            SplitEmailAddresses(mmMessage.CC, strCC)
            SplitEmailAddresses(mmMessage.Bcc, strBCC)
            SplitEmailAddresses(mmMessage.ReplyToList, strReplyTo)

            'Add from address, subject adn body
            mmMessage.From = maFromAddr
            mmMessage.Subject = strSubject
            mmMessage.Body = strBody

            'Send Message
            mmMessage.IsBodyHtml = blnHTML

            'Set up the object with values from the web config
            objSmtpClient.Port = CInt(AppSettings.Get("SmtpPort"))
            objSmtpClient.Host = AppSettings.Get("SmtpHost")
            objSmtpClient.Credentials = New System.Net.NetworkCredential(AppSettings.Get("SmtpUser"), AppSettings.Get("SmtpPassword"))

            'Set time out send message and cleanup
            objSmtpClient.Timeout = 15000
            objSmtpClient.Send(mmMessage)
            mmMessage.Dispose()

            SendMail = RetStatus.Pass
        End Function

        ''' <summary>
        ''' Starts a new thread to send an email through the SMTP server specified in the web.config 
        ''' </summary>
        ''' <param name="strSendTo">A comma or semi-colon separated list of email addresses to send the email to</param>
        ''' <param name="strSendFrom">The email address that the email should be sent from</param>
        ''' <param name="strSubject">The subject of the email</param>
        ''' <param name="strBody">The body of the email</param>
        ''' <param name="blnHTML">Boolean value indicating if the body contains HTML</param>
        ''' <param name="strFromDisplay">The name that should be displayed instead of the from address</param>
        ''' <param name="strCC">A comma or semi-colon separated list of email addresses to CC the email to</param>
        ''' <param name="strBCC">A comma or semi-colon separated list of email addresses to BCC the email to</param>
        ''' <param name="strReplyTo">A comma or semi-colon separated list of email addresses to include as the reply to addresses</param>
        ''' <returns>A return status object indicating if the new thread was created succesfully or not</returns>
        Public Function SendMail_BG(ByVal strSendTo As String, ByVal strSendFrom As String, _
            ByVal strSubject As String, ByVal strBody As String, _
            Optional ByVal blnHTML As Boolean = True, Optional ByVal strFromDisplay As String = "", _
            Optional ByVal strCC As String = "", Optional ByVal strBCC As String = "", _
            Optional ByVal strReplyTo As String = "") As RetStatus
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Sends an email through the SMTP server specified in the web.config
            ' Modified:     2012.10.30 by JR - Added reply to address parameter
            ' *****************************************************************************
            SendMail_BG = RetStatus.Fail

            Threading.Tasks.Task.Factory.StartNew(Function() SendMail(strSendTo, strSendFrom, strSubject, strBody, blnHTML, strFromDisplay, strCC, strBCC, strReplyTo))

            SendMail_BG = RetStatus.Pass
        End Function

    End Module

End Namespace