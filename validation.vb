Imports System.Text
Imports System.Text.RegularExpressions
Imports Common.Clean

Namespace Validation

    Public Module mValidation

        ''' <summary>
        ''' Checks the given email address to see if it is valid
        ''' </summary>
        ''' <param name="strEmail">The email address string to be validated</param>
        ''' <param name="blnEmptyOk">Boolean value indicating if blank strings are valid</param>
        ''' <param name="blnAllowMultiple">Boolean value indicating if multiple email addresses are allowed</param>
        ''' <returns>A boolean value indicating if the email address is valid or not</returns>
        ''' <remarks>Multiple email addresses can be separated by either semicolons or commas</remarks>
        Public Function ValidEmail(ByVal strEmail As String, Optional ByVal blnEmptyOk As Boolean = True, _
            Optional ByVal blnAllowMultiple As Boolean = False) As Boolean
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks the given email address to see if it is valid
            ' *****************************************************************************
            Dim blnValid As Boolean = False
            Dim astrEmails() As String

            'Does the option allow empty string as a valid email
            If strEmail.Trim() = "" Then
                Return blnEmptyOk
            End If

            'check the email formatting
            If Regex.IsMatch(strEmail.Trim.ToLower, "^[a-z0-9\._%+-]+@(?:[a-z0-9-]+\.)+(?:[a-z]{2}|aero|biz|com|co\.uk|edu|gov|info|jobs|mil|mobi|museum|name|net|org|tv|us|ws)$") Then
                blnValid = True
            End If

            'check for multiple email addresses, separated by semi-colons or commas
            If Not blnValid And strEmail.Contains(";") And blnAllowMultiple Then
                blnValid = True
                astrEmails = strEmail.Split(";")
                For Each strEmail In astrEmails
                    blnValid = blnValid And ValidEmail(strEmail.Trim, blnEmptyOk)
                Next
            ElseIf Not blnValid And strEmail.Contains(",") And blnAllowMultiple Then
                blnValid = True
                astrEmails = strEmail.Split(",")
                For Each strEmail In astrEmails
                    blnValid = blnValid And ValidEmail(strEmail.Trim, blnEmptyOk)
                Next
            End If

            Return blnValid
        End Function

        ''' <summary>
        ''' Checks the given web address to see if it is valid
        ''' </summary>
        ''' <param name="strURL">The url string to be validated</param>
        ''' <param name="blnEmptyOK">Boolean value indicating if blank strings are valid</param>
        ''' <returns>A boolean value indicating if the url is valid or not</returns>
        Public Function ValidURL(ByVal strURL As String, Optional ByVal blnEmptyOK As Boolean = True) As Boolean
            '*********************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks the given web address to see if it is valid
            '*********************************************************************************
            If strURL.Trim = "" And blnEmptyOK Then
                ValidURL = True
            Else
                ValidURL = Regex.IsMatch(strURL.Trim.ToLower, "^(https?://)?(www\.)?[a-z0-9\-]*\.(?:[a-z]{2}|aero|asia|biz|com|coop|co\.uk|edu|gov|info|int|jobs|mil|mobi|museum|name|net|org|pro|tel|travel)(/?.*)$")
            End If
        End Function

        ''' <summary>
        ''' Checks the given phone number to see if it is valid
        ''' </summary>
        ''' <param name="strPhone">The phone number string to be validated</param>
        ''' <param name="blnEmptyOk">Boolean value indicating if blank strings are valid</param>
        ''' <returns>A boolean value indicating if the phone number is valid or not</returns>
        Public Function ValidPhone(ByVal strPhone As String, Optional ByVal blnEmptyOk As Boolean = True) As Boolean
            '*********************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks the given phone number to see if it is valid
            '*********************************************************************************
            Dim intLength As Integer = NumbersOnly(strPhone.Trim).Length

            If (intLength = 0 And blnEmptyOk) Or intLength = 7 Or _
                intLength = 10 Or intLength = 11 Then
                Return True
            Else : Return False
            End If
        End Function

        ''' <summary>
        ''' Checks the given zip code to see if it is valid
        ''' </summary>
        ''' <param name="strZip">The zip code string to be validated</param>
        ''' <param name="blnEmptyOk">Boolean value indicating if blank strings are valid</param>
        ''' <returns>A boolean value indicating if the zip code is valid or not</returns>
        Public Function ValidZipCode(ByVal strZip As String, Optional ByVal blnEmptyOk As Boolean = True) As Boolean
            '*********************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks the given zip code to see if it is valid
            '*********************************************************************************
            Dim intLength As Integer = NumbersOnly(strZip.Trim).Length

            If (intLength = 0 And blnEmptyOk) Or _
                intLength = 5 Or intLength = 9 Then
                Return True
            Else : Return False
            End If

        End Function

        ''' <summary>
        ''' Checks the given login name to see if it is valid
        ''' </summary>
        ''' <param name="strLogin">The login string to be validated</param>
        ''' <returns>A boolean value indicating if the login name is valid or not</returns>
        Public Function ValidLogin(ByVal strLogin As String) As Boolean
            '*********************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks the given login name to see if it is valid
            '*********************************************************************************
            Dim intLen As Integer
            Dim intCtr As Integer
            Dim strChar As String

            'returns true if all characters in a string are alphanumeric
            'returns false otherwise or for empty string
            ValidLogin = False

            intLen = Len(strLogin)
            If intLen > 0 Then
                For intCtr = 1 To intLen
                    strChar = Mid(strLogin, intCtr, 1)
                    If Not strChar Like "[0-9A-Za-z]" Then Exit Function
                Next

                ValidLogin = True
            End If

        End Function

        ''' <summary>
        ''' Checks the http response code for the given url to see if the remote file is available
        ''' </summary>
        ''' <param name="strURL">The URL of the file to check</param>
        ''' <returns>A boolean value indicating if the remote file exists</returns>
        Public Function RemoteFileOk(ByVal strURL As String) As Boolean
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Created Date: 2012.02.10
            ' Description:  Checks the http response code for the given url to see if the remote file is available
            ' *****************************************************************************
            Dim objRequest As System.Net.HttpWebRequest = TryCast(System.Net.WebRequest.Create(strURL), System.Net.HttpWebRequest)

            objRequest.Method = "HEAD"
            Try
                Using objResponse As System.Net.HttpWebResponse = TryCast(objRequest.GetResponse(), System.Net.HttpWebResponse)
                    Return (objResponse.StatusCode = System.Net.HttpStatusCode.OK)
                End Using
            Catch ex As System.Net.WebException
                Return False
            End Try
        End Function

    End Module

End Namespace