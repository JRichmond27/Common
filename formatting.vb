Imports Common.Clean
Imports Common.Enums

Namespace Formatting

    Public Module mFormatting

#Region " Numbers "

        ''' <summary>
        ''' Takes a numeric string value and returns it with commas in the proper locations
        ''' </summary>
        ''' <param name="strNumber">Numeric string value to be formatted</param>
        ''' <returns>Numeric string value with commas inserted in the proper locations</returns>
        Public Function FormatNumber(ByVal strNumber As String) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Takes a numeric string value and returns it with commas in the 
            '               proper locations
            ' *****************************************************************************
            Dim intPos As Integer
            Dim intLen As Integer
            Dim strResult As String = strNumber

            Try
                'get the length of the numeric string
                intLen = strNumber.Length

                'get the position of the decimal if there is one
                intPos = strNumber.IndexOf(".")

                'set the starting point depending on if there is a decimal
                If intPos < 0 Then
                    intPos = intLen - 3
                Else
                    intPos = intPos - 3
                End If

                'loop through the string until commas have been inserted every 3 digits
                While intPos > 0
                    strResult = strNumber.Substring(0, intPos) & "," & strResult.Substring(intPos, strResult.Length - intPos)
                    intPos = intPos - 3
                End While

                'return the formatted numeric string
                Return strResult

            Catch ex As Exception
                Throw ex
                Return strNumber
            End Try
        End Function

        ''' <summary>
        ''' Converts an integer into a roman numeral string
        ''' </summary>
        ''' <param name="intNumber">The number to be converted to roman numerals</param>
        ''' <returns>A string of roman numerals</returns>
        Public Function FormatRomanNumerals(ByVal intNumber As Integer) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts an integer into a roman numeral string
            ' *****************************************************************************
            Dim i As Integer
            Dim intValues As Integer() = {1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
            Dim strNumerals As String() = {"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
            Dim objResult As New Text.StringBuilder()

            If intNumber < 0 Or intNumber > 3999 Then
                Throw New ArgumentException("Value must be in the range 0 - 3,999.")
            End If

            If intNumber = 0 Then
                Return "N"
            End If

            'loop through each of the values to diminish the number
            For i = 0 To 12
                'If the number being converted is less than the test value, append
                'the corresponding numeral or numeral pair to the resultant string
                While intNumber >= intValues(i)
                    intNumber -= intValues(i)
                    objResult.Append(strNumerals(i))
                End While
            Next i

            Return objResult.ToString()
        End Function

        ''' <summary>
        ''' Takes a numeric string value and returns it formatted as a monetary value 
        ''' </summary>
        ''' <param name="strMoney">The numeric value to be formatted</param>
        ''' <param name="intDecimalDigits">The number of decimal places to include</param>
        ''' <param name="blnLeadingDigit">Boolean value indicating if a leading zero should be used for values between 1 and -1</param>
        ''' <param name="blnGrouping">Boolean value indicating if digit grouping should be performed</param>
        ''' <param name="blnUseNegParen">Boolean value indicating if negative values should be displayed with parenthesis</param>
        ''' <param name="blnReturnZero">Boolean value indicating if a zero value should be returned for empty strings</param>
        ''' <returns>The numeric string value, formatted a monetary value</returns>
        Public Function FormatMoney(ByVal strMoney As String, Optional ByVal intDecimalDigits As Integer = 2, _
            Optional ByVal blnLeadingDigit As Boolean = True, Optional ByVal blnGrouping As Boolean = True, _
            Optional ByVal blnUseNegParen As Boolean = False, Optional ByVal blnReturnZero As Boolean = True) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Takes a numeric string value and returns it formatted as a monetary value 
            ' *****************************************************************************
            Dim strResult As String
            Dim blnNegative As Boolean = False

            'clean up the value
            strResult = NumbersOnly(strMoney & "", True, True)

            'just return if there is no numeric value
            If strResult.Trim = "" Then
                If blnReturnZero Then
                    Return "$0.00"
                Else
                    Return ""
                End If
            End If

            'return the value
            Return Strings.FormatCurrency(strMoney, intDecimalDigits, blnLeadingDigit, blnUseNegParen, blnGrouping)
        End Function

        ''' <summary>
        ''' Converts and formats a file size into a readable format
        ''' </summary>
        ''' <param name="dblSize">The file size to be formatted</param>
        ''' <param name="intUnits">The original unit type of the file size</param>
        ''' <param name="intDigits">The number of digits to round the result to</param>
        ''' <returns>A formatted file size string</returns>
        Public Function FormatFileSize(ByVal dblSize As Double, Optional ByVal intUnits As FileSizeUnits = FileSizeUnits.KB, Optional ByVal intDigits As Integer = 1) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts and formats a file size into a readable format
            ' *****************************************************************************
            Select Case intUnits
                Case FileSizeUnits.B
                    If dblSize > 999999999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.B, FileSizeUnits.TB, intDigits)) & " TB"
                    ElseIf dblSize > 999999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.B, FileSizeUnits.GB, intDigits)) & " GB"
                    ElseIf dblSize > 999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.B, FileSizeUnits.MB, intDigits)) & " MB"
                    ElseIf dblSize > 999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.B, FileSizeUnits.KB, intDigits)) & " KB"
                    Else
                        FormatFileSize = FormatNumber(dblSize) & " B"
                    End If
                Case FileSizeUnits.KB
                    If dblSize > 999999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.KB, FileSizeUnits.TB, intDigits)) & " TB"
                    ElseIf dblSize > 999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.KB, FileSizeUnits.GB, intDigits)) & " GB"
                    ElseIf dblSize > 999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.KB, FileSizeUnits.MB, intDigits)) & " MB"
                    Else
                        FormatFileSize = FormatNumber(dblSize) & " KB"
                    End If
                Case FileSizeUnits.MB
                    If dblSize > 999999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.MB, FileSizeUnits.TB, intDigits)) & " TB"
                    ElseIf dblSize > 999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.MB, FileSizeUnits.GB, intDigits)) & " GB"
                    Else
                        FormatFileSize = FormatNumber(dblSize) & " MB"
                    End If
                Case FileSizeUnits.GB
                    If dblSize > 999 Then
                        FormatFileSize = FormatNumber(ConvertFileSize(dblSize, FileSizeUnits.GB, FileSizeUnits.TB, intDigits)) & " TB"
                    Else
                        FormatFileSize = FormatNumber(dblSize) & " GB"
                    End If
                Case FileSizeUnits.TB
                    FormatFileSize = FormatNumber(dblSize) & " TB"
                Case Else
                    FormatFileSize = FormatNumber(dblSize)
            End Select
        End Function

        ''' <summary>
        ''' Converts a file size from one unit type to another
        ''' </summary>
        ''' <param name="intSize">The file size to be converted</param>
        ''' <param name="intFrom">The original unit type of the file size</param>
        ''' <param name="intTo">The new unit type to convert the file size to</param>
        ''' <param name="intDigits">The number of digits to round the result to</param>
        ''' <returns>The original file size, converted to the new unit type</returns>
        Public Function ConvertFileSize(ByVal intSize As Long, ByVal intFrom As FileSizeUnits, ByVal intTo As FileSizeUnits, Optional ByVal intDigits As Integer = 1) As Double
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts a file size from one unit type to another
            ' *****************************************************************************
            ConvertFileSize = intSize

            Select Case intFrom
                Case FileSizeUnits.B
                    Select Case intTo
                        Case FileSizeUnits.KB
                            ConvertFileSize = Math.Round(intSize / 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.MB
                            ConvertFileSize = Math.Round(intSize / 1048576, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.GB
                            ConvertFileSize = Math.Round(intSize / 1073741824, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.TB
                            ConvertFileSize = Math.Round(intSize / 1099511627776, intDigits, MidpointRounding.AwayFromZero)
                    End Select
                Case FileSizeUnits.KB
                    Select Case intTo
                        Case FileSizeUnits.B
                            ConvertFileSize = Math.Round(intSize * 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.MB
                            ConvertFileSize = Math.Round(intSize / 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.GB
                            ConvertFileSize = Math.Round(intSize / 1048576, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.TB
                            ConvertFileSize = Math.Round(intSize / 1073741824, intDigits, MidpointRounding.AwayFromZero)
                    End Select
                Case FileSizeUnits.MB
                    Select Case intTo
                        Case FileSizeUnits.B
                            ConvertFileSize = Math.Round(intSize * 1048576, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.KB
                            ConvertFileSize = Math.Round(intSize * 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.GB
                            ConvertFileSize = Math.Round(intSize / 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.TB
                            ConvertFileSize = Math.Round(intSize / 1048576, intDigits, MidpointRounding.AwayFromZero)
                    End Select
                Case FileSizeUnits.GB
                    Select Case intTo
                        Case FileSizeUnits.B
                            ConvertFileSize = Math.Round(intSize * 1073741824, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.KB
                            ConvertFileSize = Math.Round(intSize * 1048576, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.MB
                            ConvertFileSize = Math.Round(intSize * 1024, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.TB
                            ConvertFileSize = Math.Round(intSize / 1024, intDigits, MidpointRounding.AwayFromZero)
                    End Select
                Case FileSizeUnits.TB
                    Select Case intTo
                        Case FileSizeUnits.B
                            ConvertFileSize = Math.Round(intSize * 1099511627776, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.KB
                            ConvertFileSize = Math.Round(intSize * 1073741824, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.MB
                            ConvertFileSize = Math.Round(intSize * 1048576, intDigits, MidpointRounding.AwayFromZero)
                        Case FileSizeUnits.GB
                            ConvertFileSize = Math.Round(intSize * 1024, intDigits, MidpointRounding.AwayFromZero)
                    End Select
            End Select
        End Function

        Public Function PadLeft(ByVal strValue As String, ByVal intLength As Integer, Optional ByVal chrPadding As Char = "0") As String
            While strValue.Length < intLength
                strValue = chrPadding & strValue
            End While

            Return strValue
        End Function

#End Region

#Region " Contact Info "

        ''' <summary>
        ''' Formats a numeric string as a phone number
        ''' </summary>
        ''' <param name="strPhone">The numeric string to be formatted</param>
        ''' <param name="blnUseDots">Boolean value indicating if dot formatting should be used</param>
        ''' <returns>The original string formatted as a phone number</returns>
        Public Function FormatPhone(ByVal strPhone As String, Optional ByVal blnUseDots As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Formats a numeric string as a phone number
            ' Modified:     2011.06.23 by JR - Added option for dot formatting (555.555.1234)
            ' *****************************************************************************
            Dim intLen As Integer

            intLen = strPhone.Length
            FormatPhone = strPhone

            If blnUseDots Then
                If intLen >= 7 Then
                    FormatPhone = strPhone.Substring(intLen - 4, 4)
                    FormatPhone = strPhone.Substring(intLen - 7, 3) & "." & FormatPhone
                    If intLen >= 10 Then
                        FormatPhone = strPhone.Substring(intLen - 10, 3) & "." & FormatPhone
                    End If
                End If
            Else
                If intLen >= 7 Then
                    FormatPhone = strPhone.Substring(intLen - 4, 4)
                    FormatPhone = strPhone.Substring(intLen - 7, 3) & "-" & FormatPhone
                    If intLen >= 10 Then
                        FormatPhone = "(" & strPhone.Substring(intLen - 10, 3) & ") " & FormatPhone
                    End If
                End If
            End If
        End Function

        ''' <summary>
        ''' Formats a numeric string as a social security number
        ''' </summary>
        ''' <param name="strSSN">The numeric string to be formatted</param>
        ''' <returns>The original string formatted as a social security number</returns>
        Public Function FormatSSN(ByVal strSSN As String) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Formats a numeric string as a social security number
            ' *****************************************************************************
            Dim intLen As Integer

            intLen = strSSN.Length
            FormatSSN = strSSN

            If intLen = 9 Then
                FormatSSN = strSSN.Substring(0, 3)
                FormatSSN = FormatSSN & "-" & strSSN.Substring(3, 2)
                FormatSSN = FormatSSN & "-" & strSSN.Substring(5, 4)
            End If
        End Function

        ''' <summary>
        ''' Formats a string as a clickable html mailto link
        ''' </summary>
        ''' <param name="strEmail">The email address to be formatted</param>
        ''' <param name="strName">The name of the person to be emailed</param>
        ''' <param name="strSubject">The subject to be used in the email</param>
        ''' <param name="strText">The text to be displayed as a link</param>
        ''' <param name="strTitle">The title attibute to be used in the link</param>
        ''' <returns>An html anchor element string that links to the specified email address</returns>
        Public Function FormatEmail(ByVal strEmail As String, Optional ByVal strName As String = "", _
            Optional ByVal strSubject As String = "", Optional ByVal strText As String = "", _
            Optional ByVal strTitle As String = "") As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Formats a string as a clickable html mailto link
            ' Modified:     2011.06.23 by JR - Added html encoding to help obfuscate email address
            ' *****************************************************************************
            If strEmail.Trim <> "" Then
                'encode the email address
                strEmail = strEmail.Trim.ToLower.Replace("@", "&#064;").Replace(".", "&#46;").Replace("c", "&#99;").Replace("o", "&#111;").Replace("m", "&#109;").Replace("r", "&#114;").Replace("g", "&#103;").Replace("n", "&#110;").Replace("e", "&#101;").Replace("t", "&#116;")

                'build the href
                FormatEmail = "<a href=""&#109;&#097;&#105;&#108;&#116;&#111;&#58;" & strEmail

                'add the subject
                If strSubject.Trim = "" Then
                    FormatEmail &= """"
                Else
                    FormatEmail &= "?subject=" & strSubject.Trim.Replace(" ", "%20") & """"
                End If

                'add the title
                If strTitle.Trim = "" Then
                    If strName.Trim = "" Then
                        FormatEmail &= " title=""Send an email to " & strEmail & """>"
                    Else
                        FormatEmail &= " title=""Send an email to " & strName.Trim & """>"
                    End If
                Else
                    FormatEmail &= " title=""" & strTitle.Trim & """>"
                End If

                'add the link text
                If strText.Trim = "" Then
                    FormatEmail &= strEmail & "</a>"
                Else
                    FormatEmail &= strText.Trim & "</a>"
                End If
            Else
                FormatEmail = ""
            End If
        End Function

        ''' <summary>
        ''' Formats a numeric string as a zip code
        ''' </summary>
        ''' <param name="strZip">The numeric string to be formatted</param>
        ''' <returns>The original string formatted as a zip code</returns>
        Public Function FormatZip(ByVal strZip As String) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Formats a numeric string as a zip code
            ' *****************************************************************************
            Dim intLen As Integer

            intLen = strZip.Trim.Length

            If intLen = 9 Then
                FormatZip = strZip.Substring(0, 5)
                FormatZip = FormatZip & "-" & strZip.Substring(5, 4)
            ElseIf intLen >= 5 Then
                FormatZip = strZip.Substring(0, 5)
            Else
                FormatZip = ""
            End If
        End Function

        ''' <summary>
        ''' Generates an html formatted address from the given address values
        ''' </summary>
        ''' <param name="strAddr1">Address line 1</param>
        ''' <param name="strAddr2">Address line 2</param>
        ''' <param name="strCity">City</param>
        ''' <param name="strST">State</param>
        ''' <param name="strZip">Zip code</param>
        ''' <returns>An html formatted address generated from the given address values</returns>
        Public Function FormatAddress(ByVal strAddr1 As String, ByVal strAddr2 As String, _
            ByVal strCity As String, ByVal strST As String, ByVal strZip As String) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Generates an html formatted address from the given address values
            ' *****************************************************************************

            'format the address
            FormatAddress = strAddr1
            If strAddr2 <> "" Then
                FormatAddress &= "<br />" & vbCrLf & strAddr2
            End If
            If strCity <> "" Or strST <> "" Or strZip <> "" Then
                FormatAddress &= "<br />" & vbCrLf
                If strCity <> "" Then FormatAddress &= strCity
                If strST <> "" And strCity <> "" Then
                    FormatAddress &= ", " & strST & " "
                ElseIf strST <> "" And strCity = "" Then
                    FormatAddress &= strST & " "
                End If
                FormatAddress &= strZip
            End If

            'trim off extra whitespace
            FormatAddress = FormatAddress.Trim
        End Function

#End Region

#Region " Dates "

        ''' <summary>
        ''' Converts a date string to a friendly date
        ''' </summary>
        ''' <param name="strDate">The date to be converted into a friendly date</param>
        ''' <param name="blnShortMonth">Boolean value indicating if the full month name should be used, or just a 3 letter abbreviation</param>
        ''' <param name="blnShortDay">Boolean value indicating if one or two digits should be used for the day</param>
        ''' <param name="blnShortYear">Boolean value indicating if two or four digits should be used for the year</param>
        ''' <param name="blnDateOnly">Boolean value idicating if only dates should be used, rather than weekdays or relative terms (today, tomorrow, yesterday, etc.)</param>
        ''' <returns>The original date reformatted in a friendly readable format</returns>
        Public Function FriendlyDate(ByVal strDate As String, Optional ByVal blnShortMonth As Boolean = True, _
            Optional ByVal blnShortDay As Boolean = True, Optional ByVal blnShortYear As Boolean = False, _
            Optional ByVal blnDateOnly As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts a date string to a friendly date
            ' *****************************************************************************
            If IsDate(strDate) Then
                Return FriendlyDate(CDate(strDate), blnShortMonth, blnShortDay, blnShortYear, blnDateOnly)
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' Return a friendly date string in relation to today's date
        ''' </summary>
        ''' <param name="objDate">The date to be converted into a friendly date</param>
        ''' <param name="blnShortMonth">Boolean value indicating if the full month name should be used, or just a 3 letter abbreviation</param>
        ''' <param name="blnShortDay">Boolean value indicating if one or two digits should be used for the day</param>
        ''' <param name="blnShortYear">Boolean value indicating if two or four digits should be used for the year</param>
        ''' <param name="blnDateOnly">Boolean value idicating if only dates should be used, rather than weekdays or relative terms (today, tomorrow, yesterday, etc.)</param>
        ''' <returns>The original date reformatted in a friendly readable format</returns>
        Public Function FriendlyDate(ByVal objDate As Date, Optional ByVal blnShortMonth As Boolean = True, _
            Optional ByVal blnShortDay As Boolean = True, Optional ByVal blnShortYear As Boolean = False, _
            Optional ByVal blnDateOnly As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Return a friendly date string in relation to today's date
            ' *****************************************************************************
            Dim datDueDate As Date = CDate(objDate.ToString("MM/dd/yyyy"))
            Dim datNow As Date = CDate(Now.ToString("MM/dd/yyyy"))
            Dim strMonth As String = "MMM"
            Dim strDay As String = "dd"
            Dim strYear As String = "yyyy"

            If Not blnShortMonth Then strMonth = "MMMM"
            If blnShortDay Then strDay = "d"
            If blnShortYear Then strYear = "yy"

            If datDueDate = datNow And Not blnDateOnly Then
                FriendlyDate = "Today"
            ElseIf datDueDate = datNow.AddDays(1) And Not blnDateOnly Then
                FriendlyDate = "Tomorrow"
            ElseIf datDueDate = datNow.AddDays(-1) And Not blnDateOnly Then
                FriendlyDate = "Yesterday"
            ElseIf DatePart(DateInterval.WeekOfYear, datDueDate) = DatePart(DateInterval.WeekOfYear, datNow) And datDueDate.Year = Now.Year And Not blnDateOnly Then
                FriendlyDate = datDueDate.DayOfWeek.ToString()
                If datDueDate < datNow Then
                    FriendlyDate = "Last " & FriendlyDate
                End If
            ElseIf datDueDate.Year <> Now.Year Then
                FriendlyDate = datDueDate.ToString(strMonth & " " & strDay & ", " & strYear)
            Else
                FriendlyDate = datDueDate.ToString(strMonth & " " & strDay)
            End If
        End Function

        Public Enum Specificity
            Specific = 1
            Medium = 2
            Vague = 3
        End Enum

        ''' <summary>
        ''' Converts a date/time string to a friendly date
        ''' </summary>
        ''' <param name="strDate">The date to be converted into a friendly date</param>
        ''' <returns>The original date reformatted in a friendly readable format</returns>
        Public Function FriendlyDateTime(ByVal strDate As String,
            Optional ByVal enmSpecificity As Specificity = Specificity.Medium) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2012.05.25
            ' Description:  Converts a date/time string to a friendly date
            ' *****************************************************************************
            If IsDate(strDate) Then
                Return FriendlyDateTime(CDate(strDate), enmSpecificity)
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' Return a friendly date and time string in relation to today's date
        ''' </summary>
        ''' <param name="objDate">The date to be converted into a friendly date</param>
        ''' <returns>The original date and time reformatted in a friendly readable format</returns>
        Public Function FriendlyDateTime(ByVal objDate As Date,
            Optional ByVal enmSpecificity As Specificity = Specificity.Medium) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2012.05.25
            ' Description:  Return a friendly date and time string in relation to today's date
            ' *****************************************************************************
            Dim datDateOnly As Date = CDate(objDate.ToString("MM/dd/yyyy"))
            Dim datToday As Date = CDate(Now.ToString("MM/dd/yyyy"))
            Dim strMonth As String = "MMMM"
            Dim strDay As String = "dddd"
            Dim strYear As String = "yyyy"

            FriendlyDateTime = ""

            If datDateOnly = datToday Then
                Select Case enmSpecificity
                    Case Specificity.Specific
                        FriendlyDateTime = "Today at " & objDate.ToString("h:mm tt")
                    Case Specificity.Medium
                        Dim intMinutes As Integer = DateDiff(DateInterval.Minute, Now, objDate)

                        If intMinutes = 0 Then
                            FriendlyDateTime = "Right Now"
                        ElseIf intMinutes > 0 Then
                            FriendlyDateTime = " From Now"
                            If intMinutes > 60 Then
                                Dim intHours As Integer = Math.Round(intMinutes / 60)
                                If intHours = 1 Then
                                    FriendlyDateTime = "An Hour" & FriendlyDateTime
                                Else
                                    FriendlyDateTime = intHours & " Hours" & FriendlyDateTime
                                End If
                            Else
                                If intMinutes = 1 Then
                                    FriendlyDateTime = "A Minute" & FriendlyDateTime
                                Else
                                    FriendlyDateTime = intMinutes & " Minutes" & FriendlyDateTime
                                End If
                            End If
                        ElseIf intMinutes < 0 Then
                            FriendlyDateTime = " Ago"
                            If intMinutes < -60 Then
                                intMinutes = intMinutes * -1
                                Dim intHours As Integer = Math.Round(intMinutes / 60)
                                If intHours = 1 Then
                                    FriendlyDateTime = "An Hour" & FriendlyDateTime
                                Else
                                    FriendlyDateTime = intHours & " Hours" & FriendlyDateTime
                                End If
                            Else
                                intMinutes = intMinutes * -1
                                If intMinutes = 1 Then
                                    FriendlyDateTime = "A Minute" & FriendlyDateTime
                                Else
                                    FriendlyDateTime = intMinutes & " Minutes" & FriendlyDateTime
                                End If
                            End If
                        End If
                    Case Specificity.Vague
                        If objDate > CDate(Now.ToString("MM/d/yyyy") & " 4:59:00 pm") Then
                            FriendlyDateTime = "Tonight"
                        ElseIf objDate > CDate(Now.ToString("MM/d/yyyy") & " 12:00:00 pm") Then
                            FriendlyDateTime = "This Afternoon"
                        Else
                            FriendlyDateTime = "This Morning"
                        End If
                End Select
            ElseIf datDateOnly = datToday.AddDays(1) Then
                Select Case enmSpecificity
                    Case Specificity.Specific
                        FriendlyDateTime = "Tomorrow at " & objDate.ToString("h:mm tt")
                    Case Specificity.Medium
                        If objDate > CDate(datToday.AddDays(1).ToString("MM/d/yyyy") & " 4:59:00 pm") Then
                            FriendlyDateTime = "Tomorrow Night"
                        ElseIf objDate > CDate(datToday.AddDays(1).ToString("MM/d/yyyy") & " 12:00:00 pm") Then
                            FriendlyDateTime = "Tomorrow Afternoon"
                        Else
                            FriendlyDateTime = "Tomorrow Morning"
                        End If
                    Case Specificity.Vague
                        FriendlyDateTime = "Tomorrow"
                End Select
            ElseIf datDateOnly = datToday.AddDays(-1) Then
                Select Case enmSpecificity
                    Case Specificity.Specific
                        FriendlyDateTime = "Yesterday at " & objDate.ToString("h:mm tt")
                    Case Specificity.Medium
                        If objDate > CDate(datToday.AddDays(-1).ToString("MM/d/yyyy") & " 4:59:00 pm") Then
                            FriendlyDateTime = "Last Night"
                        ElseIf objDate > CDate(datToday.AddDays(-1).ToString("MM/d/yyyy") & " 12:00:00 pm") Then
                            FriendlyDateTime = "Yesterday Afternoon"
                        Else
                            FriendlyDateTime = "Yesterday Morning"
                        End If
                    Case Specificity.Vague
                        FriendlyDateTime = "Yesterday"
                End Select
            ElseIf DateDiff(DateInterval.Day, Now, objDate) >= -7 And DateDiff(DateInterval.Day, Now, objDate) <= 7 Then
                Dim intDays As Integer = DateDiff(DateInterval.Day, Now, objDate)
                Select Case enmSpecificity
                    Case Specificity.Specific
                        If intDays < -2 Then
                            FriendlyDateTime = "Last " & objDate.ToString("dddd") & " at " & objDate.ToString("h:mm tt")
                        ElseIf intDays < 5 Then
                            FriendlyDateTime = objDate.ToString("dddd") & " at " & objDate.ToString("h:mm tt")
                        ElseIf intDays >= 5 Then
                            FriendlyDateTime = "Next " & objDate.ToString("dddd") & " at " & objDate.ToString("h:mm tt")
                        End If
                    Case Specificity.Medium
                        If intDays < -2 Then
                            FriendlyDateTime = "Last " & objDate.ToString("dddd")
                        ElseIf intDays < 5 Then
                            FriendlyDateTime = objDate.ToString("dddd")
                        ElseIf intDays >= 5 Then
                            FriendlyDateTime = "Next " & objDate.ToString("dddd")
                        End If
                    Case Specificity.Vague
                        If intDays < -2 Then
                            FriendlyDateTime = "A Few Days Ago"
                        ElseIf intDays = -2 Then
                            FriendlyDateTime = "A Couple Days Ago"
                        ElseIf intDays = -1 Then
                            FriendlyDateTime = "Yesterday"
                        ElseIf intDays = 0 Then
                            FriendlyDateTime = "Today"
                        ElseIf intDays = 1 Then
                            FriendlyDateTime = "Tomorrow"
                        ElseIf intDays = 2 Then
                            FriendlyDateTime = "A Couple Days From Now"
                        ElseIf intDays > 2 Then
                            FriendlyDateTime = "A Few Days From Now"
                        End If
                End Select
            ElseIf DateDiff(DateInterval.Month, Now, objDate) > -6 And DateDiff(DateInterval.Month, Now, objDate) < 6 And objDate.Year = Now.Year Then
                Select Case enmSpecificity
                    Case Specificity.Specific
                        FriendlyDateTime = objDate.ToString("MMMM d") & " at " & objDate.ToString("h:mm tt")
                    Case Specificity.Medium
                        FriendlyDateTime = objDate.ToString("MMMM d")
                    Case Specificity.Vague
                        Dim intMonths As Integer = DateDiff(DateInterval.Month, Now, objDate)
                        If intMonths < -4 Then
                            FriendlyDateTime = "Several Months Ago"
                        ElseIf intMonths < -2 Then
                            FriendlyDateTime = "A Few Months Ago"
                        ElseIf intMonths = -2 Then
                            FriendlyDateTime = "A Couple Months Ago"
                        ElseIf intMonths = -1 Then
                            FriendlyDateTime = "Last Month"
                        ElseIf intMonths = 0 Then
                            If DateDiff(DateInterval.Day, Now, objDate) < 0 Then
                                FriendlyDateTime = "Earlier This Month"
                            Else
                                FriendlyDateTime = "Later This Month"
                            End If
                        ElseIf intMonths = 1 Then
                            FriendlyDateTime = "Next Month"
                        ElseIf intMonths = 2 Then
                            FriendlyDateTime = "A Couple Months From Now"
                        ElseIf intMonths < 5 Then
                            FriendlyDateTime = "A Few Months From Now"
                        Else
                            FriendlyDateTime = "Several Months From Now"
                        End If
                End Select
            Else
                Select Case enmSpecificity
                    Case Specificity.Specific
                        FriendlyDateTime = objDate.ToString("MMMM d, yyyy") & " at " & objDate.ToString("h:mm tt")
                    Case Specificity.Medium
                        FriendlyDateTime = objDate.ToString("MMMM d, yyyy")
                    Case Specificity.Vague
                        Dim intYears As Integer = DateDiff(DateInterval.Year, Now, objDate)
                        If intYears < -4 Then
                            FriendlyDateTime = "Several Years Ago"
                        ElseIf intYears < -2 Then
                            FriendlyDateTime = "A Few Years Ago"
                        ElseIf intYears = -2 Then
                            FriendlyDateTime = "A Couple Years Ago"
                        ElseIf intYears = -1 Then
                            FriendlyDateTime = "Last Year"
                        ElseIf intYears = 0 Then
                            If DateDiff(DateInterval.Day, Now, objDate) < 0 Then
                                FriendlyDateTime = "Earlier This Year"
                            Else
                                FriendlyDateTime = "Later This Year"
                            End If
                        ElseIf intYears = 1 Then
                            FriendlyDateTime = "Next Year"
                        ElseIf intYears = 2 Then
                            FriendlyDateTime = "A Couple Years From Now"
                        ElseIf intYears < 5 Then
                            FriendlyDateTime = "A Few Years From Now"
                        Else
                            FriendlyDateTime = "Several Years From Now"
                        End If
                End Select
            End If
        End Function

        ''' <summary>
        ''' Formats a date based on the given format string
        ''' </summary>
        ''' <param name="objDate">The raw date to be formatted</param>
        ''' <param name="strFormat">The format string to use for formatting the date</param>
        ''' <param name="blnLowercaseAMPM">Boolean value indicating if the AM/PM part should be lowercase</param>
        ''' <returns>A string containing the formatted date</returns>
        ''' <remarks>Returns an empty string if the date is not valid</remarks>
        Public Function FormatDate(ByVal objDate As Object, ByVal strFormat As String, Optional ByVal blnLowercaseAMPM As Boolean = True) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Formats a date based on the given format string
            ' Modified:     2011.01.24 by JR - Added lowercase AM/PM option
            ' *****************************************************************************
            Dim datInputDate As DateTime
            Dim strDate As String

            If objDate & "" = "" Then
                Return ""
            ElseIf Not IsDate(objDate) Then
                Return ""
            Else
                datInputDate = CDate(objDate)
            End If

            If blnLowercaseAMPM And strFormat.Contains("tt") Then
                strDate = datInputDate.ToString(strFormat.Replace("tt", "\t\t"))
                strDate = strDate.Replace("tt", datInputDate.ToString("tt").ToLower)
            Else
                strDate = datInputDate.ToString(strFormat)
            End If

            Return strDate
        End Function

        ''' <summary>
        ''' Returns the portion of the day that the given time falls within
        ''' </summary>
        ''' <param name="datTime">The time to be formatted</param>
        ''' <returns>A friendly name of the portion of the day the the given time falls within</returns>
        Public Function TimeOfDay(ByVal datTime As Date) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Returns the portion of the day that the given time falls within
            ' *****************************************************************************
            Dim objTime As TimeSpan = datTime.TimeOfDay

            Select Case objTime
                Case Is < New TimeSpan(12, 0, 0)
                    TimeOfDay = "Morning"
                Case Is < New TimeSpan(17, 0, 0)
                    TimeOfDay = "Afternoon"
                Case Is <= New TimeSpan(0, 23, 59, 59, 99999)
                    TimeOfDay = "Evening"
                Case Else
                    TimeOfDay = "Morning"
            End Select
        End Function

#End Region

#Region " Miscellaneous "

        ''' <summary>
        ''' Converts a boolean value to a more readable string format
        ''' </summary>
        ''' <param name="blnValue">The boolean value to be converted</param>
        ''' <param name="objType">The string format to be used</param>
        ''' <returns>A string representation of the boolean value</returns>
        Public Function BooleanToString(ByVal blnValue As Boolean, Optional ByVal objType As BoolStringTypes = BoolStringTypes.TrueFalse) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts a boolean value to a more readable string format
            ' *****************************************************************************
            Select Case objType
                Case BoolStringTypes.TrueFalse
                    If blnValue Then
                        BooleanToString = "True"
                    Else
                        BooleanToString = "False"
                    End If
                Case BoolStringTypes.YesNo
                    If blnValue Then
                        BooleanToString = "Yes"
                    Else
                        BooleanToString = "No"
                    End If
                Case BoolStringTypes.ActiveInactive
                    If blnValue Then
                        BooleanToString = "Active"
                    Else
                        BooleanToString = "Inactive"
                    End If
                Case BoolStringTypes.OnlineOffline
                    If blnValue Then
                        BooleanToString = "Online"
                    Else
                        BooleanToString = "Offline"
                    End If
                Case Else
                    If blnValue Then
                        BooleanToString = "True"
                    Else
                        BooleanToString = "False"
                    End If
            End Select
        End Function

#End Region

    End Module

End Namespace