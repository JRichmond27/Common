Imports System.Web
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Clean

    Public Module mCleanData

        ''' <summary>
        ''' Removes all special characters from a string
        ''' </summary>
        ''' <param name="strPhrase">The string to remove special characters from</param>
        ''' <param name="blnAllowPunctuation">Boolean value indicating if punctuation should be allowed</param>
        ''' <returns>The original string with all special characters removed</returns>
        Public Function RemoveSpecChars(ByVal strPhrase As String, Optional ByVal blnAllowPunctuation As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Removes all special characters from a string
            ' *****************************************************************************
            Dim sbEditPhrase As New StringBuilder()
            Dim intCount As Integer = 0

            For intCount = 0 To strPhrase.Length - 1
                If Char.IsLetterOrDigit(strPhrase(intCount)) Or Char.IsWhiteSpace(strPhrase(intCount)) Or _
                    (Char.IsPunctuation(strPhrase(intCount)) And blnAllowPunctuation) Then

                    sbEditPhrase.Append(strPhrase(intCount))
                End If
            Next intCount

            RemoveSpecChars = sbEditPhrase.ToString
        End Function

        ''' <summary>
        ''' Removes all non-numeric characters from a string
        ''' </summary>
        ''' <param name="strString">The numeric string to be cleaned up</param>
        ''' <param name="blnAllowNegative">Boolean value indicating if negative values should be allowed</param>
        ''' <param name="blnAllowDecimal">Boolean value indicating if decimal places should be allowed</param>
        ''' <param name="blnReplaceZero">Boolean value indicating if a zero should be returned in the case of a non-numeric result</param>
        ''' <param name="blnAllowTwoDecimals">Boolean value indicating if two decimals should be allowed</param>
        ''' <param name="blnAllowHex">Boolean value indicating if the string contains a hexidecimal value</param>
        ''' <returns>The original string with all non-numeric characters stripped out</returns>
        Public Function NumbersOnly(ByVal strString As String, Optional ByVal blnAllowNegative As Boolean = False, _
            Optional ByVal blnAllowDecimal As Boolean = False, Optional ByVal blnReplaceZero As Boolean = False, _
            Optional ByVal blnAllowTwoDecimals As Boolean = False, Optional ByVal blnAllowHex As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Removes all non-numeric characters from a string
            ' *****************************************************************************
            Dim strNumbers As String = ""
            Dim blnNegative As Boolean = False
            Dim blnDecimal As Boolean = False
            Dim intDecimalCount As Integer = 2
            Dim blnHex As Boolean
            Dim i As Integer

            'make sure the string is not null
            strString = strString & ""

            If strString <> "" Then
                'check if the first character is a negative sign
                If blnAllowNegative And strString.Substring(0, 1) = "-" Then
                    blnNegative = True
                End If
            End If

            'extract only the numbers
            For i = 0 To strString.Length - 1
                If IsNumeric(strString.Substring(i, 1)) Then
                    strNumbers = strNumbers & strString.Substring(i, 1)
                ElseIf blnAllowDecimal And strString.Substring(i, 1) = "." And Not blnDecimal Then
                    strNumbers = strNumbers & "."
                    blnDecimal = True
                ElseIf blnAllowTwoDecimals And strString.Substring(i, 1) = "." And intDecimalCount > 0 Then
                    strNumbers = strNumbers & "."
                    intDecimalCount -= 1
                ElseIf blnAllowHex And strString.Substring(i, 1) = "#" And Not blnHex Then
                    strNumbers = strNumbers & strString.Substring(i, 1)
                    blnHex = True
                ElseIf blnAllowHex And Regex.IsMatch(strString.Substring(i, 1), "[A-Fa-f0-9]") Then
                    strNumbers = strNumbers & strString.Substring(i, 1).ToUpper
                End If
            Next i

            'put the negative sign back in
            If blnNegative And strNumbers <> "" And strNumbers <> "0" Then
                strNumbers = "-" & strNumbers
            End If

            'make sure that more than just a decimal is returned
            If strNumbers = "." Then strNumbers = ""

            'make sure the returned value is not empty
            If strNumbers = "" And blnReplaceZero Then strNumbers = 0

            Return strNumbers
        End Function

        ''' <summary>
        ''' Removes all numeric characters from a string
        ''' </summary>
        ''' <param name="strString">The string to be cleaned</param>
        ''' <returns>The original string with all numeric charactes stripped out</returns>
        Public Function NonNumbersOnly(ByVal strString As String) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Removes all numeric characters from a string
            ' *****************************************************************************
            Dim strNumbers As String = ""
            Dim i As Integer

            'make sure the string is not null
            strString = strString & ""

            'extract only the numbers
            For i = 0 To strString.Length - 1
                If Not IsNumeric(strString.Substring(i, 1)) Then
                    strNumbers = strNumbers & strString.Substring(i, 1)
                End If
            Next i

            Return strNumbers
        End Function

        ''' <summary>
        ''' Substitutes empty, blank, or null values with a DBNull value
        ''' </summary>
        ''' <param name="datValue">The date object to be substituted</param>
        ''' <returns>The original date object, or DBNull in the case of empty, blank or null values</returns>
        Public Function SubstituteNull(ByVal datValue As Date) As Object
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2011.09.29
            ' Description:  Substitutes empty, blank, or null values with a DBNull value
            ' *****************************************************************************
            If datValue = Nothing Then
                SubstituteNull = DBNull.Value
            ElseIf IsDate(datValue) AndAlso CDate(datValue) < CDate("1/1/1900") Then
                SubstituteNull = DBNull.Value
            Else
                SubstituteNull = datValue.ToString.Trim
            End If
        End Function

        ''' <summary>
        ''' Substitutes empty, blank, or null values with a DBNull value
        ''' </summary>
        ''' <param name="objValue">The object to be substituted</param>
        ''' <returns>The original object, or DBNull in the case of empty, blank or null values</returns>
        Public Function SubstituteNull(ByVal objValue As Object) As Object
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Substitutes empty, blank, or null values with a DBNull value
            ' Modified:     2011.09.29 by JR - Moved date object substitution to it's own function
            ' *****************************************************************************
            If objValue Is Nothing Then
                SubstituteNull = DBNull.Value
            ElseIf objValue.ToString.Trim = "" Or objValue.ToString.Trim = "null" Then
                SubstituteNull = DBNull.Value
            Else
                SubstituteNull = objValue.ToString.Trim
            End If
        End Function

        ''' <summary>
        ''' Checks if a value is null. If it is null, then it is replaced with an alternate string. Returns the original value if not null.
        ''' </summary>
        ''' <param name="strValue">The string to be checked for a null value</param>
        ''' <param name="strAlt">An alternate string to return in case of a null value</param>
        ''' <returns>The original string if not null, otherwise the alternate string</returns>
        Public Function IsNull(ByVal strValue As String, Optional ByVal strAlt As String = "") As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks if a value is null. If it is null, then it is replaced 
            '               with an alternate string. Returns the original value if not null.
            ' *****************************************************************************
            Return IsNull("", strValue, "", strAlt)
        End Function

        ''' <summary>
        ''' Checks if a value is null. If it is null, then it is replaced with an alternate string. Returns the original value with the pre and post strings appended if not null.
        ''' </summary>
        ''' <param name="strPre">String value to prepend to the returned value if not null</param>
        ''' <param name="strValue">The string to be checked for a null value</param>
        ''' <param name="strPost">String value to append to the returned value if not null</param>
        ''' <param name="strAlt">An alternate string to return in case of a null value</param>
        ''' <returns>The original string with the pre and post strings appended if not null, otherwise the alternate string</returns>
        Public Function IsNull(ByVal strPre As String, ByVal strValue As String, ByVal strPost As String, Optional ByVal strAlt As String = "") As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks if a value is null. If it is null, then it is replaced 
            '               with an alternate string. Returns the original value with the 
            '               pre and post strings appended if not null.
            ' *****************************************************************************
            If (strValue & "").Trim = "" Then
                IsNull = strAlt
            Else
                IsNull = strPre & strValue & strPost
            End If
        End Function

        ''' <summary>
        ''' Truncates a string at the whitespace character nearest to the specified length
        ''' </summary>
        ''' <param name="strValue">The string to be truncated</param>
        ''' <param name="intLength">The desired length of the resulting string</param>
        ''' <param name="blnEllipses">Boolean value indicating if an ellises should be added to the end</param>
        ''' <returns>The original string, truncated near the desired length</returns>
        Public Function Truncate(ByVal strValue As String, ByVal intLength As Integer, Optional ByVal blnEllipses As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Truncates a string to the end of the last word within the 
            '               requested character length; encodes html special characters
            ' *****************************************************************************

            'check if the string is shorter than the requested length
            If strValue.Length > intLength Then
                Dim strTemp As String = strValue.Substring(0, intLength + 1)
                Dim intTempLength As Integer = 0

                'check if the last word ends exactly at the specified length
                If Regex.IsMatch(strTemp, "^.+[^a-z^0-9]$", RegexOptions.IgnoreCase Or RegexOptions.Singleline) Then
                    'the last word ends at the specified length
                    'just return the string trimmed to the requested length
                    strTemp = strTemp.Substring(0, intLength)
                Else
                    'get the location of the last space character
                    intTempLength = strTemp.LastIndexOf(" ")

                    If intTempLength > 0 Then
                        'trim to the last occurance of a space character
                        strTemp = strTemp.Substring(0, intTempLength)
                    Else
                        'there are no spaces, so trim to the specified length
                        strTemp = strTemp.Substring(0, intLength)
                    End If
                End If

                'get the current length
                intTempLength = strTemp.Length

                'trim off any leftover punctuation
                While Regex.IsMatch(strTemp, "^.+[^a-z^0-9]$", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                    strTemp = strTemp.Substring(0, strTemp.Length - 1)

                    'check if the string is empty
                    If strTemp = "" Then
                        'set the string to the previous length and break out of the loop
                        strTemp = strValue.Substring(0, intTempLength)
                        Exit While
                    End If
                End While

                'add ellipses if requested
                If blnEllipses Then
                    Truncate = HttpUtility.HtmlEncode(strTemp) & "&#8230;"
                Else : Truncate = HttpUtility.HtmlEncode(strTemp)
                End If
            Else
                Truncate = HttpUtility.HtmlEncode(strValue)
            End If
        End Function

        ''' <summary>
        ''' Converts the requested file name into a clean file name
        ''' </summary>
        ''' <param name="strName">The file name to be cleaned up</param>
        ''' <param name="blnRemoveSpaces">Boolean value indicating if spaces should be removed</param>
        ''' <param name="intLength">The maximum length of the file name</param>
        ''' <param name="blnRemoveCommonWords">Boolean value indicating if certain common words should be removed</param>
        ''' <param name="blnRemovePeriods">Boolean value indicating if periods should be removed</param>
        ''' <param name="blnToLower">Boolean value indicating if the result should be changed to lower case</param>
        ''' <returns>The original string, cleaned up to work as a nice file name</returns>
        Public Function CleanFileName(ByVal strName As String, Optional ByVal blnRemoveSpaces As Boolean = False, _
            Optional ByVal intLength As Integer = 75, Optional ByVal blnRemoveCommonWords As Boolean = True, _
            Optional ByVal blnRemovePeriods As Boolean = False, Optional ByVal blnToLower As Boolean = True) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  converts the requested file name into a clean file name
            ' Modified:     2011.06.23 by JR - Fixed a bug that caused the returned text to 
            '               always be lowercased
            ' *****************************************************************************
            Dim strTemp As String = ""

            If blnToLower Then
                strTemp = Regex.Replace(strName.Trim.ToLower, "[^a-z^0-9^ ^.^_^-]", "", RegexOptions.Singleline)
            Else
                strTemp = Regex.Replace(strName.Trim, "[^a-z^A-Z^0-9^ ^.^_^-]", "", RegexOptions.Singleline)
            End If

            'replace multiple whitespaces in a row
            strTemp = Regex.Replace(strTemp, "\s{2,}", " ", RegexOptions.Singleline)

            'remove common words from the name
            If blnRemoveCommonWords Then
                strTemp = Regex.Replace(strTemp, "(^|\s)(and|or|of|a|an)(?=(\s|$))", "", RegexOptions.Singleline)
            End If

            'remove periods
            If blnRemovePeriods Then
                strTemp = strTemp.Replace(".", "")
            End If

            'truncate the length of the name
            strTemp = Truncate(strTemp, intLength)

            'remove or replace all spaces
            If blnRemoveSpaces Then
                strTemp = strTemp.Replace(" ", "")
                strTemp = strTemp.Replace("_", "")
            Else
                strTemp = strTemp.Replace(" ", "-")
                strTemp = strTemp.Replace("_", "-")
            End If

            'replace multiple dashes in a row with a single dash
            strTemp = Regex.Replace(strTemp, "\-{2,}", "-", RegexOptions.Singleline)

            CleanFileName = strTemp
        End Function

        ''' <summary>
        ''' Removes or cleans up all html in a given string
        ''' </summary>
        ''' <param name="strValue">The string to be cleaned up</param>
        ''' <param name="blnSafeHTML">Boolean value indicating if html should be removed or simply made "safe"</param>
        ''' <returns></returns>
        Public Function HTMLClean(ByVal strValue As String, Optional ByVal blnSafeHTML As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Removes or cleans up all html in a given string
            ' *****************************************************************************
            If blnSafeHTML Then
                strValue = Regex.Replace(strValue, "<h1[^>]*>", "<h1>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h1[^>]*>", "</h1>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<h2[^>]*>", "<h2>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h2[^>]*>", "</h2>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<h3[^>]*>", "<h3>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h3[^>]*>", "</h3>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<h4[^>]*>", "<h4>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h4[^>]*>", "</h4>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<h5[^>]*>", "<h5>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h5[^>]*>", "</h5>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<h6[^>]*>", "<h6>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</h6[^>]*>", "</h6>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<b[^>]*>", "<strong>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</b[^>]*>", "</strong>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<strong[^>]*>", "<strong>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</strong[^>]*>", "</strong>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<i[^>]*>", "<em>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</i[^>]*>", "</em>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<em[^>]*>", "<em>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</em[^>]*>", "</em>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<p[^>]*>", "<p>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</p[^>]*>", "</p>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<ul[^>]*>", "<ul>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</ul[^>]*>", "</ul>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<ol[^>]*>", "<ol>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</ol[^>]*>", "</ol>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<li[^>]*>", "<li>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</li[^>]*>", "</li>", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</?font[^>]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</?div[^>]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</?span[^>]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "</?form[^>]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<script[^>]*>.*?</script[^>]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<script[^>]*>.*?</script[^>]*>", " ", RegexOptions.None)
            Else
                strValue = Regex.Replace(strValue, "</[^>]*><[^>^/]*>", " ", RegexOptions.None)
                strValue = Regex.Replace(strValue, "<[^>]*>", "", RegexOptions.None)
                strValue = Web.HttpUtility.HtmlDecode(strValue)
            End If

            Return strValue
        End Function

        ''' <summary>
        ''' Checks a url and makes sure that it includes the http protocol
        ''' </summary>
        ''' <param name="strURL">The url to be checked</param>
        ''' <param name="blnSecure">Boolean value indicating if https should be used instead</param>
        ''' <returns>The original URL with the correct protocol prefix</returns>
        Public Function FixURL(ByVal strURL As String, Optional ByVal blnSecure As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Checks a url and makes sure that it includes the http protocol
            ' *****************************************************************************
            Dim strProtocol = "http://"

            If blnSecure Then strProtocol = "https://"

            If strURL.IndexOf(strProtocol) < 0 Then
                Return strProtocol & strURL
            Else
                Return strURL
            End If
        End Function

        ''' <summary>
        ''' Converts a string into a valid date object
        ''' </summary>
        ''' <param name="strDate">The string to be interpreted as a date</param>
        ''' <returns>The original string as a date object</returns>
        ''' <remarks>Returns nothing if the string can not be cast to a valid date</remarks>
        Public Function StringToDate(ByVal strDate As String) As Date
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Converts a string into a valid date object
            ' *****************************************************************************
            If (strDate & "").Trim = "" Or Not IsDate(strDate) Then
                StringToDate = Nothing
            Else
                StringToDate = CDate(strDate.Trim)
            End If
        End Function

    End Module

End Namespace