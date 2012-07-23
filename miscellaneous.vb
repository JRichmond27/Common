Imports System.Text

Namespace Misc

    Public Module mMiscellaneous

        ''' <summary>
        ''' Generates a random string of numbers for use as a key
        ''' </summary>
        ''' <param name="intLength">The length of the key to be generated</param>
        ''' <param name="blnHex">Boolean value indicating if hexidecimal numbers should be included</param>
        ''' <returns>A random string of numbers to be used as a key</returns>
        Public Function GenerateRandomKey(ByVal intLength As Integer, Optional ByVal blnHex As Boolean = True) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Generates a random string of numbers for use as a key
            ' *****************************************************************************
            Dim i As Integer
            Dim strKey As New StringBuilder()
            Dim rndNumber As New Random()

            For i = 0 To intLength - 1
                If blnHex Then
                    strKey.Append(Hex(rndNumber.Next(0, 16)))
                Else
                    strKey.Append(rndNumber.Next(0, 9))
                End If
            Next i
            Return strKey.ToString()
        End Function

    End Module

End Namespace