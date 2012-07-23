Imports System.IO

Namespace Files

    Public Module mFiles

        ''' <summary>
        ''' Moves or renames one file to another
        ''' </summary>
        ''' <param name="strOldPath">The path to the current location of the file</param>
        ''' <param name="strNewPath">The path to the new location of the file</param>
        Public Sub MoveFile(ByVal strOldPath As String, ByVal strNewPath As String)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Moves or renames one file to another
            ' *****************************************************************************
            Dim objFile As New FileInfo(strOldPath)

            If objFile.Exists Then
                File.Move(strOldPath, strNewPath)
            Else
                Throw New FileNotFoundException
            End If
        End Sub

        ''' <summary>
        ''' Returns the file name of the given file
        ''' </summary>
        ''' <param name="strPath">The path of the file to extract the name from</param>
        ''' <param name="blnIncludeExt">Boolena value indicating if the file extension should be included</param>
        ''' <returns>Just the name of the file specified in the path</returns>
        Public Function FileName(ByVal strPath As String, Optional ByVal blnIncludeExt As Boolean = False) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Returns the file name of the given file
            ' *****************************************************************************
            Dim objFile As New FileInfo(strPath)

            If objFile.Exists Then
                FileName = objFile.Name
            Else
                Throw New FileNotFoundException("Cannot extract file name.", strPath)
            End If

            If Not blnIncludeExt Then
                FileName = FileName.Replace(objFile.Extension, "")
            End If
        End Function

        ''' <summary>
        ''' Returns the file extension of the given file
        ''' </summary>
        ''' <param name="strPath">The path of the file to extract the file extension from</param>
        ''' <param name="blnIncludeDot">Boolena value indicating if the period before the extension should be included</param>
        ''' <returns>Just the extension part of the file specified in the path</returns>
        Public Function FileExtension(ByVal strPath As String, Optional ByVal blnIncludeDot As Boolean = True) As String
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Returns the file extension of the given file
            ' *****************************************************************************
            Dim objFile As New FileInfo(strPath)

            If objFile.Exists Then
                FileExtension = objFile.Extension
            Else
                Throw New FileNotFoundException("Cannot extract file extension.", strPath)
            End If

            If Not blnIncludeDot Then
                FileExtension = FileExtension.Replace(".", "")
            End If
        End Function

        ''' <summary>
        ''' Returns the size in bytes of the given file
        ''' </summary>
        ''' <param name="strPath">The path of the file to get the file size for</param>
        ''' <returns>The size (in bytes) of the file specified in the path</returns>
        Public Function FileSize(ByVal strPath As String) As Long
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Returns the size in bytes of the given file
            ' *****************************************************************************
            Dim objFile As New FileInfo(strPath)

            If objFile.Exists Then
                FileSize = objFile.Length
            Else
                Throw New FileNotFoundException("Cannot find the file size.", strPath)
            End If
        End Function

    End Module

End Namespace