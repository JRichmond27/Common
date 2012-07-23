Namespace Enums

    Public Module mEnumerations

        ''' <summary>
        ''' Message box content types
        ''' </summary>
        Public Enum MsgType
            Ok = 0
            Info = 1
            Warn = 2
            Err = 3
            Query = 4
        End Enum

        ''' <summary>
        ''' Message box screen locations
        ''' </summary>
        Public Enum MsgLoc
            TopLeft = 0
            TopCenter = 1
            TopRight = 2
            Center = 3
            BotLeft = 4
            BotCenter = 5
            BotRight = 6
        End Enum

        ''' <summary>
        ''' Types of changes that can be made to database records
        ''' </summary>
        Public Enum SQLChangeType
            Insert = 0
            Update = 1
            Delete = 2
        End Enum

        ''' <summary>
        ''' Possible function return statuses
        ''' </summary>
        Public Enum RetStatus
            Fail = 0
            Pass = 1
            Err = 2
        End Enum

        ''' <summary>
        ''' Possible string formats for boolean values
        ''' </summary>
        Public Enum BoolStringTypes
            TrueFalse = 0
            YesNo = 1
            ActiveInactive = 2
            OnlineOffline = 3
        End Enum

        ''' <summary>
        ''' File size units
        ''' </summary>
        Public Enum FileSizeUnits
            B = 1
            KB = 2
            MB = 3
            GB = 4
            TB = 5
        End Enum

        ''' <summary>
        ''' Email address types
        ''' </summary>
        Public Enum EmailType
            EmailTo = 0
            EmailFrom = 1
            EmailCC = 2
            EmailBCC = 3
        End Enum

    End Module

End Namespace