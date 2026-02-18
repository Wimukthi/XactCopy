Namespace Models
    ''' <summary>
    ''' Determines whether a run performs byte transfer or read-only surface analysis.
    ''' </summary>
    Public Enum JobOperationMode
        ''' <summary>
        ''' Standard source-to-destination copy run.
        ''' </summary>
        Copy = 0

        ''' <summary>
        ''' Read-only scan that records unreadable ranges without writing output files.
        ''' </summary>
        ScanOnly = 1
    End Enum
End Namespace
