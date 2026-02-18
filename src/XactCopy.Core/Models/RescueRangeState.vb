Namespace Models
    ''' <summary>
    ''' Coverage state for a logical byte range in a file.
    ''' </summary>
    Public Enum RescueRangeState
        Pending = 0
        Good = 1
        Bad = 2
        Recovered = 3

        ''' <summary>
        ''' Pre-seeded from bad-range map; worker can skip read attempts when configured.
        ''' </summary>
        KnownBad = 4
    End Enum
End Namespace
