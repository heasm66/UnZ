Public Enum MemoryMapType
    MM_HEADER_TABLE
    MM_HEADER_EXT_TABLE
    MM_ABBREVIATION_STRINGS
    MM_ABBREVIATION_TABLE
    MM_PROPERTY_DEFAULTS_TABLE
    MM_OBJECT_TREE_TABLE
    MM_OBJECT_PROPERTIES_TABLES
    MM_GLOBAL_VARIABLES
    MM_TERMINATING_CHARS_TABLE
    MM_GRAMMAR_TABLE
    MM_GRAMMAR_TABLE_DATA
    MM_ACTION_TABLE
    MM_PREPOSITION_TABLE
    MM_PREPOSITION_TABLE_COUNT
    MM_PREACTION_TABLE
    MM_PREACTION_PARSING_TABLE
    MM_DICTIONARY
    MM_ZCODE
    MM_STATIC_STRINGS
    MM_CHRSET
    MM_IFID
    MM_UNIDENTIFIED_DATA
    MM_PADDING
    MM_WORD_FLAGS_TABLE
End Enum

Public Class MemoryMapEntry
    Public Sub New(pName As String, pAddressStart As Integer, pAddressEnd As Integer, pType As MemoryMapType)
        name = pName
        addressStart = pAddressStart
        addressEnd = pAddressEnd
        type = pType
    End Sub

    Public name As String = ""
    Public addressStart As Integer = 0
    Public addressEnd As Integer = 0
    Public type As MemoryMapType = MemoryMapType.MM_UNIDENTIFIED_DATA
    Public startOfDynamic As Boolean = False
    Public startOfStatic As Boolean = False
    Public startOfHigh As Boolean = False


    Public ReadOnly Property SizeString As String
        Get
            Dim size As Integer = addressEnd - addressStart + 1
            If size = 1 Then Return "1 byte"
            Return size.ToString("#,##0 bytes")
        End Get
    End Property

    Public Sub PrintMemoryLabel()
        If startOfDynamic Then
            Console.WriteLine()
            Console.WriteLine("*****************************************************")
            Console.WriteLine("********** START OF DYNAMIC MEMORY 0x{0:X5} **********", addressStart)
            Console.WriteLine("*****************************************************")
            Console.WriteLine()
        End If
        If startOfStatic Then
            Console.WriteLine()
            Console.WriteLine("****************************************************")
            Console.WriteLine("********** START OF STATIC MEMORY 0x{0:X5} **********", addressStart)
            Console.WriteLine("****************************************************")
            Console.WriteLine()
        End If
        If startOfHigh Then
            Console.WriteLine()
            Console.WriteLine("**************************************************")
            Console.WriteLine("********** START OF HIGH MEMORY 0x{0:X5} **********", addressStart)
            Console.WriteLine("**************************************************")
            Console.WriteLine()
        End If
    End Sub
End Class
