'MIT License

'Copyright(c) 2021-2024 Henrik Åsman

'Permission Is hereby granted, free Of charge, to any person obtaining a copy
'of this software And associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
'copies of the Software, And to permit persons to whom the Software Is
'furnished to do so, subject to the following conditions:

'The above copyright notice And this permission notice shall be included In all
'copies Or substantial portions of the Software.

'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
'LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
'OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN THE
'SOFTWARE.

Public Enum EnumCompilerSource
    ZILCH
    ZILF
    INFORM5
    INFORM6
    DIALOG
    UNKNOWN
End Enum

Public Enum EnumGrammarVer
    VERSION_1 = 1
    VERSION_2 = 2
    UNKNOWN = 0
End Enum
Public Class Helper
    Public Shared Function GetAdressFromWord(byteGame() As Byte, index As Integer) As Integer
        Return byteGame(index) * 256 + byteGame(index + 1)
    End Function

    Public Shared Function GetAdressFromPacked(byteGame() As Byte, index As Integer, isRoutine As Boolean) As Integer
        Dim byteAddress As Integer = GetAdressFromWord(byteGame, index)
        Return UnpackAddress(byteAddress, byteGame, isRoutine)
    End Function

    Public Shared Function UnpackAddress(packedAddress As Integer, byteGame() As Byte, isRoutine As Boolean) As Integer
        Dim zVersion As Integer = byteGame(0)
        Dim offset As Integer = GetAdressFromWord(byteGame, &H28)
        If Not isRoutine Then offset = GetAdressFromWord(byteGame, &H2A)
        Select Case zVersion
            Case 1, 2, 3
                Return 2 * packedAddress
            Case 4, 5
                Return 4 * packedAddress
            Case 6, 7
                Return 4 * packedAddress + 8 * offset
            Case 8
                Return 8 * packedAddress
            Case Else
                Return 2 * packedAddress
        End Select
    End Function

    Public Shared Function GetNextValidPackedAddress(byteGame() As Byte, address As Integer) As Integer
        Dim Zversion As Integer = byteGame(0)
        Dim scaler As Integer = 2
        Select Case Zversion
            Case 4, 5, 6, 7
                scaler = 4
            Case 8
                scaler = 8
        End Select
        'Dim offset As Integer = GetAdressFromWord(byteGame, &H28)
        'If Not IsRoutine Then offset = GetAdressFromWord(byteGame, &H2A)
        'offset = offset * 8
        Return CInt(Math.Truncate((address + scaler - 1) / scaler) * scaler) ' - offset
    End Function

    Public Shared Function ExtractZString(byteGame() As Byte, piStringAddress As Integer, sAbbreviations() As String, pAlphabet() As String, Optional pbHighlightAbbrevs As Boolean = False) As String
        Dim iCounter As Integer = 0
        Dim bLastW As Boolean
        Dim iWord As Integer
        Dim iAlpabeth As Integer = 0
        Dim sRet As String = ""
        Dim iAbbrevTable As Integer = -1
        Dim iZSCIIEscape As Integer = -1
        Dim iZSCIIEscapeCount As Integer = -1

        Do
            iWord = byteGame(piStringAddress + iCounter) * 256 + byteGame(piStringAddress + iCounter + 1)
            bLastW = CBool(iWord And 32768)
            For i As Integer = 2 To 0 Step -1
                Dim iZChar As Integer = CInt((iWord And CInt(((32 ^ i) * 31))) / (32 ^ i))
                If iAbbrevTable > -1 Then
                    ' Insert abbreviation
                    If pbHighlightAbbrevs Then
                        sRet = sRet & "{" & sAbbreviations((iAbbrevTable - 1) * 32 + iZChar) & "}"
                    Else
                        sRet &= sAbbreviations((iAbbrevTable - 1) * 32 + iZChar)
                    End If
                    iAbbrevTable = -1
                ElseIf iZSCIIEscapeCount <> -1 Then
                    iZSCIIEscapeCount += 1
                    If iZSCIIEscapeCount = 1 Then
                        iZSCIIEscape = iZChar
                    ElseIf iZSCIIEscapeCount = 2 Then
                        iZSCIIEscapeCount = -1
                        iAlpabeth = 0
                        iZSCIIEscape = 32 * iZSCIIEscape + iZChar
                        If iZSCIIEscape <= &H80 Then
                            sRet &= "                                 !~#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_'abcdefghijklmnopqrstuvwxyz{!}~ ".Substring(iZSCIIEscape, 1)
                        End If
                    End If
                ElseIf iZChar = 4 Then
                    iAlpabeth = 1
                ElseIf iZChar = 5 Then
                    iAlpabeth = 2
                ElseIf iZChar = 0 Then
                    sRet &= " "
                ElseIf iZChar < 6 Then
                    ' Insert abbreviation
                    iAbbrevTable = iZChar
                ElseIf iZChar = 6 And iAlpabeth = 2 Then
                    ' ZSCII Escape
                    iZSCIIEscapeCount += 1
                Else
                    Select Case iAlpabeth
                        Case 0
                            sRet &= pAlphabet(0).Substring(iZChar, 1)
                        Case 1
                            sRet &= pAlphabet(1).Substring(iZChar, 1)
                        Case 2
                            sRet &= pAlphabet(2).Substring(iZChar, 1)
                    End Select
                    iAlpabeth = 0
                End If
            Next
            iCounter += 2
        Loop Until bLastW

        Return sRet
    End Function
End Class

Public Class StringData
    Public name As String = ""
    Public number As Integer = 0
    Public text As String = ""
    Public textWithAbbrevs As String = ""
    Public entryPoint As Integer = 0
    Public endPoint As Integer = 0
    Public entryPointPacked As Integer = 0
    Public ReadOnly Property GetText(showAbbrevsInsertion As Boolean) As String
        Get
            If showAbbrevsInsertion Then Return textWithAbbrevs Else Return text
        End Get
    End Property
End Class

Public Class RoutineData
    Public name As String = ""
    Public entryPoint As Integer = 0
    Public endPoint As Integer = 0
    Public entryPointPacked As Integer = 0
    Public callsFrom As New List(Of String)
End Class

Public Class CallsFromTo
    Public fromAddress As Integer
    Public toAddress As Integer
End Class
Public Class InlineString
    Public text As String = ""
    Public size As Integer = 0
End Class

Public Class DictionaryEntry
    Public dictWord As String = ""
    Public dictAddress As Integer = 0
    Public Flags As Integer = 0
    Public V1 As Integer = 0
    Public V2 As Integer = 0
    Public VerbNum As Integer = 0
    Public AdjNum As Integer = 0
    Public PrepNum As Integer = 0
    Public DirNum As Integer = 0
    Public Byte6ToLast As Integer = 0
    Public Byte5ToLast As Integer = 0
    Public Byte4ToLast As Integer = 0
    Public Byte3ToLast As Integer = 0
    Public Byte2ToLast As Integer = 0
    Public Byte1ToLast As Integer = 0
    Public v2_ClassificationNumber As Integer = 0
    Public v2_SemanticStuff As Integer = 0
    Public v2_WordFlags As Integer = 0
End Class

Public Class DictionaryEntries
    Inherits List(Of DictionaryEntry)

    Public WordSize As Integer = 0
    Public v1_CompactVocabulary As Boolean = False
    Public v2_OneBytePartsOfSpeech As Boolean = False
    Public v2_WordFlagsInTable As Boolean = True
    Public v2_PartsOfSpeech As New Dictionary(Of Integer, String)
    Public v2_VerbAddresses As List(Of Integer)
    Public v2_VerbWord As Integer = 0
    Public v2_NounWord As Integer = 0
    Public v2_AdjWord As Integer = 0
    Public v2_AdvWord As Integer = 0
    Public v2_QuantWord As Integer = 0
    Public v2_MiscWord As Integer = 0
    Public v2_CommaWord As Integer = 0
    Public v2_ParticleWord As Integer = 0
    Public v2_PrepWord As Integer = 0
    Public v2_ToBeWord As Integer = 0
    Public v2_ApostrWord As Integer = 0
    Public v2_OfWord As Integer = 0
    Public v2_ArticleWord As Integer = 0
    Public v2_QuoteWord As Integer = 0
    Public v2_EOIWord As Integer = 0
    Public v2_DirWord As Integer = 0
    Public v2_CanDoWord As Integer = 0
    Public v2_QWord As Integer = 0
    Public v2_AskWord As Integer = 0

    Public ReadOnly Property DirectionCount() As Integer
        Get
            Dim oList As New List(Of Integer)
            For Each oEntry As DictionaryEntry In Me
                If oEntry.DirNum > 0 AndAlso Not oList.Contains(oEntry.DirNum) Then oList.Add(oEntry.DirNum)
            Next
            Return oList.Count
        End Get
    End Property

    Public ReadOnly Property PrepositionCountUnique() As Integer
        Get
            Dim oList As New List(Of Integer)
            For Each oEntry As DictionaryEntry In Me
                If oEntry.PrepNum > 0 AndAlso Not oList.Contains(oEntry.PrepNum) Then oList.Add(oEntry.PrepNum)
            Next
            Return oList.Count
        End Get
    End Property

    Public ReadOnly Property PrepositionCountTotal() As Integer
        Get
            Dim count As Integer = 0
            For Each oEntry As DictionaryEntry In Me
                If (oEntry.Flags And 8) = 8 Then
                    count += 1
                End If
            Next
            Return count
        End Get
    End Property

    Private _adjectiveCount As Integer = -1
    Public ReadOnly Property AdjectiveCount() As Integer
        ' Only for ZIL, version 1-3
        Get
            If _adjectiveCount > -1 Then Return _adjectiveCount
            Dim oList As New List(Of Integer)
            For Each oEntry As DictionaryEntry In Me
                If oEntry.AdjNum > 0 AndAlso Not oList.Contains(oEntry.AdjNum) Then oList.Add(oEntry.AdjNum)
            Next
            _adjectiveCount = oList.Count
            Return _adjectiveCount
        End Get
    End Property

    Public Function GetVerb(piVerbNum As Integer) As DictionaryEntries
        Dim listRet As New DictionaryEntries
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.VerbNum = piVerbNum Then listRet.Add(dictEntry)
        Next
        Return listRet
    End Function

    Public Function GetLowestVerbNum() As Integer
        Dim iLowestVerbNum = 256
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.VerbNum > 0 And dictEntry.VerbNum < iLowestVerbNum Then iLowestVerbNum = dictEntry.VerbNum
        Next
        Return iLowestVerbNum
    End Function

    Public Function GetAdjective(piAdjNum As Integer) As DictionaryEntries
        Dim listRet As New DictionaryEntries
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.AdjNum = piAdjNum Then listRet.Add(dictEntry)
        Next
        Return listRet
    End Function

    Public Function GetDirection(piDirNum As Integer) As DictionaryEntries
        Dim listRet As New DictionaryEntries
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.DirNum = piDirNum Then listRet.Add(dictEntry)
        Next
        Return listRet
    End Function

    Public Function GetPreposition(piPrepNum As Integer) As DictionaryEntries
        Dim listRet As New DictionaryEntries
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.PrepNum = piPrepNum Then listRet.Add(dictEntry)
        Next
        Return listRet
    End Function

    Public Function GetEntryAtAddress(piAddress As Integer) As DictionaryEntry
        For Each dictEntry As DictionaryEntry In Me
            If dictEntry.dictAddress = piAddress Then Return dictEntry
        Next
        Return Nothing
    End Function
End Class
