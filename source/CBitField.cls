VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBitField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'THIS CLASS
'Creates a 4 boolean arrays that can be converted to primitive data types by converting such arrays into bitfields.

Private Const BYTE_BITSIZE As Long = 8
Private Const WORD_BITSIZE As Long = 16
Private Const DWORD_BITSIZE As Long = 32
Private Const QWORD_BITSIZE As Long = 64

Private Const BYTE_SIZE As Long = 1
Private Const WORD_SIZE As Long = 2
Private Const DWORD_SIZE As Long = 4
Private Const QWORD_SIZE As Long = 8

Public Enum eTableSize
    TS_BYTE = 0
    TS_WORD = 1
    TS_DWORD = 2
    TS_QWORD = 3
End Enum

Private Const ASC_ZERO As Long = 48 'Asc("0")
Private Const ASC_ONE As Long = 49 'Asc("1")

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal ByteLen As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (ByRef Destination As Any, ByVal Length As Long)

Private p_FlagTable_Byte(0 To BYTE_BITSIZE - 1) As Boolean
Private p_FlagTable_Word(0 To WORD_BITSIZE - 1) As Boolean
Private p_FlagTable_DWord(0 To DWORD_BITSIZE - 1) As Boolean
Private p_FlagTable_QWord(0 To QWORD_BITSIZE - 1) As Boolean

Private Sub Class_Initialize()
    'Constructor
    
End Sub

Public Property Get flagTable(ByVal byteKey As eTableSize) As String
    Dim retVal As String
    
    On Error GoTo ErrHandler
    
    Select Case byteKey
    Case eTableSize.TS_BYTE
        retVal = flagTableToString(p_FlagTable_Byte)
    Case eTableSize.TS_WORD
        retVal = flagTableToString(p_FlagTable_Word)
    Case eTableSize.TS_DWORD
        retVal = flagTableToString(p_FlagTable_DWord)
    Case eTableSize.TS_QWORD
        retVal = flagTableToString(p_FlagTable_QWord)
    Case Else
        Call RaiseUserError(PERR_INVALIDBYTESIZE, GetErrorMessage(PERR_INVALIDBYTESIZE))
    End Select
    
    flagTable = retVal
    
Exit Property
ErrHandler:
    Call ErrorLevelManager(ELC_SOURCECHAIN, "FlagTable [Get]", Err.Description, Err.number, Err.Source)
End Property

Public Property Let flagTable(ByVal byteKey As eTableSize, ByVal newValue As String)
    On Error GoTo ErrHandler
    
    Call stringToFlagTable(byteKey, StrReverse(newValue))
    
Exit Property
ErrHandler:
    Call ErrorLevelManager(ELC_SOURCECHAIN, "FlagTable [Let]", Err.Description, Err.number, Err.Source)
End Property

Private Function flagTableToString(ByRef fTable() As Boolean) As String
    Dim upBound As Integer, lwBound As Integer
    Dim i As Long
    Dim strOut As String
    
    lwBound = LBound(fTable)
    upBound = UBound(fTable)
    
    For i = lwBound To upBound
        If fTable(i) Then
            strOut = strOut & "1"
        Else
            strOut = strOut & "0"
        End If
    Next i
    
    flagTableToString = StrReverse(strOut)
End Function

Private Sub stringToFlagTable(ByVal byteKey As eTableSize, ByVal strVal As String)
    'inside each select there is an if that asserts that:
    'The value being passed as parameter has the correct lenght, and that it contains only 0s or 1s
    'In case that the assertion is not true, an error will be raised
    
    Dim ucodeArray() As Byte
    Dim i As Long 'counter variable for For-Loops
    Dim ftIndex As Long 'flag table index
    
    On Error GoTo ErrHandler
    
    Select Case byteKey
    Case eTableSize.TS_BYTE
        If Len(strVal) <> BYTE_BITSIZE And Not (strVal Like nString(BYTE_BITSIZE, "[0-1]")) Then _
            Call RaiseUserError(PERR_BITMISTACH, GetErrorMessage(PERR_BITMISTACH))
        
        ucodeArray = strVal
        
        For i = 0 To BYTE_BITSIZE * 2 - 1 Step 2
            ftIndex = i \ 2
            If ucodeArray(i) = ASC_ZERO Then
                p_FlagTable_Byte(ftIndex) = False
            Else 'ucodeArray(i) = ASC_ONE
                p_FlagTable_Byte(ftIndex) = True
            End If
        Next i
        
    Case eTableSize.TS_WORD
        If Len(strVal) <> WORD_BITSIZE And Not (strVal Like nString(WORD_BITSIZE, "[0-1]")) Then _
            Call RaiseUserError(PERR_BITMISTACH, GetErrorMessage(PERR_BITMISTACH))
        
        ucodeArray = strVal
        
        For i = 0 To WORD_BITSIZE * 2 - 1 Step 2
            ftIndex = i \ 2
            If ucodeArray(i) = ASC_ZERO Then
                p_FlagTable_Word(ftIndex) = False
            Else 'ucodeArray(i) = ASC_ONE
                p_FlagTable_Word(ftIndex) = True
            End If
        Next i
        
    Case eTableSize.TS_DWORD
        If Len(strVal) <> DWORD_BITSIZE And Not (strVal Like nString(DWORD_BITSIZE, "[0-1]")) Then _
            Call RaiseUserError(PERR_BITMISTACH, GetErrorMessage(PERR_BITMISTACH))
        
        ucodeArray = strVal
        
        For i = 0 To DWORD_BITSIZE * 2 - 1 Step 2
            ftIndex = i \ 2
            If ucodeArray(i) = ASC_ZERO Then
                p_FlagTable_DWord(ftIndex) = False
            Else 'ucodeArray(i) = ASC_ONE
                p_FlagTable_DWord(ftIndex) = True
            End If
        Next i
        
    Case eTableSize.TS_QWORD
        If Len(strVal) <> QWORD_BITSIZE And Not (strVal Like nString(QWORD_BITSIZE, "[0-1]")) Then _
            Call RaiseUserError(PERR_BITMISTACH, GetErrorMessage(PERR_BITMISTACH))
        
        ucodeArray = strVal
        
        For i = 0 To QWORD_BITSIZE * 2 - 1 Step 2
            ftIndex = i \ 2
            If ucodeArray(i) = ASC_ZERO Then
                p_FlagTable_QWord(ftIndex) = False
            Else 'ucodeArray(i) = ASC_ONE
                p_FlagTable_QWord(ftIndex) = True
            End If
        Next i
        
    Case Else
        Call RaiseUserError(PERR_INVALIDBYTESIZE, GetErrorMessage(PERR_INVALIDBYTESIZE))
    End Select
    
Exit Sub
ErrHandler:
    Call ErrorLevelManager(ELC_SOURCECHAIN, "stringToFlagTable", Err.Description, Err.number, Err.Source)
End Sub

'Returns a string that contains seedStr as many times as determined by number
Private Function nString(ByVal number As Long, ByVal seedStr As String) As String
    Dim retVal As String, i As Long
    Dim chrLen As Long, bytLen As Long
    
    If number > 0 Then
        chrLen = Len(seedStr) 'Gets the number of characters in the string
        bytLen = LenB(seedStr) 'Gets the number of bytes needed to store the string
        
        'padds the return variable with as many blank spaces as needed to accomdate the output
        retVal = String(number * chrLen, " ")
        
        For i = 0 To number - 1
            Call CopyMemory(ByVal StrPtr(retVal) + i * bytLen, ByVal StrPtr(seedStr), bytLen)
        Next i
        
        nString = retVal
    ElseIf number = 0 Then
        nString = ""
    Else
        Err.Raise 5
    End If
End Function

Private Function getBitField(ByVal sizeCode As eTableSize) As Byte()
    Dim retVal() As Byte 'array of bytes that will contain the bitfield
    Dim i As Long 'counter variable for For-Loops
    Dim bfIndex As Long 'bit field index
    Dim bitMask As Byte 'or mask to aggregate values
    
    'Selects the correct Flag Table to create an appropriately sized bit field array
    Select Case sizeCode
    Case eTableSize.TS_BYTE
        ReDim retVal(0 To BYTE_SIZE - 1) 'allocates storage for a byte-sized bitfield (8 bits)
        
        bfIndex = 0
        'goes through the flag table to calculate the bits
        For i = 0 To BYTE_BITSIZE - 1
            'if flag is true add bit to array
            If p_FlagTable_Byte(i) Then
                bitMask = 2 ^ (i Mod BYTE_BITSIZE)
                retVal(bfIndex) = retVal(bfIndex) Or bitMask
            End If
        Next i
        
    Case eTableSize.TS_WORD
        ReDim retVal(0 To WORD_SIZE - 1) 'allocates storage for a word-sized bitfield (16 bits)
        
        'goes through the flag table to calculate the bits
        For i = 0 To WORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE 'ensures that the bits fall in the right index
            
            'if flag is true add bit to array
            If p_FlagTable_Word(i) Then
                bitMask = 2 ^ (i Mod BYTE_BITSIZE)
                retVal(bfIndex) = retVal(bfIndex) Or bitMask
            End If
        Next i
        
    Case eTableSize.TS_DWORD
        ReDim retVal(0 To DWORD_SIZE - 1) 'allocates storage for a double word-sized bitfield (32 bits)
        
        'goes through the flag table to calculate the bits
        For i = 0 To DWORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE 'ensures that the bits fall in the right index
            
            'if flag is true add bit to array
            If p_FlagTable_DWord(i) Then
                bitMask = 2 ^ (i Mod BYTE_BITSIZE)
                retVal(bfIndex) = retVal(bfIndex) Or bitMask
            End If
        Next i
        
    Case eTableSize.TS_QWORD
        ReDim retVal(0 To QWORD_SIZE - 1) 'allocates storage for a quad word-sized bitfield (64 bits)
        
        'goes through the flag table to calculate the bits
        For i = 0 To QWORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE 'ensures that the bits fall in the right index
            
            'if flag is true add bit to array
            If p_FlagTable_QWord(i) Then
                bitMask = 2 ^ (i Mod BYTE_BITSIZE)
                retVal(bfIndex) = retVal(bfIndex) Or bitMask
            End If
        Next i
        
    End Select
    
    'returns correctly filled bitfield array
    getBitField = retVal
End Function

Private Sub setFlagTable(ByRef bitField() As Byte, ByVal byteKey As eTableSize)
    Dim i As Long 'counter variable for For-Loops
    Dim bfIndex As Long 'flag table index
    Dim bitMask As Byte 'and mask to find true flags
    
    Select Case byteKey
    Case eTableSize.TS_BYTE
        bfIndex = 0
        For i = 0 To BYTE_BITSIZE - 1
            'bfIndex = i \ BYTE_BITSIZE
            bitMask = 2 ^ (i Mod BYTE_BITSIZE) '&H80& \ 2 ^ i
            p_FlagTable_Byte(i) = bitField(bfIndex) And bitMask
        Next i
        
    Case eTableSize.TS_WORD
        For i = 0 To WORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE
            bitMask = 2 ^ (i Mod BYTE_BITSIZE) '&H80& \ 2 ^ i
            p_FlagTable_Word(i) = bitField(bfIndex) And bitMask
        Next i
        
        
    Case eTableSize.TS_DWORD
        For i = 0 To DWORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE
            bitMask = 2 ^ (i Mod BYTE_BITSIZE) '&H80& \ 2 ^ i
            p_FlagTable_DWord(i) = bitField(bfIndex) And bitMask
        Next i
        
        
    Case eTableSize.TS_QWORD
        For i = 0 To QWORD_BITSIZE - 1
            bfIndex = i \ BYTE_BITSIZE
            bitMask = 2 ^ (i Mod BYTE_BITSIZE) '&H80& \ 2 ^ i
            p_FlagTable_QWord(i) = bitField(bfIndex) And bitMask
        Next i
        
    End Select
End Sub

Public Property Get b_Byte() As Byte
    Dim retVal As Byte, bitField() As Byte
    
    bitField = getBitField(TS_BYTE)
    
    CopyMemory retVal, bitField(0), BYTE_SIZE
    b_Byte = retVal
End Property

Public Property Let b_Byte(ByVal newNumber As Byte)
    Dim bitField(0 To BYTE_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, BYTE_SIZE
    Call setFlagTable(bitField, TS_BYTE)
End Property

Public Property Get w_Integer() As Integer
    Dim retVal As Integer, bitField() As Byte
    
    bitField = getBitField(TS_WORD)
    
    CopyMemory retVal, bitField(0), WORD_SIZE
    w_Integer = retVal
End Property

Public Property Let w_Integer(ByVal newNumber As Integer)
    Dim bitField(0 To WORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, WORD_SIZE
    Call setFlagTable(bitField, TS_WORD)
End Property

Public Property Get w_Boolean() As Boolean
    Dim retVal As Boolean, bitField() As Byte
    
    bitField = getBitField(TS_WORD)
    
    CopyMemory retVal, bitField(0), WORD_SIZE
    w_Boolean = retVal
End Property

Public Property Let w_Boolean(ByVal newNumber As Boolean)
    Dim bitField(0 To WORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, WORD_SIZE
    Call setFlagTable(bitField, TS_WORD)
End Property

Public Property Get dw_Long() As Long
    Dim retVal As Long, bitField() As Byte
    
    bitField = getBitField(TS_DWORD)
    
    CopyMemory retVal, bitField(0), DWORD_SIZE
    dw_Long = retVal
End Property

Public Property Let dw_Long(ByRef newNumber As Long)
    Dim bitField(0 To DWORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, DWORD_SIZE
    Call setFlagTable(bitField, TS_DWORD)
End Property

Public Property Get dw_Single() As Single
    Dim retVal As Single, bitField() As Byte
    
    bitField = getBitField(TS_DWORD)
    
    CopyMemory retVal, bitField(0), DWORD_SIZE
    dw_Single = retVal
End Property

Public Property Let dw_Single(ByRef newNumber As Single)
    Dim bitField(0 To DWORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, DWORD_SIZE
    Call setFlagTable(bitField, TS_DWORD)
End Property

Public Property Get qw_Currency() As Currency
    Dim retVal As Currency, bitField() As Byte
    
    bitField = getBitField(TS_QWORD)
    
    CopyMemory retVal, bitField(0), QWORD_SIZE
    qw_Currency = retVal
End Property

Public Property Let qw_Currency(ByRef newNumber As Currency)
    Dim bitField(0 To QWORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, QWORD_SIZE
    Call setFlagTable(bitField, TS_QWORD)
End Property

Public Property Get qw_Double() As Double
    Dim retVal As Double, bitField() As Byte
    
    bitField = getBitField(TS_QWORD)
    
    CopyMemory retVal, bitField(0), QWORD_SIZE
    qw_Double = retVal
End Property

Public Property Let qw_Double(ByRef newNumber As Double)
    Dim bitField(0 To QWORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, QWORD_SIZE
    Call setFlagTable(bitField, TS_QWORD)
End Property

Public Property Get qw_Date() As Date
    Dim retVal As Double, bitField() As Byte
    
    bitField = getBitField(TS_QWORD)
    
    CopyMemory retVal, bitField(0), QWORD_SIZE
    qw_Date = retVal
End Property

Public Property Let qw_Date(ByRef newNumber As Date)
    Dim bitField(0 To QWORD_SIZE - 1) As Byte
    
    CopyMemory bitField(0), newNumber, QWORD_SIZE
    Call setFlagTable(bitField, TS_QWORD)
End Property
