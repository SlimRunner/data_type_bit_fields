Attribute VB_Name = "mdlPrivateErrorDispatch"
Option Explicit

Public Enum enumCBitFieldError
    PERR_GENERIC = 0
    PERR_INVALIDBYTESIZE = 101
    PERR_BITMISTACH = 110
End Enum

Public Function GetErrorMessage(ByVal errNum As enumCBitFieldError) As String
    Select Case errNum
    Case enumCBitFieldError.PERR_GENERIC
        GetErrorMessage = "Generic error"
    Case enumCBitFieldError.PERR_INVALIDBYTESIZE
        GetErrorMessage = "That is not a compatible type for a variable input or output."
    Case enumCBitFieldError.PERR_BITMISTACH
        GetErrorMessage = "The string provided doesn't match the length of the bitfield being assigned."
    End Select
End Function
