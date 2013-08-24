Attribute VB_Name = "bytes"
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

 '*----------------------------------------------------------*
 '* Name       : MAKELONG                                    *
 '*----------------------------------------------------------*
 '* Purpose    : Combines two integers into a long integer.  *
 '*----------------------------------------------------------*
 '* Parameters : wLow   Required. Low WORD.                  *
 '*            : wHigh  Required. High WORD.                 *
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language MAKELONG macro.                    *
 '*----------------------------------------------------------*
 Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
   MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
 End Function

 '*----------------------------------------------------------*
 '* Name       : MAKELPARAM                                  *
 '*----------------------------------------------------------*
 '* Purpose    : Combines two integers into a long integer.  *
 '*----------------------------------------------------------*
 '* Parameters : wLow   Required. Low WORD.                  *
 '*            : wHigh  Required. High WORD.                 *
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language MAKELPARAM macro.                  *
 '*----------------------------------------------------------*
 Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
   MAKELPARAM = MAKELONG(wLow, wHigh)
 End Function

 '*----------------------------------------------------------*
 '* Name       : MAKEWORD                                    *
 '*----------------------------------------------------------*
 '* Purpose    : Combines two integers into a 16-bit unsigned*
 '*            : integer (word).                             *
 '*----------------------------------------------------------*
 '* Parameters : wLow   Required. Low BYTE.                  *
 '*            : wHigh  Required. High BYTE.                 *
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language MAKELONG macro.                    *
 '*----------------------------------------------------------*
 Public Function MAKEWORD(wLow As Long, wHigh As Long) As Long
   MAKEWORD = LOBYTE(wLow) Or (&H100& * LOBYTE(wHigh))
 End Function

 '*----------------------------------------------------------*
 '* Name       : LOWORD                                      *
 '*----------------------------------------------------------*
 '* Purpose    : Returns the low 16-bit integer from a 32-bit*
 '*            : long integer.                               *
 '*----------------------------------------------------------*
 '* Parameters : dwValue Required. 32-bit long integer value.*
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language LOWORD macro.                      *
 '*----------------------------------------------------------*
 Public Function LOWORD(dwValue As Long) As Integer
   MoveMemory LOWORD, dwValue, 2
 End Function

 '*----------------------------------------------------------*
 '* Name       : HIWORD                                      *
 '*----------------------------------------------------------*
 '* Purpose    : Returns the high 16-bit integer from a      *
 '*            : 32-bit long integer.                        *
 '*----------------------------------------------------------*
 '* Parameters : dwValue Required. 32-bit long integer value.*
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language HIWORD macro.                      *
 '*----------------------------------------------------------*
 Public Function HIWORD(dwValue As Long) As Integer
   MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
 End Function

 '*----------------------------------------------------------*
 '* Name       : LOBYTE                                      *
 '*----------------------------------------------------------*
 '* Purpose    : Returns the low 8-bit byte from a low word  *
 '*            : of 32-bit long integer.                     *
 '*----------------------------------------------------------*
 '* Parameters : dwValue Required. 32-bit long integer value.*
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language LOBYTE macro.                      *
 '*----------------------------------------------------------*
 Public Function LOBYTE(dwValue As Long) As Byte
   MoveMemory LOBYTE, LOWORD(dwValue), 1
 End Function

 '*----------------------------------------------------------*
 '* Name       : HIBYTE                                      *
 '*----------------------------------------------------------*
 '* Purpose    : Returns the high 8-bit byte from a low word *
 '*            : of 32-bit long integer.                     *
 '*----------------------------------------------------------*
 '* Parameters : dwValue Required. 32-bit long integer value.*
 '*----------------------------------------------------------*
 '* Description: This function is equivalent to the 'C'      *
 '*            : language HIBYTE macro.                      *
 '*----------------------------------------------------------*
 Public Function HIBYTE(dwValue As Long) As Byte
   MoveMemory HIBYTE, ByVal VarPtr(LOWORD(dwValue)) + 1, 1
 End Function

