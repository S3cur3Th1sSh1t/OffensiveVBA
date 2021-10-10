# OffensiveVBA

In preparation for a VBS AV Evasion Stream/Video I was doing some research for Office Macro code execution methods.

The list got longer and longer and I found no central place for offensive VBA templates - so this repo can be used for such.

## Examples in this repo

| File | Description |
| ---  | --- |
| [ShellApplication_ShellExecute.vba](../master/src/ShellApplication_ShellExecute.vba) | Execute an OS command via ShellApplication object and ShellExecute method |
| [ShellApplication_ShellExecute_privileged.vba](../master/src/ShellApplication_ShellExecute_privileged.vba) | Execute an privileged OS command via ShellApplication object and ShellExecute method - UAC prompt |
| [Shellcode_CreateThread.vba](../master/src/Shellcode_CreateThread.vba) | Execute shellcode in the current process via Win32 CreateThread |
| [Shellcode_EnumChildWindowsCallback.vba](../master/src/Shellcode_EnumChildWindowsCallback.vba) | Execute shellcode in the current process via EnumChildWindows |
| [Win32_CreateProcess.vba](../master/src/Win32_CreateProcess.vba) | Create a new process for code execution via Win32 CreateProcess function |
| [Win32_ShellExecute.vba](../master/src/Win32_ShellExecute.vba) | Create a new process for code execution via Win32 ShellExecute function |
| [WMI_Process_Create.vba](../master/src/WMI_Process_Create.vba) | Create a new process via WMI for code execution |
| [WMI_Process_Create2.vba](../master/src/WMI_Process_Create2.vba) | Another WMI code execution example |
| [WscriptShell_Exec.vba](../master/src/WscriptShell_Exec.vba) | Execute an OS command via WscriptShell object and Exec method |
| [WscriptShell_run.vba](../master/src/WscriptShell_run.vba) | Execute an OS command via WscriptShell object and Run method |
| [VBA-RunPE](../master/src/VBA-RunPE/) | [@itm4n's](https://twitter.com/itm4n) RunPE technique in VBA |
| [GadgetToJScript](../master/src/GadgetToJScript/) | [med0x2e's](https://github.com/med0x2e) C# script for generating .NET serialized gadgets that can trigger .NET assembly load/execution when deserialized using BinaryFormatter from JS/VBS/VBA based scripts.  |
| [PPID_Spoof.vba](../master/src/PPID_Spoof.vba) | [christophetd's](https://github.com/christophetd)  [spoofing-office-macro](https://github.com/christophetd/spoofing-office-macro) copy |
| [AMSIBypass_AmsiScanBuffer_Patch.vba](../master/src/AMSIBypass_AmsiScanBuffer_Patch.vba) | [rmdavy's](https://github.com/rmdavy) AMSI Bypass to patch AmsiScanBuffer using ordinal values for a signature bypass |
| [AMSIBypass_Heap.vba](../master/src/AMSIBypass_Heap.vba) | [rmdavy's](https://github.com/rmdavy) [HeapsOfFun](https://github.com/rmdavy/HeapsOfFun) repo copy  |
| [AMSIbypasses.vba](../master/src/AMSIbypasses.vba) | [outflanknl's](https://github.com/outflanknl) [AMSI bypass blog](https://outflank.nl/blog/2019/04/17/bypassing-amsi-for-vba/) |

## Obfuscators / Payload generators

1) [VBad](https://github.com/Pepitoh/VBad)
2) [wePWNise](https://github.com/FSecureLABS/wePWNise)
3) [wePWNise](https://github.com/FSecureLABS/wePWNise)
4) [VisualBasicObfuscator](https://github.com/mgeeky/VisualBasicObfuscator/tree/master)
5) [macro_pack](https://github.com/sevagas/macro_pack)
6) [shellcode2vbscript.py](https://github.com/DidierStevens/DidierStevensSuite/blob/master/shellcode2vbscript.py)
7) [EvilClippy](https://github.com/outflanknl/EvilClippy)
8) [OfficePurge](https://github.com/mandiant/OfficePurge)


