# OffensiveVBA

In preparation for a VBS AV Evasion Stream/Video I was doing some research for Office Macro code execution methods and evasion techniques.

The list got longer and longer and I found no central place for offensive VBA templates - so this repo can be used for such.

## Templates in this repo

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
| [AMSIBypass_AmsiScanBuffer_ordinal.vba](../master/src/AMSIBypass_AmsiScanBuffer_ordinal.vba) | [rmdavy's](https://github.com/rmdavy) AMSI Bypass to patch AmsiScanBuffer using ordinal values for a signature bypass |
| [AMSIBypass_AmsiScanBuffer_Classic.vba](../master/src/AMSIBypass_AmsiScanBuffer_Classic.vba) | [rasta-mouse's](https://github.com/rasta-mouse) classic AmsiScanBuffer patch |
| [AMSIBypass_Heap.vba](../master/src/AMSIBypass_Heap.vba) | [rmdavy's](https://github.com/rmdavy) [HeapsOfFun](https://github.com/rmdavy/HeapsOfFun) repo copy  |
| [AMSIbypasses.vba](../master/src/AMSIbypasses.vba) | [outflanknl's](https://github.com/outflanknl) [AMSI bypass blog](https://outflank.nl/blog/2019/04/17/bypassing-amsi-for-vba/) |
| [COMHijack_DLL_Load.vba](../master/src/COMHijack_DLL_Load.vba) | Load DLL via COM Hijacking |
| [COM_Process_create.vba](../master/src/COM_Process_create.vba) | Create process via COM object |
| [Download_Autostart.vba](../master/src/Download_Autostart.vba) | Download a file from a remote webserver and put it into the StartUp folder |
| [Download_Autostart_WinAPI.vba](../master/src/Download_Autostart_WinAPI.vba) | Download a file from a remote webserver via URLDownloadtoFileA and put it into the StartUp folder |
| [Dropper_Autostart.vba](../master/src/Dropper_Autostart.vba) | Drop batch file into the StartUp folder |
| [Dropper_Embeded_msf.vba](../master/src/Dropper_Embeded_msf.vba) | MSF generated dropper template |
| [Registry_Persist_wmi.vba](../master/src/Registry_Persist_wmi.vba) | Create StartUp registry key for persistence via WMI |
| [Registry_Persist_wscript.vba](../master/src/Registry_Persist_wscript.vba) | Create StartUp registry key for persistence via wscript object |
| [ScheduledTask_Create.vba](../master/src/ScheduledTask_Create.vba) | Create and start sheduled task for code execution/persistence |
| [XMLDOM_Load_XSL_Process_create.vba](../master/src/XMLDOM_Load_XSL_Process_create.vba) | Load XSL from a remote webserver to execute code |
| [regsvr32_sct_DownloadExecute.vba](../master/src/regsvr32_sct_DownloadExecute.vba) | Execute regsvr32 to download a remote webservers SCT file for code execution |
| [BlockETW.vba](../master/src/BlockETW.vba) | Patch EtwEventWrite in ntdll.dll to block ETW data collection |
| [ShellWindows_Process_create.vba](../master/src/ShellWindows_Process_create.vba) | ShellWindows Process create to get explorer.exe as parent process |
| [AES.vba](../master/src/AES.vba) | An example to use AES encryption/decryption in VBA from [Here](https://github.com/susam/aes.vbs/blob/a0cb5f9ffbd90b435622f5cfdb84264e1a319bf2/aes.vbs) |



## Obfuscators / Payload generators

1) [VBad](https://github.com/Pepitoh/VBad)
2) [wePWNise](https://github.com/FSecureLABS/wePWNise)
3) [VisualBasicObfuscator](https://github.com/mgeeky/VisualBasicObfuscator/tree/master)
4) [macro_pack](https://github.com/sevagas/macro_pack)
5) [shellcode2vbscript.py](https://github.com/DidierStevens/DidierStevensSuite/blob/master/shellcode2vbscript.py)
6) [EvilClippy](https://github.com/outflanknl/EvilClippy)
7) [OfficePurge](https://github.com/mandiant/OfficePurge)
8) [SharpShooter](https://github.com/mdsecactivebreach/SharpShooter)
9) [VBS-Obfuscator-in-Python](https://github.com/kkar/VBS-Obfuscator-in-Python)

## Credits / usefull resources

ASR bypass:
http://blog.sevagas.com/IMG/pdf/bypass_windows_defender_attack_surface_reduction.pdf

Shellcode to VBScript conversion:
https://github.com/DidierStevens/DidierStevensSuite/blob/master/shellcode2vbscript.py

Bypass AMSI in VBA:
https://outflank.nl/blog/2019/04/17/bypassing-amsi-for-vba/usercontent.com/27858067/138140466-5bcac24c-e3db-47f6-a38b-6eb5d7605fa0.png)





