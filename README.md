# OffensiveVBA

In preparation for a VBS AV Evasion Stream/Video I was doing some research for Office Macro code execution methods and evasion techniques.

The list got longer and longer and I found no central place for offensive VBA templates - so this repo can be used for such. It is very far away from being complete. If you know any other cool technique or useful template feel free to contribute and create a pull request!

Most of the templates in this repo were already published somewhere. I just copy pasted most templates from ms-docs sites, blog posts or from other tools.

## Templates in this repo

| File | Description |
| ---  | --- |
| [ShellApplication_ShellExecute.vba](./src/ShellApplication_ShellExecute.vba) | Execute an OS command via ShellApplication object and ShellExecute method |
| [ShellApplication_ShellExecute_privileged.vba](./src/ShellApplication_ShellExecute_privileged.vba) | Execute an privileged OS command via ShellApplication object and ShellExecute method - UAC prompt |
| [Shellcode_CreateThread.vba](./src/Shellcode_CreateThread.vba) | Execute shellcode in the current process via Win32 CreateThread |
| [Shellcode_EnumChildWindowsCallback.vba](./src/Shellcode_EnumChildWindowsCallback.vba) | Execute shellcode in the current process via EnumChildWindows |
| [Win32_CreateProcess.vba](./src/Win32_CreateProcess.vba) | Create a new process for code execution via Win32 CreateProcess function |
| [Win32_ShellExecute.vba](./src/Win32_ShellExecute.vba) | Create a new process for code execution via Win32 ShellExecute function |
| [WMI_Process_Create.vba](./src/WMI_Process_Create.vba) | Create a new process via WMI for code execution |
| [WMI_Process_Create2.vba](./src/WMI_Process_Create2.vba) | Another WMI code execution example |
| [WscriptShell_Exec.vba](./src/WscriptShell_Exec.vba) | Execute an OS command via WscriptShell object and Exec method |
| [WscriptShell_run.vba](./src/WscriptShell_run.vba) | Execute an OS command via WscriptShell object and Run method |
| [VBA-RunPE](./src/VBA-RunPE/) | [@itm4n's](https://twitter.com/itm4n) RunPE technique in VBA |
| [GadgetToJScript](./src/GadgetToJScript/) | [med0x2e's](https://github.com/med0x2e) C# script for generating .NET serialized gadgets that can trigger .NET assembly load/execution when deserialized using BinaryFormatter from JS/VBS/VBA based scripts.  |
| [PPID_Spoof.vba](./src/PPID_Spoof.vba) | [christophetd's](https://github.com/christophetd)  [spoofing-office-macro](https://github.com/christophetd/spoofing-office-macro) copy |
| [AMSIBypass_AmsiScanBuffer_ordinal.vba](./src/AMSIBypass_AmsiScanBuffer_ordinal.vba) | [rmdavy's](https://github.com/rmdavy) AMSI Bypass to patch AmsiScanBuffer using ordinal values for a signature bypass |
| [AMSIBypass_AmsiScanBuffer_Classic.vba](./src/AMSIBypass_AmsiScanBuffer_Classic.vba) | [rasta-mouse's](https://github.com/rasta-mouse) classic AmsiScanBuffer patch |
| [AMSIBypass_Heap.vba](./src/AMSIBypass_Heap.vba) | [rmdavy's](https://github.com/rmdavy) [HeapsOfFun](https://github.com/rmdavy/HeapsOfFun) repo copy  |
| [AMSIbypasses.vba](./src/AMSIbypasses.vba) | [outflanknl's](https://github.com/outflanknl) [AMSI bypass blog](https://outflank.nl/blog/2019/04/17/bypassing-amsi-for-vba/) |
| [COMHijack_DLL_Load.vba](./src/COMHijack_DLL_Load.vba) | Load DLL via COM Hijacking |
| [COM_Process_create.vba](./src/COM_Process_create.vba) | Create process via COM object |
| [Download_Autostart.vba](./src/Download_Autostart.vba) | Download a file from a remote webserver and put it into the StartUp folder |
| [Download_Autostart_WinAPI.vba](./src/Download_Autostart_WinAPI.vba) | Download a file from a remote webserver via URLDownloadtoFileA and put it into the StartUp folder |
| [Dropper_Autostart.vba](./src/Dropper_Autostart.vba) | Drop batch file into the StartUp folder |
| [Registry_Persist_wmi.vba](./src/Registry_Persist_wmi.vba) | Create StartUp registry key for persistence via WMI |
| [Registry_Persist_wscript.vba](./src/Registry_Persist_wscript.vba) | Create StartUp registry key for persistence via wscript object |
| [ScheduledTask_Create.vba](./src/ScheduledTask_Create.vba) | Create and start sheduled task for code execution/persistence |
| [XMLDOM_Load_XSL_Process_create.vba](./src/XMLDOM_Load_XSL_Process_create.vba) | Load XSL from a remote webserver to execute code |
| [regsvr32_sct_DownloadExecute.vba](./src/regsvr32_sct_DownloadExecute.vba) | Execute regsvr32 to download a remote webservers SCT file for code execution |
| [BlockETW.vba](./src/BlockETW.vba) | Patch EtwEventWrite in ntdll.dll to block ETW data collection |
| [BlockETW_COMPLUS_ETWEnabled_ENV.vba](./src/BlockETW_COMPLUS_ETWEnabled_ENV.vba) | Block ETW data collection by setting the environment variable COMPLUS_ETWEnabled to 0, credit to [@_xpn_](https://twitter.com/_xpn_) |
| [ShellWindows_Process_create.vba](./src/ShellWindows_Process_create.vba) | ShellWindows Process create to get explorer.exe as parent process |
| [AES.vba](./src/AES.vba) | An example to use AES encryption/decryption in VBA from [Here](https://github.com/susam/aes.vbs/blob/a0cb5f9ffbd90b435622f5cfdb84264e1a319bf2/aes.vbs) |
| [Dropper_Executable_Autostart.vba](./src/Dropper_Executable_Autostart.vba) | Get executable bytes from VBA and drop into Autostart - no download in this case |
| [MarauderDrop.vba](./src/MarauderDrop.vba) | Drop a COM registered .NET DLL into temp, import the function and execute code - in this case loads a remote C# binary from a webserver to memory and executes it - credit to [@Jean_Maes_1994](https://twitter.com/Jean_Maes_1994) for [MaraudersMap](https://github.com/NVISOsecurity/blogposts/tree/master/MaraudersMap) |
| [Dropper_Workfolders_lolbas_Execute.vba](./src/Dropper_Workfolders_lolbas_Execute.vba) | Drop an embedded executable into the TEMP directory and execute it using C:\windows\system32\Workfolders.exe as LOLBAS - credit to [@YoSignals](https://www.ctus.io/2021/04/12/exploading/) |
| [SandBoxEvasion](./src/SandBoxEvasion/) | Some SandBox Evasion templates |
| [Evasion Dropper Autostart.vba](./src/Evasion_Dropper_Autostart.vba)| Drops a file to the Startup directory bypassing file write monitoring via renamed folder operation|
|[Evasion MsiInstallProduct.vba](./src/Evasion%20MsiInstallProduct.vba)| Installs a remote MSI package using WindowsInstaller ActiveXObject avoiding spawning suspicious office child process, the msi installation will be executed as a child of the `MSIEXEC /V service`|
|[StealNetNTLMv2.vba](./src/StealNetNTLMv2.vba)| Steal NetNTLMv2 Hash via share connection - credit to [https://book.hacktricks.xyz/windows/ntlm/places-to-steal-ntlm-creds](https://book.hacktricks.xyz/windows/ntlm/places-to-steal-ntlm-creds)|
|[Parse-Outlook.vba](./src/Parse-Outlook.vba)| Parses Outlook for sensitive keywords and file extensions, and exfils them via email - credit to [JohnWoodman](https://github.com/JohnWoodman/VBA-Macro-Projects)|
|[Reverse-Shell.vba](./src/Reverse-Shell.vba)| Reverse shell written entirely in VBA using Windows API calls - credit to [JohnWoodman](https://github.com/JohnWoodman/VBA-Macro-Projects)|



## Missing - ToDos
| File | Description |
| ---  | --- |
| [Unhooker.vba](./src/Unhooker.vba) | Unhook API's in memory to get rid of hooks |
| [Syscalls.vba](./src/Syscalls.vba) | Syscall usage - fresh from disk or Syswhispers like |
| [Manymore.vba](./src/Manymore.vba) | If you have any more ideas feel free to contribute |


## Obfuscators / Payload generators

1) [VBad](https://github.com/Pepitoh/VBad)
2) [wePWNise](https://github.com/FSecureLABS/wePWNise)
3) [VisualBasicObfuscator](https://github.com/mgeeky/VisualBasicObfuscator/tree/master) - needs some modification as it doesn't split up lines and is therefore not usable for office document macros
4) [macro_pack](https://github.com/sevagas/macro_pack)
5) [shellcode2vbscript.py](https://github.com/DidierStevens/DidierStevensSuite/blob/master/shellcode2vbscript.py)
6) [EvilClippy](https://github.com/outflanknl/EvilClippy)
7) [OfficePurge](https://github.com/mandiant/OfficePurge)
8) [SharpShooter](https://github.com/mdsecactivebreach/SharpShooter)
9) [VBS-Obfuscator-in-Python](https://github.com/kkar/VBS-Obfuscator-in-Python) - - needs some modification as it doesn't split up lines and is therefore not usable for office document macros

## Credits / usefull resources

ASR bypass:
http://blog.sevagas.com/IMG/pdf/bypass_windows_defender_attack_surface_reduction.pdf

Shellcode to VBScript conversion:
https://github.com/DidierStevens/DidierStevensSuite/blob/master/shellcode2vbscript.py

Bypass AMSI in VBA:
https://outflank.nl/blog/2019/04/17/bypassing-amsi-for-vba/

VBA purging:
https://www.mandiant.com/resources/purgalicious-vba-macro-obfuscation-with-vba-purging

F-Secure VBA Evasion and detection post:
https://blog.f-secure.com/dechaining-macros-and-evading-edr/

One more F-Secure blog:
https://labs.f-secure.com/archive/dll-tricks-with-vba-to-improve-offensive-macro-capability/
