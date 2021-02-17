@ECHO OFF
@ECHO REGISTRANDO DLL ERPSPSHEETS...

C:

CD \
CD C:\Windows\Microsoft.NET\Framework\v4.0.30319\

RegAsm.exe "E:\Desenvolvimento\Projetos\ErpSpSheets\ErpSpSheets\bin\Release\ErpSpSheets.dll" /tlb "E:\Desenvolvimento\Projetos\ErpSpSheets\ErpSpSheets\bin\Release\ErpSpSheets.tlb"

pause