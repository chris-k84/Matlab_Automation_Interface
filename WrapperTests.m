clear classes
clc

Tc3_VS = VisualStudioHandler;

VsV = Tc3_VS.GetInstalledVisualStudios;

Tc3_VS.CreateNewVisualStudioInstance(VsV(end), true);

if exist('Test', 'dir') == 7
   rmdir Test s % delete incl. subfolders
end
Tc3_VS.CreateTwinCatSolution('C:\Users\chrisk\Desktop\Test', 'RealTimeTempControl');

sysMan = Tc3_VS.sysManager;

Tc3_Tc = TwinCATHandler(sysMan);

Tc3_Tc.CreateTask();

Tc3_Plc = PLCHandler(sysMan);

Tc3_Plc.CreatePLC("Bob");