clear classes
clc

Tc3_VS = VisualStudioHandler;

% scan system for installed VS versions
VsV = Tc3_VS.GetInstalledVisualStudios;

% new VS instance (use latest VS)
Tc3_VS.CreateNewVisualStudioInstance(VsV(end), true);

% new Tc Solution  \Test/CreateNewVisualStudioInstance.sln
% if allready exist, delete folder 
if exist('Test', 'dir') == 7
   rmdir Test s % delete incl. subfolders
end
Tc3_VS.CreateTwinCatSolution('C:\Users\chrisk\Desktop\Test', 'RealTimeTempControl');

sysMan = Tc3_VS.sysManager;

Tc3_Tc = TwinCATHandler(sysMan);

Tc3_Tc.CreateTask();
