clear classes
clc

Tc3_VS = VisualStudioHandler;

VsV = Tc3_VS.GetInstalledVisualStudios;

Tc3_VS.CreateNewVisualStudioInstance(VsV{end}, false);

Tc3_VS.CreateTwinCatSolution('C:\Users\chrisk\Desktop\Test','bob');

Tc3_IO = IOHandler(Tc3_VS.sysManager);

%Tc3_IO.AddCanInterface();

Tc3_IO.CreateEtherCAT();

%Tc3_IO.AddEcSlave();

%Tc3_VS.SaveSolution();
