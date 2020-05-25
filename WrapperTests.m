clear classes
clc

Tc3_VS = VisualStudioHandler;

VsV = Tc3_VS.GetInstalledVisualStudios;

Tc3_VS.CreateNewVisualStudioInstance(VsV{end}, true);
