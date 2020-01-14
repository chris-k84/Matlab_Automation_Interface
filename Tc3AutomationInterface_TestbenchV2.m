%% Test Tc3 Automation Interface Wrapper
% 
% This m.-File uses the Tc3AutomationInterface-wrapper class to
% automatically setup the temperature controller sample including automatic
% start on the local TC3 runtime
% Sample Code provided by Beckhoff Automation GmbH & Co. KG


% clear workspace
clear classes
clc

% Call constructor
Tc3_AI = Tc3AutomationInterfaceV2;



% scan system for installed VS versions
VsV = Tc3_AI.GetInstalledVisualStudios;

% new VS instance (use latest VS)
Tc3_AI.CreateNewVisualStudioInstance(VsV(end), true);

% new Tc Solution  \Test/CreateNewVisualStudioInstance.sln
% if allready exist, delete folder 
if exist('Test', 'dir') == 7
   rmdir Test s % delete incl. subfolders
end
Tc3_AI.CreateTwinCatSolution([cd,'\Test'], 'RealTimeTempControl');

% Option: Load a prepared TwinCAT Solution 
% Tc3_AI.OpenTwinCatSolution([cd,'\Test'], 'RealTimeTempControl')

% Option: scan for boxes at remote target
%AmsNetId = '172.17.36.3.1.1';
%AmsNetId = '169.254.45.114.1.1'; %Xts machine
AmsNetId = '5.52.64.65.1.1'; %demo rig
%Boxes = Tc3_AI.ScanEtherCatMasters(AmsNetId);

% new Object of TctSmplTempCtrl
%Tc3_AI.CreateTcCOM('TctSmplTempCtrl_'); 
% new Object of TctSmplCtrlSysPT2
%Tc3_AI.CreateTcCOM('TctSmplCtrlSysPT2_');

% Linking inputs and outputs of Temp Control setup
%Tc3_AI.LinkVariables([char(Tc3_AI.TcCOM{1}.PathName),'^Input^FeedbackTemp'], [char(Tc3_AI.TcCOM{2}.PathName),'^Output^Temp']);
%Tc3_AI.LinkVariables([char(Tc3_AI.TcCOM{1}.PathName),'^Output^HeaterOn'], [char(Tc3_AI.TcCOM{2}.PathName),'^Input^Heat_On']);

% save
%Tc3_AI.SaveSolution();

% activate at local XAR
%Tc3_AI.ActivateOnDevice('Local');