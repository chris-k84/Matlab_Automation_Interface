classdef Tc3AutomationInterfaceV2 < handle
% Tc3AutomationInterface  wrapper class for the Tc3 Automation Interface.
%
%   Sample code to show how a wrapper class for the Automation Interface
%   can look like. The code comes as free sample and without any warranty.
%   Author: Beckhoff Automation GmbH & Co. KG
    
    properties
        AssembliesLoaded     % BOOL flag indicating successfully loaded assemblies
        VsDTE                % interface of the opended VS instance
        solution             % interface of the Solution in VS
        project              % interface pf the Tc Project
        sysManager           % interface to the system manager object
        solutionPath         % path of the solution incl. solution name
        devices              % interface to devices in TC (see ScanForBoxes); type cell array 
        TcCOM                % interface to instaciated TcCOM (see Create TcCOM); type cell array
    end
    
    methods
        %% Constructor of class
        function this = Tc3AutomationInterfaceV2
        % Constructor of class
        %   load all assemblies necessary for the Automation Interface
        %   set porperty: AssembliesLoaded = TRUE/FALSE
            try                
                % load message filter dll- for handling threading errors
                messageFilterAsmPath = which('TctComMessageFilter.dll');
                if (isempty(messageFilterAsmPath))
                    error('MessageFilter was not found');
                end
                NET.addAssembly(messageFilterAsmPath);
                
                if (~TctComMessageFilter.MessageFilter.IsRegistered)
                    TctComMessageFilter.MessageFilter.Register();
                end
                
                % load assemblies
                
                % COM lib for Visual Studio Core Automation - for VS specific commands
                NET.addAssembly('EnvDTE');   % doc -> https://docs.microsoft.com/en-us/dotnet/api/envdte
                NET.addAssembly('EnvDTE80'); % doc -> https://docs.microsoft.com/en-us/dotnet/api/envdte80
                
                % Lib for parsing XML docs like tmc files
                NET.addAssembly('System.Xml'); % doc -> https://msdn.microsoft.com/library/system.xml.aspx
                
                % COM lib for TwinCAT: TC3 Automation Interface
                NET.addAssembly('TCatSysManagerLib');
                
                % TwinCAT ADS, e.g. to set TC3 into run or config mode
                NET.addAssembly('TwinCAT.Ads');
                
                this.AssembliesLoaded = true;
            catch e
                disp(e.message);
                this.AssembliesLoaded = false;
            end
            
            % initate empty cell arrays
            this.TcCOM = {};
            this.devices = {};
        end
        
        %% functions
        function CreateNewVisualStudioInstance(this, VsVersion, visible)
        % CreateNewVisualStudioInstance  creates new Visual Studio instance
        %   
        %   CreateNewVisualStudioInstance(VsVersion, visible)
        %   VsVersion = [10] or .. 11 12 14 15
        %   visible = true/false sets visibility of VS windows
        %   set property: VsDTE
            
            % get VS Type and create Instance
            VsProgID = ['VisualStudio.DTE.' num2str(VsVersion) '.0']; % program identifier Visual Studio 2015
            t = System.Type.GetTypeFromProgID(VsProgID); % get assicianted type with ProgID
            this.VsDTE = EnvDTE80.DTE2(System.Activator.CreateInstance(t)); % create VS instance
            
            % set Visibility
            this.VsDTE.MainWindow.Visible = visible; % set visibility of main developement window of VS (true/false)
            this.VsDTE.SuppressUI = ~visible;   % set weather UI should be displayed
            
            % activate siltent mode
            % doc -> https://infosys.beckhoff.com/content/1033/tc3_automationinterface/2489025803.html
            settings = TCatSysManagerLib.ITcAutomationSettings(this.VsDTE.GetObject('TcAutomationSettings'));
            settings.SilentMode = true;
        end
        
        function OpenTwinCatSolution(this, SolutionDirectory, SolutionName, ProjectNumber)
        % OpenTwinCatSolution  opens existing TwinCAT Solution 
        % 
        %   Tc3_AI.OpenTwinCatSolution(SolutionDirectory, SolutionName, ProjectNumber)
        %   open an existing TwinCAT solution in Visual Studio (VsDTE) 
        %   in the specified directors (SolutionDirectory) with the given 
        %   solution name (SolutionName). The properties project and 
        %   sysManager are set to the specified project (ProjectNumber).
        %
        %   set properties: solutionPath, solution, project, sysManager
        %
        %   See also
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/36028797261893003.html"
        %   >Beckhoff Infosys</a>    
        
            % set default value
            if nargin < 4 
                ProjectNumber = 1;
            end
            
            % set solutionPath, open solution
            this.solutionPath = [SolutionDirectory,'\',SolutionName,'.sln'];
            this.solution = this.VsDTE.Solution;
            this.solution.Open(this.solutionPath);
            
            % set project
            if this.solution.Projects.Count > 1
                disp('Found more than one project in opened solution.')
            end
            this.project = this.solution.Projects.Item(ProjectNumber);
            
            % create sysManager
            this.sysManager = TCatSysManagerLib.ITcSysManager7(this.project.Object);
        end
        
        function CreateTwinCatSolution(this, SolutionDirectory, SolutionName)
        % CreateTwinCatSolution  creates a new TwinCAT Solution
        %   
        %   CreateNewTcSolution(SolutionDirectory, SolutionName)
        %   create a new Standard TwinCAT Solution in VS (VsDTE) in the
        %   specified directory (SolutionDirectory) with the given 
        %   solution name (SolutionName).
        %
        %   set properties: solutionPath, solution, project, sysManager
        %   
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/45035996516426763.html"
        %   >Beckhoff Infosys</a>
            
            % set solutionPath
            this.solutionPath = [SolutionDirectory,'\',SolutionName,'.sln'];
            
            % create directory and solution
            [solDir, solName, ~] = fileparts(this.solutionPath);
            if (~exist(solDir, 'dir'))
                mkdir(solDir);
            end
            this.solution = this.VsDTE.Solution;   % handle to solution layer
            this.solution.Create(solDir,solName);  % create a new empty solution
            
            % set project and template path (standard TwinCAT project template)
            projectPath  = [SolutionDirectory,'\',SolutionName,'\',SolutionName,'.tsproj'];
            templatePath = [getenv('TwinCAT3Dir') 'Components\Base\PrjTemplate\TwinCAT Project.tsproj'];
            
            % creaty project and sysManager
            [projDir, projName, ~] = fileparts(projectPath);
            this.project = this.solution.AddFromTemplate(templatePath, projDir, projName, false); % add new project using a template
            this.sysManager = TCatSysManagerLib.ITcSysManager7(this.project.Object); % from now we are inside a TwinCAT Project, so we use the TC Automation Interface
        end
        
        function TcComObject = CreateTcCOM(this, Modelname)
        % CreateTcCOM  creates a new instance of a TcCOM
        %
        %   TcComObject = CreateTcCOM(Modelname)
        %   Instanciates the TcCOM with the specified name (Modelname).
        %   Also a task with a matching cycle time is created and linked to
        %   the TcCOM-Object.
        %
        %   set properties: TcCOM
        %
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/45035996991714699.html"
        %   >Beckhoff Infosys</a>   
        
            % (1) parse the tmc file to get necessary task settings
            customModulesPath = [getenv('TwinCat3Dir') '\CustomConfig\Modules\'];
            classFactoryName = Modelname;
            if length(classFactoryName) > 35
                classFactoryName = classFactoryName(1:35);
            end
            xTmcFile = System.Xml.XmlDocument;
            xTmcFile.Load(strcat(customModulesPath,classFactoryName,'\',classFactoryName,'.tmc'));
            
            classId = xTmcFile.SelectSingleNode('TcModuleClass/Modules/Module/CLSID').InnerText;
            
            xTasks = xTmcFile.SelectNodes('TcModuleClass/Modules/Module/Contexts/Context');
            cycleTime = cell(1, numel(xTasks));
            priority = cell(1, numel(xTasks));
            
            for i = 1:xTasks.Count
                cycleTime{i} = char(xTasks.Item(i-1).SelectSingleNode('CycleTime').InnerText);
                priority{i} = char(xTasks.Item(i-1).SelectSingleNode('Priority').InnerText);
            end
            
            % (2) Create a TcCom Object
            
            % doc -> https://infosys.beckhoff.com/content/1033/tc3_automationinterface/27021598006995083.html?id=5908679173508422650
            parentTreeItem = this.sysManager.LookupTreeItem('TIRC^TcCOM Objects'); % get a pointer to a tree item; shortcut TIRC = real-time configuration, see also link above
            TcComObject = parentTreeItem.CreateChild(Modelname,0,'',classId); % create TcCOM using model name and the class ID
            
            % (3) create cylic task(s)
            Tasks = this.sysManager.LookupTreeItem('TIRT'); % TIRT: Real-Time Configuration" TAB "Additional Tasks"
            xDocTask = System.Xml.XmlDocument;
            TaskOID = cell(1,numel(cycleTime));
            for i = 1:numel(cycleTime)
                Task = Tasks.CreateChild(['TaskFor' Modelname '_' num2str(i)],1,[],[]);
                Task.ConsumeXml(['<TreeItem><TaskDef><Priority>' priority{i} '</Priority></TaskDef></TreeItem>']);
                Task.ConsumeXml(['<TreeItem><TaskDef><CycleTime>' cycleTime{i}(1:length(cycleTime{i})-2) '</CycleTime></TaskDef></TreeItem>']); 
                % Has to be scaled to base tick so delete last two zeroes: cycleTime(1:length(cycleTime)-2)
                
                xDocTask.LoadXml(Task.ProduceXml());
                TaskOID{i} = char(xDocTask.SelectSingleNode('TreeItem/ObjectId').InnerXml);
            end
            
            % (4) Append Task(s) to TcCom Object
            for i = 1:numel(TaskOID)
                xDocTComObj = System.Xml.XmlDocument;
                xDocTComObj.LoadXml(TcComObject.ProduceXml());
                xContext = xDocTComObj.SelectSingleNode(['TreeItem/TcModuleInstance/Module/Contexts/Context[Id=' num2str(i-1) ']']);
                XManualConfig = xContext.OwnerDocument.CreateElement('ManualConfig');
                xOTCID = xContext.OwnerDocument.CreateElement('OTCID');
                xOTCID.InnerText = char(TaskOID{i});
                XManualConfig.AppendChild(xOTCID);
                xContext.AppendChild(XManualConfig);
                TcComObject.ConsumeXml(xDocTComObj.InnerXml);
            end
                        
            this.TcCOM = [this.TcCOM {TcComObject}];            
        end
        
        function LinkVariables(this, TreePathVar1, TreePathVar2)
        % LinkVariables  links 2 variables
        %
        %   LinkVariables(TreePathVar1, TreePathVar2)
        %   Links two variables, which are specified by their paths 
        %   (TreePathVar1, TreePathVar2).
        %   
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/27021598596060555.html"
        %   >Beckhoff Infosys</a>
            
            this.sysManager.LinkVariables(TreePathVar1, TreePathVar2);
            
        end
        
        function Boxes = ScanEtherCatMasters(this, AmsNetId)
        % ScanEtherCatMasters scans the target system for EtherCAT masters
        %
        %   Boxes = ScanEtherCatMasters(AmsNetId)
        %   If the target system specified by its Net-ID (AmsNetId) is in
        %   Config-Mode its scanned for EtherCAT masters. 
        %   
        %   set property: devices
        %
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/45035996516641675.html"
        %   >Beckhoff Infosys</a>
            Boxes = {};

            if strcmpi(AmsNetId, 'Local')
                AmsNetId = char(TwinCAT.Ads.AmsNetId.Local.ToString()); % here local
            end

            if (this.ReadTwinCatState(AmsNetId) ~= TwinCAT.Ads.AdsState.Config)
                error('target system is not in config mode, devices could not be scanned');                
            else
                % change to Target systen with AmsNetId
                this.sysManager.SetTargetNetId(AmsNetId);

                pause(2) % wait for target

                % get handle to devices item 
                ioDevicesItem = this.sysManager.LookupTreeItem('TIID');

                % xml work
                scannedXml = ioDevicesItem.ProduceXml(false);
                xmlDoc = System.Xml.XmlDocument;
                xmlDoc.LoadXml(scannedXml);
                xmlDeviceList = xmlDoc.SelectNodes('TreeItem/DeviceGrpDef/FoundDevices/Device');

                this.devices = {};
                ECDevicesFound = 0;
                % get devices
                for ii = 0:1:(xmlDeviceList.Count-1) % devices start with 0

                    % get next device item
                    node = xmlDeviceList.Item(ii);

                    % get node specification
                    typeName    = node.SelectSingleNode('ItemSubTypeName').InnerText;
                    xmlAddress  = node.SelectSingleNode('AddressInfo');                                
                    itemSubType = int32(str2num(char(node.SelectSingleNode('ItemSubType').InnerText)));
                    
                    % ignore devices that are not EtherCAT masters
                    if(itemSubType == 111)
                        % add found node to device tree
                        ECDevicesFound = ECDevicesFound + 1;
                        device      = ioDevicesItem.CreateChild(System.String.Format('Device_{0}_{1}',ii,typeName),itemSubType,'',{});                
                        xml         = System.String.Format('<TreeItem><DeviceDef>{0}</DeviceDef></TreeItem>',xmlAddress.OuterXml);

                        device.ConsumeXml(xml);

                        this.devices = [this.devices {device}];
                     
                    end
                end
                X =[num2str(ECDevicesFound), ' EtherCAT Masters Found'];
                disp(X)
                % scan for boxes
                for ii = 1:1:(numel(this.devices)) 
                    xml = '<TreeItem><DeviceDef><ScanBoxes>1</ScanBoxes></DeviceDef></TreeItem>';
                    try
                        disp(['---- ' char(this.devices{ii}.Name) ' ----'])
                        this.devices{ii}.ConsumeXml(xml);                   
                    catch e
                        disp(e.message)
                    end

                    %for jj = 1:1:(this.devices{ii}.ChildCount) % childs of devices start with 1
                    %    Boxes = [Boxes, char(this.devices{ii}.Child(jj).PathName)]; %#ok<AGROW>
                    %    disp(['Found Box: ', char(this.devices{ii}.Child(jj).PathName)]) % use this for linking
                    %end
                    if this.devices{ii}.ChildCount == 0
                        disp('No Terminals found.')
                    else
                        X = [num2str(this.devices{ii}.ChildCount),' Terminals Found'];
                        disp(X)
                    end
                end
            end
        end
        
        function SaveSolution(this)
        % SaveSolution  saves the current TC3 Project and VS Solution
            this.project.Save();
            this.solution.SaveAs(this.solutionPath);
        end
        
        function ActivateOnDevice(this, AmsNetId)
        % ActivateOnDevice  activates project and starts TC on target system
        %
        %   ActivateOnDevice(AmsNetId)
        %   Activates the project held in sysMamager on the specified
        %   TwinCAT system (AmsNetId). The specified system is
        %   started/restarted afterwards.
            
            if strcmpi(AmsNetId, 'Local')
                AmsNetId = char(TwinCAT.Ads.AmsNetId.Local.ToString()); 
            end
            
            % set to config mode
            if ~(this.ReadTwinCatState(AmsNetId) == TwinCAT.Ads.AdsState.Config)
                this.SwitchTwinCatState(AmsNetId, TwinCAT.Ads.AdsState.Reconfig);
            end
            
            % Automation Interface again
            this.sysManager.SetTargetNetId(AmsNetId);
            pause(2);
            this.sysManager.ActivateConfiguration();
            pause(2);
            this.sysManager.StartRestartTwinCAT();
            pause(5);
            
        end
        
        function AdsState = ReadTwinCatState(~, AmsNetId)
        % ReadTwinCatState  reads the state of a TC system
        %   
        %   AdsState = ReadTwinCatState(AmsNetId)
        %   Reads the State (AdsState) of the target system specified by its
        %   Net-ID (AmsNetId).
        %
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1031/tcadswcf/html/tcadswcf.tcadsservice.enumerations.adsstate.html?"
        %   >Beckhoff Infosys: Ads State</a>
        
            if strcmpi(AmsNetId, 'Local')
                AmsNetId = char(TwinCAT.Ads.AmsNetId.Local.ToString()); 
            end
            client = TwinCAT.Ads.TcAdsClient();
            client.Connect(AmsNetId, 10000);
            
            % get system state
            [~, currentState] = client.TryReadState(); % read current Twincat state (e.g. run, config, ...)
            AdsState = (currentState.AdsState);
        end
        
        function SwitchTwinCatState(~, AmsNetId, AdsState)
        % SwitchTwinCatState  switchs the state of a TC system
        %   
        %   SwitchTwinCatState(AmsNetId, AdsState)
        %   Sets the state (AdsState) of the target system specified by its
        %   Net-ID (AmsNetId).
        %
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1031/tcadswcf/html/tcadswcf.tcadsservice.enumerations.adsstate.html?"
        %   >Beckhoff Infosys: Ads State</a>
        
            if strcmpi(AmsNetId, 'Local')
                AmsNetId = char(TwinCAT.Ads.AmsNetId.Local.ToString()); 
            end
            client = TwinCAT.Ads.TcAdsClient();
            client.Connect(AmsNetId, 10000);
            
            % get system state
            [~, currentState] = client.TryReadState(); % read current Twincat state (e.g. run, config, ...)
            
            % set ads state
            currentState.AdsState = AdsState;
            client.WriteControl(currentState);
            pause(5);
        end
    end
    
    methods (Static)        
        function vsVersions = GetInstalledVisualStudios
        % GetInstalledVisualStudios  gets all installes VS versions
        %
        %   vsVersions = GetInstalledVisualStudios
        %   Static Method. Looks up all installed Visual Studio versions.
        %   Returns an array of numbers (vsVersions)
        %   
        %   10 -> VS 2010
        %   11 -> VS 2012
        %   12 -> VS 2013
        %   14 -> VS 2015
        %   15 -> VS 2017
        %
        %   see also:
        %   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/27021598006969227.html?id=5787876362957722402"
        %   >Beckhoff Infosys</a>
        
            vsVersions = [];
            
            for vsVer = [10 11 12 14, 15] 
                for vsKey={'SOFTWARE\Microsoft\VisualStudio\' 'SOFTWARE\Wow6432Node\Microsoft\VisualStudio\'}
                    try  %#ok<TRYNC>
                        key = [vsKey{1} num2str(vsVer) '.0'];
                        winqueryreg('HKEY_LOCAL_MACHINE', key , 'InstallDir');
                        vsVersions = vertcat(vsVersions,vsVer);  %#ok<AGROW>
                    end
                end
            end
        end
    end
    
end
