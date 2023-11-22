classdef TwinCATHandler < handle
    %TWINCATHANDLER class used to maipulate TwinCAT project
    
    properties
        sysManager           
        TcCOM  
    end
    
    methods
        
        function this = TwinCATHandler(sysManager)
            this.sysManager = sysManager;
        end
        
        function ActivateOnDevice(this, AmsNetId)
        % ActivateOnDevice  activates project and starts TC on target system            
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
        
        function LinkVariables(this, TreePathVar1, TreePathVar2)
        % LinkVariables  links 2 variables
            
            this.sysManager.LinkVariables(TreePathVar1, TreePathVar2);
            
        end
        
        function CreateTask(this)
            Tasks = this.sysManager.LookupTreeItem('TIRT'); % TIRT: Real-Time Configuration" TAB "Additional Tasks"
            Task = Tasks.CreateChild('Test',1,[],[]);
        end
        
        function TcComObject = AddTcCOM(this, Modelname, vendorName, driverName)
        % CreateTcCOM  creates a new instance of a TcCOM
			%
			%   TcComObject = CreateTcCOM(Modelname, vendorName, driverName)
			%   Instanciates the TcCOM with the specified name (Modelname).
			%   Also a task with a matching cycle time is created and linked to
			%   the TcCOM-Object.
			%   For versioned modules, i.e. Repository drivers, the vendor name
			%   and driver name needs to be given. If no driverName is passed,
			%   the ModelName is assumed to be driverName. The module with the 
			%   maximum version number will get created.
			%
			%   set properties: TcCOM
			%
			%   see also:
			%   <a href="https://infosys.beckhoff.com/content/1033/tc3_automationinterface/45035996991714699.html"
			%   >Beckhoff Infosys</a>
			
			% (1) parse the tmc file to get necessary task settings
			if nargin > 2
				if nargin < 4
					driverName = Modelname(1:min(length(Modelname),35));
				end
				
				try
					repoDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'RepositoryDir');
				catch
					repoDir = fullfile(getenv('TwinCAT3Dir'),'Repository');
				end
				
				versionDir = fullfile(repoDir,vendorName,driverName);
				if ~isfolder(versionDir)
					error(['The folder ' versionDir ' does not exist']);
				end
				
				% select max version number
				maxVer = System.Version.Parse('0.0.0.0');
				versionNumberDirs = dir(versionDir);
				for i=1:length(versionNumberDirs)
					if versionNumberDirs(i).isdir
						[isValidVersion,version] = System.Version.TryParse(versionNumberDirs(i).name);
						if isValidVersion && ~isempty(version) && version > maxVer
							maxVer = version;
						end
					end
				end
				
				versionNo = char(maxVer.ToString());
				if strcmp(versionNo,'0.0.0.0')
					error('No driver with valid version number found');
				end
				
				tmcFolder = fullfile(repoDir,vendorName,driverName,versionNo,[driverName '.tmc']);
			else
				try
					customConfigDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'CustomConfigDir');
				catch
					customConfigDir = fullfile(getenv('TwinCAT3Dir'),'CustomConfig');
				end
				customModulesPath = fullfile(customConfigDir, 'Modules');
				classFactoryName = Modelname;
				if length(classFactoryName) > 35
					classFactoryName = classFactoryName(1:35);
				end
				
				tmcFolder = fullfile(customModulesPath,classFactoryName,[classFactoryName '.tmc']);
			end
			
			if ~isfile(tmcFolder)
				error(['Unable to open file ' tmcFolder]);
			end
			
			xTmcFile = System.Xml.XmlDocument;
			xTmcFile.Load(tmcFolder);
			
			xModule = [];
			xModules = xTmcFile.SelectNodes('TcModuleClass/Modules/Module');
			if xModules.Count > 1
				for i = 1:xModules.Count
					module = xModules.Item(i-1);
					xModuleName = char(module.SelectSingleNode('Name').InnerText);
					if strcmp(xModuleName,Modelname)
						xModule = module;
						break
					end
				end
			elseif xModules.Count == 0
				error('The loaded tmc file does not contain a Module Node');
			else
				xModule = xModules.Item(0);
			end
			
			if isempty(xModule)
				error('Unable to locate the Module node in the tmc file');
			end
			
			classId = xModule.SelectSingleNode('CLSID').InnerText;
			
			
			xTasks = xModule.SelectNodes('Contexts/Context');
			cycleTime = cell(1, numel(xTasks));
			priority = cell(1, numel(xTasks));
			
			for i = 1:xTasks.Count
				xCycleTime = xTasks.Item(i-1).SelectSingleNode('CycleTime');
				if ~isempty(xCycleTime)
					cycleTime{i} = char(xCycleTime.InnerText);
				end
				xPriority = xTasks.Item(i-1).SelectSingleNode('Priority');
				if ~isempty(xPriority)
					priority{i} = char(xPriority.InnerText);
				end
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
				if ~isempty(priority{i})
					Task.ConsumeXml(['<TreeItem><TaskDef><Priority>' priority{i} '</Priority></TaskDef></TreeItem>']);
				end
				if ~isempty(cycleTime{i})
					Task.ConsumeXml(['<TreeItem><TaskDef><CycleTime>' cycleTime{i}(1:length(cycleTime{i})-2) '</CycleTime></TaskDef></TreeItem>']); % Has to be scaled to base tick so delete last two zeroes: cycleTime(1:length(cycleTime)-2)
				end
				
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
        
        function TcComObject = AddTcCOMModule(this, Modelname, vendorName, driverName)
			if nargin > 2
				if nargin < 4
					driverName = Modelname(1:min(length(Modelname),35));
				end
				
				try
					repoDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'RepositoryDir');
				catch
					repoDir = fullfile(getenv('TwinCAT3Dir'),'Repository');
				end
				
				versionDir = fullfile(repoDir,vendorName,driverName);
				if ~isfolder(versionDir)
					error(['The folder ' versionDir ' does not exist']);
				end
				
				% select max version number
				maxVer = System.Version.Parse('0.0.0.0');
				versionNumberDirs = dir(versionDir);
				for i=1:length(versionNumberDirs)
					if versionNumberDirs(i).isdir
						[isValidVersion,version] = System.Version.TryParse(versionNumberDirs(i).name);
						if isValidVersion && ~isempty(version) && version > maxVer
							maxVer = version;
						end
					end
				end
				
				versionNo = char(maxVer.ToString());
				if strcmp(versionNo,'0.0.0.0')
					error('No driver with valid version number found');
				end
				
				tmcFolder = fullfile(repoDir,vendorName,driverName,versionNo,[driverName '.tmc']);
			else
				try
					customConfigDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'CustomConfigDir');
				catch
					customConfigDir = fullfile(getenv('TwinCAT3Dir'),'CustomConfig');
				end
				customModulesPath = fullfile(customConfigDir, 'Modules');
				classFactoryName = Modelname;
				if length(classFactoryName) > 35
					classFactoryName = classFactoryName(1:35);
				end
				
				tmcFolder = fullfile(customModulesPath,classFactoryName,[classFactoryName '.tmc']);
			end
			
			if ~isfile(tmcFolder)
				error(['Unable to open file ' tmcFolder]);
			end
			
			xTmcFile = System.Xml.XmlDocument;
			xTmcFile.Load(tmcFolder);
			
			xModule = [];
			xModules = xTmcFile.SelectNodes('TcModuleClass/Modules/Module');
			if xModules.Count > 1
				for i = 1:xModules.Count
					module = xModules.Item(i-1);
					xModuleName = char(module.SelectSingleNode('Name').InnerText);
					if strcmp(xModuleName,Modelname)
						xModule = module;
						break
					end
				end
			elseif xModules.Count == 0
				error('The loaded tmc file does not contain a Module Node');
			else
				xModule = xModules.Item(0);
			end
			
			if isempty(xModule)
				error('Unable to locate the Module node in the tmc file');
			end
			
			classId = xModule.SelectSingleNode('CLSID').InnerText;

			parentTreeItem = this.sysManager.LookupTreeItem('TIRC^TcCOM Objects'); % get a pointer to a tree item; shortcut TIRC = real-time configuration, see also link above
			TcComObject = parentTreeItem.CreateChild(Modelname,0,'',classId); % create TcCOM using model name and the class ID

			this.TcCOM = [this.TcCOM {TcComObject}];            
        end

        function GetTcCOMContext()
			if nargin > 2
				if nargin < 4
					driverName = Modelname(1:min(length(Modelname),35));
				end
				
				try
					repoDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'RepositoryDir');
				catch
					repoDir = fullfile(getenv('TwinCAT3Dir'),'Repository');
				end
				
				versionDir = fullfile(repoDir,vendorName,driverName);
				if ~isfolder(versionDir)
					error(['The folder ' versionDir ' does not exist']);
				end
				
				% select max version number
				maxVer = System.Version.Parse('0.0.0.0');
				versionNumberDirs = dir(versionDir);
				for i=1:length(versionNumberDirs)
					if versionNumberDirs(i).isdir
						[isValidVersion,version] = System.Version.TryParse(versionNumberDirs(i).name);
						if isValidVersion && ~isempty(version) && version > maxVer
							maxVer = version;
						end
					end
				end
				
				versionNo = char(maxVer.ToString());
				if strcmp(versionNo,'0.0.0.0')
					error('No driver with valid version number found');
				end
				
				tmcFolder = fullfile(repoDir,vendorName,driverName,versionNo,[driverName '.tmc']);
			else
				try
					customConfigDir = winqueryreg('HKEY_LOCAL_MACHINE', 'Software\Wow6432Node\Beckhoff\TwinCAT3\3.1', 'CustomConfigDir');
				catch
					customConfigDir = fullfile(getenv('TwinCAT3Dir'),'CustomConfig');
				end
				customModulesPath = fullfile(customConfigDir, 'Modules');
				classFactoryName = Modelname;
				if length(classFactoryName) > 35
					classFactoryName = classFactoryName(1:35);
				end
				
				tmcFolder = fullfile(customModulesPath,classFactoryName,[classFactoryName '.tmc']);
			end
			
			if ~isfile(tmcFolder)
				error(['Unable to open file ' tmcFolder]);
			end
			
			xTmcFile = System.Xml.XmlDocument;
			xTmcFile.Load(tmcFolder);
			
			xModule = [];
			xModules = xTmcFile.SelectNodes('TcModuleClass/Modules/Module');
			if xModules.Count > 1
				for i = 1:xModules.Count
					module = xModules.Item(i-1);
					xModuleName = char(module.SelectSingleNode('Name').InnerText);
					if strcmp(xModuleName,Modelname)
						xModule = module;
						break
					end
				end
			elseif xModules.Count == 0
				error('The loaded tmc file does not contain a Module Node');
			else
				xModule = xModules.Item(0);
			end
			
			if isempty(xModule)
				error('Unable to locate the Module node in the tmc file');
			end
			
			classId = xModule.SelectSingleNode('CLSID').InnerText;
			
			
			xTasks = xModule.SelectNodes('Contexts/Context');
			cycleTime = cell(1, numel(xTasks));
			priority = cell(1, numel(xTasks));
			
			for i = 1:xTasks.Count
				xCycleTime = xTasks.Item(i-1).SelectSingleNode('CycleTime');
				if ~isempty(xCycleTime)
					cycleTime{i} = char(xCycleTime.InnerText);
				end
				xPriority = xTasks.Item(i-1).SelectSingleNode('Priority');
				if ~isempty(xPriority)
					priority{i} = char(xPriority.InnerText);
				end
            end         
        end

        function TaskId = CreateTaskwithContext(this, Name, Priority, Cycle)
			Tasks = this.sysManager.LookupTreeItem('TIRT'); % TIRT: Real-Time Configuration" TAB "Additional Tasks"
			xDocTask = System.Xml.XmlDocument;
			Task = Tasks.CreateChild(Name,1,[],[]);
			if ~isempty(Priority)
				Task.ConsumeXml(['<TreeItem><TaskDef><Priority>' num2str(Priority) '</Priority></TaskDef></TreeItem>']);
			end
			if ~isempty(Cycle)
				Task.ConsumeXml(['<TreeItem><TaskDef><CycleTime>' num2str(Cycle) '</CycleTime></TaskDef></TreeItem>']); % Has to be scaled to base tick so delete last two zeroes: cycleTime(1:length(cycleTime)-2)
			end
			
			xDocTask.LoadXml(Task.ProduceXml());
			TaskId = char(xDocTask.SelectSingleNode('TreeItem/ObjectId').InnerXml);
        end

        function AssignTaskToTcCOM(this, TaskId, TcComObject)
			xDocTComObj = System.Xml.XmlDocument;
			xDocTComObj.LoadXml(TcComObject.ProduceXml());
			xContext = xDocTComObj.SelectSingleNode(['TreeItem/TcModuleInstance/Module/Contexts/Context[Id=' num2str(0) ']']);
			XManualConfig = xContext.OwnerDocument.CreateElement('ManualConfig');
			xOTCID = xContext.OwnerDocument.CreateElement('OTCID');
			xOTCID.InnerText = char(TaskId);
			XManualConfig.AppendChild(xOTCID);
			xContext.AppendChild(XManualConfig);
			TcComObject.ConsumeXml(xDocTComObj.InnerXml);
        end

        function setTaskProperties(~,currTask,Priority, Cycle,CpuAffinity)
            currTask.ConsumeXml(['<TreeItem><TaskDef><Priority>' Priority '</Priority></TaskDef></TreeItem>']);
            currTask.ConsumeXml(['<TreeItem><TaskDef><CycleTime>' Cycle(1:length(Cycle)-2) '</CycleTime></TaskDef></TreeItem>']); 
            % Has to be scaled to base tick so delete last two zeroes: cycleTime(1:length(cycleTime)-2)
            currTask.ConsumeXml(['<TreeItem><TaskDef><CpuAffinity>' CpuAffinity '</CpuAffinity></TaskDef></TreeItem>']);
        end    
    end
end

