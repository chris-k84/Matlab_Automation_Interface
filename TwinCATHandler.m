classdef TwinCATHandler
    %TWINCATHANDLER Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
%         Property1
            project              % interface pf the Tc Project
            sysManager           % interface to the system manager object
    end
    
    methods
        
        function this = TwinCATHandler(sysManager)
            sysManager = sysManager;
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
    end
end

