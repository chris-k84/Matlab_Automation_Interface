classdef PLCHandler < handle
    %PLCHANDLER Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        sysManager
        plcProject
    end
    
    methods
        function this = PLCHandler(sysManager)
            this.sysManager = sysManager;
        end
        
        function CreatePLC(this, projectName)
            try
                PLC = this.sysManager.LookupTreeItem('TIPC');
                %templatePath = [getenv('TwinCAT3Dir') 'Components\Plc\PlcTemplate\Plc Templates\Standard PLC Template.plcproj'];
                %this.plcProject = PLC.CreateChild(projectName, 1, [], templatePath);
                this.plcProject = PLC.CreateChild(projectName, 0, [], 'Standard PLC Template.plcproj');
            catch e
                disp(e.message);
                disp('test');
                warning('So this happened');
            end
        end
        
    end
end

