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
                this.plcProject = PLC.CreateChild(projectName, 0, [], "Standard PLC Template");
            catch e
                disp(e.message);
                disp('test');
                warning('So this happened');
            end
        end
    end
end

