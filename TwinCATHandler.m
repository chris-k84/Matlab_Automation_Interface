classdef TwinCATHandler
    %TWINCATHANDLER Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
%         Property1
            project              % interface pf the Tc Project
            sysManager           % interface to the system manager object
    end
    
    methods
%         function obj = TwinCATHandler(inputArg1,inputArg2)
%             %TWINCATHANDLER Construct an instance of this class
%             %   Detailed explanation goes here
%             obj.Property1 = inputArg1 + inputArg2;
%         end
        
%         function outputArg = method1(obj,inputArg)
%             METHOD1 Summary of this method goes here
%               Detailed explanation goes here
%             outputArg = obj.Property1 + inputArg;
%         end
        
         function CreateTwinCatSolution(this, SolutionDirectory, SolutionName) 
        %   CreateNewTcSolution(SolutionDirectory, SolutionName)
        %   create a new Standard TwinCAT Solution in VS (VsDTE) in the
        %   specified directory (SolutionDirectory) with the given 
        %   solution name (SolutionName).
        %
        %   set properties: solutionPath, solution, project, sysManager
            
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
    end
end

