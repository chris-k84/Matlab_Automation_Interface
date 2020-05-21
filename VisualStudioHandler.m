classdef VisualStudioHandler < handle
    %Visual Studio Handler Class to load in VS instances
    
    properties
        AssembliesLoaded     % BOOL flag indicating successfully loaded assemblies
        VsDTE                % interface of the opended VS instance
        sysManager
        solutionPath 
        solution             % interface of the Solution in VS
        project              % interface pf the Tc Project
    end
    
    methods
        %Constructor
         function this = VisualStudioHandler
        % Constructor of class
        %   load all assemblies necessary for the Automation Interface
        %   set porperty: AssembliesLoaded = TRUE/FALSE
            try                
                %load message filter dll- for handling threading errors
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
                disp('test');
                warning('So this happened');
                this.AssembliesLoaded = false;
            end
         end
         
         function CreateNewVisualStudioInstance(this, VsVersion, visible)
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
        
         function a = get.sysManager(this)
             a = this.sysManager;
         end
             
         function SaveSolution(this)
        % SaveSolution  saves the current TC3 Project and VS Solution
            this.project.Save();
            this.solution.SaveAs(this.solutionPath);
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