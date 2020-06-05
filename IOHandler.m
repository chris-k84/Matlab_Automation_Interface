classdef IOHandler < handle
    %IOHANDLER class to handle IO elements of TwinCAT
    
    properties
        sysManager
        Devices
        CAN
        xmlDoc
        EcMaster
    end
    
    methods
        
        function this = IOHandler(sysManager)
            this.sysManager = sysManager;
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
        end %needs rewriting into class
        
        function EtherCATMaster = CreateEtherCAT(this)
            this.Devices = this.sysManager.LookupTreeItem("TIID");
            this.EcMaster = this.Devices.CreateChild("EtherCAT Master", 111, '', {});
        end %need to retuern EcMaster ref
        function this = AddEcSlave(this)%rewrite this to add slave 1 at a time
            ek1100 = this.EcMaster.CreateChild('EK1100', 9099, '', 'EK1100-0000-0001');
            el1004 = ek1100.CreateChild('EL1004', 9099, '', 'EL1004-0000-0000');
            xml = el1004.ProduceXml();
            this.xmlDoc = System.Xml.XmlDocument;
            this.xmlDoc.LoadXml(xml);
            % get element in xml by tag name
            %itemName = this.xmlDoc.GetElementsByTagName("ItemName");
            % change inner text of tag
            %%if itemName ~= 0
            %    itemName.setAttribute('ChangedByAi');
            %end
            item = this.xmlDoc.GetElementsByTagName('Name')
            thislist = item.item(0);
            thislist.SetAttribute('Name','Test');
            
            el1004.ConsumeXml(this.xmlDoc.InnerXml);
        end
         
        function CanDevice = AddCanInterface(this)
            this.Devices = this.sysManager.LookupTreeItem('TIID');
            this.CAN = this.Devices.CreateChild("CANDevice", 87, '', {});
            %CanDevice = this.CAN.CreateChild("CAN", 5051, '', '');
            CanDevice = this.CAN.ImportChild(['C:\Users\chrisk\Desktop\', 'Box 27 (CAN Interface).xti'],'',true,'CanDevice-1');
            CanDeviceXml = CanDevice.ProduceXml(false);
            this.xmlDoc = System.Xml.XmlDocument;
            this.xmlDoc.LoadXml(CanDeviceXml);
        end
    end
end

