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
            this.Devices = {};
            Boxes = {};
            if (this.ReadTwinCatState(AmsNetId) ~= TwinCAT.Ads.AdsState.Config)
                error('target system is not in config mode, devices could not be scanned');                
            else
                % change to Target systen with AmsNetId
                this.sysManager.SetTargetNetId(AmsNetId);

                this.Devices = GetDeviceMasters(sysManager);
                
                EcMasters = ScanForEcMasters(this.Devices);
                
                % scan for boxes
                for ii = 1:1:(numel(this.devices)) 
                    Boxes = GetEcNetworkBoxes(EcMasters(ii));
                end
                
            end
        end 
        
        function devices = GetDeviceMasters(this, sysMan)
            % get handle to devices item 
            ioDevicesItem = this.sysManager.LookupTreeItem('TIID');

            % xml work
            scannedXml = ioDevicesItem.ProduceXml(false);
            xmlDoc = System.Xml.XmlDocument;
            xmlDoc.LoadXml(scannedXml);
            xmlDeviceList = xmlDoc.SelectNodes('TreeItem/DeviceGrpDef/FoundDevices/Device');

            devices = xmlDeviceList;
        end
        
        function EcMasters = ScanForEcMasters(this, deviceList)
            ECDevicesFound = 0;
            % get devices
            for ii = 0:1:(deviceList.Count-1) % devices start with 0

                % get next device item
                node = deviceList.Item(ii);

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

                    EcMasters = [EcMasters {device}];

                end
            end
        end
        
        function EtherCATBoxes = GetEcNetworkBoxes(this, device)
            xml = '<TreeItem><DeviceDef><ScanBoxes>1</ScanBoxes></DeviceDef></TreeItem>';
            try
                this.devices.ConsumeXml(xml);                   
            catch e
                disp(e.message);
            end
            for jj = 1:1:(this.devices.ChildCount) % childs of devices start with 1
                EtherCATBoxes = [Boxes, char(this.devices{ii}.Child(jj).PathName)];
                %disp(['Found Box: ', char(this.devices{ii}.Child(jj).PathName)]) % use this for linking
            end
        end
        
        function EtherCATMaster = AddEtherCATNetwork(this, networkFile)
            this.Devices = this.sysManager.LookupTreeItem("TIID");
            %this.EcMaster = this.Devices.CreateChild("EtherCAT Master", 111, '', {});
            this.EcMaster = this.Devices.ImportChild(networkFile,"", true, "Device 5 (EtherCAT)"); 
            EtherCATMaster = this.EcMaster;
        end 
        
        function Child = AddEcSlave(this, ParentNode)%rewrite this to add slave 1 at a time
            ek1100 = ParentNode.CreateChild('EK1100', 9099, '', 'EK1100-0000-0001');
            el1004 = ek1100.CreateChild('EL1004', 9099, '', 'EL1004-0000-0000');
        end
         
        function CanDevice = AddCanInterface(this)
            this.Devices = this.sysManager.LookupTreeItem('TIID');
            this.CAN = this.Devices.CreateChild("CANDevice", 87, '', {});
            CanDevice = this.CAN.ImportChild(['C:\Users\chrisk\Desktop\', 'Box 27 (CAN Interface).xti'],'',true,'CanDevice-1');
        end
    end
end

