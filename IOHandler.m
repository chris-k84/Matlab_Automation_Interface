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
                
                EcMasters = CheckForEcMasters(this.Devices);
                
                % scan for boxes
                for ii = 1:1:(numel(this.devices)) 
                    Boxes = GetEcNetworkBoxes(EcMasters(ii));
                end
                
            end
        end %%remove this
        
        function devices = GetIoDeviceMasters(this, sysMan)
            %look up the IO node, return the devices listed
            ioDevices = this.sysManager.LookupTreeItem('TIID');
            scannedXml = ioDevices.ProduceXml(false);
            xmlDoc = System.Xml.XmlDocument;
            xmlDoc.LoadXml(scannedXml);
            xmlDeviceList = xmlDoc.SelectNodes('TreeItem/DeviceGrpDef/FoundDevices/Device');
            devices = xmlDeviceList;
        end
        
        function EcMasters = CheckDevicesForEcMasters(this, deviceList)
            %Take device list, search by type for EtherCAT masters (111)
            ECDevicesFound = 0;
            for ii = 0:1:(deviceList.Count-1)
                node = deviceList.Item(ii);
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
        
        function EtherCATBoxes = GetEcNetworkBoxes(this, EcMaster)
            %Consume EtherCAT master, pull out network slaves
            xml = '<TreeItem><DeviceDef><ScanBoxes>1</ScanBoxes></DeviceDef></TreeItem>';
            try
                this.EcMaster.ConsumeXml(xml);                   
            catch e
                disp(e.message);
            end
            for jj = 1:1:(this.EcMaster.ChildCount)
                EtherCATBoxes = [Boxes, char(this.EcMaster{ii}.Child(jj).PathName)];
            end
        end
        
        function EtherCATMaster = AddEtherCATNetwork(this, networkFile, networkName)
            this.Devices = this.sysManager.LookupTreeItem("TIID");
            this.EcMaster = this.Devices.ImportChild(networkFile,"", true, networkName); 
            EtherCATMaster = this.EcMaster;
        end 
         
        function CanDevice = AddCanInterface(this)
            this.Devices = this.sysManager.LookupTreeItem('TIID');
            this.CAN = this.Devices.CreateChild("CANDevice", 87, '', {});
            CanDevice = this.CAN.ImportChild(['C:\Users\chrisk\Desktop\', 'Box 27 (CAN Interface).xti'],'',true,'CanDevice-1');
        end
    end
end

