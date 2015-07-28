from read_data import *

__version__ = 3.0

#-------------------------------------------------------------
# Define the order that the configuration will be generated
#-------------------------------------------------------------
def show_all_config():
    console = sys.__stdout__

    print('\nGenerate config files:')
    for device in sorted(d.devices):
        print ('  -- New File: {0: <22} [complete]'.format('ccg-'+device+'.txt'))
        logfile = Logger('ccg-{}.txt'.format(device))
        sys.stdout = logfile

        # order listed below
        show_global_config(device,'start')
        show_vrf_config(device)
        show_vlans_config(device)
        show_interface_config(device,'physical')
        show_interface_config(device,'logical')
        show_routing_config(device)
        show_global_config(device,'end')
        sys.stdout = console

#--------------------------------------------------
# Generate the global configuration for the device
#--------------------------------------------------
def show_global_config(device_name=None, config_position=None):
    template_list = get_template_list(device_name, config_position)
    if template_list:
        print ("\n!----------------------------------------------------")
        print ("! Global configuration for ({}) @ {}  ".format(device_name,config_position))
        print ("!------------------------------------------------------")
    for template in template_list:
        config_template = d.templates[template]
        print ("\n! [{} template used]:".format(template))
        for config in config_template:
            print (config)
#------------------------------------------------
# Generate the VRF configuration for the device
#------------------------------------------------
def show_vrf_config(device_name=None):
    device = get_device(device_name)
    if device.vrfs:
        print ('!---------------------------------')
        print ('! VRF configuration              ')
        print ('!---------------------------------')
        for vrf in sorted(device.vrfs):
            v = get_vrf(device_name,vrf)
            print ('ip vrf {}' .format(v.name))
            if v.rd:
                print (' rd {}' .format(v.rd))
            if v.import_rt:
                for target in v.import_rt:
                    print (' route-target import {}' .format(target))
            if v.export_rt:
                for target in v.export_rt:
                    print (' route-target export {}' .format(target))
            if get_variable(v.variable):
                print (' {}'.format(get_variable(v.variable)))
#----------------------------------------------------------
# Generate the static routing configuration for the device
#----------------------------------------------------------
def show_routing_config(device_name=None):
    device = get_device(device_name)
    if device.static_routes:
        print ('!---------------------------------')
        print ('! Static Routes                   ')
        print ('!---------------------------------')
        for route in sorted(device.static_routes):
            r = get_route(device_name,route)
            print (r.show_route)
#-----------------------------------------------
# Generate the vlan configuration for the device
#-----------------------------------------------
def show_vlans_config(device_name=None):
    device = get_device(device_name)
    if device.vlans:
        print ('!---------------------------------')
        print ('! VLAN configuration              ')
        print ('!---------------------------------')
        for vlan in sorted(device.vlans):
            print ('vlan {}' .format(device.vlans[vlan].number))
            print (' name {}' .format(device.vlans[vlan].name))
#-----------------------------------------------------
# Generate the interface configuration for the device
#-----------------------------------------------------
def show_interface_config(device_name=None, mode=None):
    device = get_device(device_name)
    if 'logical' in mode and not device.has_logical_interfaces:
        return
    if 'physical' in mode and not device.has_physical_interfaces:
        return
    print ('!---------------------------------')
    print ('! Interface configuration [{}]    '.format(mode))
    print ('!---------------------------------')
    #--------------------------------------------------
    # Make sure only the relevant interfaces are shown
    #--------------------------------------------------
    for interface in sorted(device.interfaces):
        intf = get_interface(device.name, interface)
        if 'logical' in mode and not intf.is_logical:
            continue
        if 'physical' in mode and intf.is_logical:
            continue
        #---------------------------------------------------------------------------
        # Start generating warning message if manual user intervention is required
        #----------------------------------------------------------------------------
        if intf.is_pc_member and 'layer3' in intf.pc_type:
            print ('!......................................................................')
            print ('!  Warning: L3 PC detected, you need to manually create {} first'.format(intf.pc_parent))
            print ('!......................................................................')
        #------------------------------------------------------
        # Start generating the actual interface configuration
        #------------------------------------------------------
        print ('!\ninterface {}'.format(intf))
        if intf.comment:
            print ('{}'.format(intf.comment))
        if 'layer2' in intf.type:
            print ('  switchport')
        if 'layer3' in intf.type:
            print ('  no switchport')
        if intf.description:
            print ('  description {}'.format(intf.description))
        #-------------------------------------
        # Generate access port configuration
        #-------------------------------------
        if intf.data_vlan:
            print ('  switchport access vlan {}'.format(intf.data_vlan))
        if intf.voice_vlan:
            print ('  switchport voice vlan {}'.format(intf.voice_vlan))
        #-------------------------------------
        # Generate trunk port configuration
        #-------------------------------------
        if intf.trunk_vlans:
            print ('  switchport mode trunk')
            print ('  switchport trunk allowed vlan {}'.format(intf.get_trunk_vlans))
        if intf.native_vlan:
            print ('  switchport trunk native vlan {}'.format(intf.native_vlan))
        #-------------------------------------
        # Generate routed port configuration
        #-------------------------------------
        if intf.vrf:
            print ('  ip vrf forwarding {}'.format(intf.vrf))
        if intf.ipaddress:
            print ('  ip address {}'.format(intf.show_ipaddress))
        #-----------------------------------------------
        # Show common port configuration for all types
        #-----------------------------------------------
        if intf.mtu:
            print ('  mtu {}'.format(intf.mtu))
        if intf.variable1:
            print ('  {}'.format(get_variable(intf.variable1)))
        if intf.variable2:
            print ('  {}'.format(get_variable(intf.variable2)))
        if intf.speed:
            print ('  speed {}'.format(intf.speed))
        if intf.duplex:
            print ('  duplex {}'.format(intf.duplex))
        #-------------------------------------
        # Generate port-channel config
        #-------------------------------------
        if intf.is_pc_member:
            print ('  channel-group {} mode {}'.format(intf.pc_group,intf.pc_mode))
        #-------------------------------------
        # Generate last interface config
        #-------------------------------------
        if 'yes' in intf.enabled:
            print ('  no shutdown')
        if 'no' in intf.enabled:
            print ('  shutdown')

def main(argv):
    arg_length = len(sys.argv)

    print ('+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+')
    print ('  Cisco Config Generator v{}'.format(__version__))
    print ('+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+')
    if arg_length < 2:
        print ('Usage: python {} <spreadsheet.xlsx>' .format(sys.argv[0]))
        exit()
    if sys.argv[1]:
        filename = sys.argv[1]
    try:
        initalise_data(filename)
        print ('Data read from: \'{}\''.format(filename))
        show_all_config()
    except IOError:
        exit()

if __name__ == '__main__':
    main(sys.argv)

