import xlrd
import re
import netaddr
import sys

DATA = {}

#-----------------------------------------
# Used to redirect output to a text file
#-----------------------------------------
class Logger(object):
    def __init__(self, filename="default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "w")
    def write(self, message):
        #self.terminal.write(message)   # Shows output to screen
        self.log.write(message)         # Writes output to file


class Database(object):
    def __init__(self):
        self.templates = {}
        self.variables = {}
        self.devices = {}

class Device(object):
    def __init__(self):
        self.name = ''
        self.templates = {}
        self.vlans = {}
        self.interfaces = {}
        self.vrfs = {}
        self.static_routes = {}

    def __repr__(self):
        return self.name

    @property
    def has_logical_interfaces(self):
        for interface in self.interfaces:
            intf = get_interface(self.name, interface)
            if intf.is_logical:
                return True
        return False

    @property
    def has_physical_interfaces(self):
        for interface in self.interfaces:
            intf = get_interface(self.name, interface)
            if not intf.is_logical:
                return True
        return False

class Variable(object):
    def __init__(self):
        self.name = ''
        self.value = ''
        self.comment = ''

    def __repr__(self):
        return self.value

class Template(object):
    def __init__(self):
        self.device = ''
        self.name = ''
        self.position = 'start'

    def __repr__(self):
        return self.name

class Vlan(object):
    def __init__(self):
        self.number = ''
        self.name = ''

    def __repr__(self):
        return self.number

class Vrf(object):
    def __init__(self):
        self.name = ''
        self.rd = ''
        self.import_rt = []
        self.export_rt = []
        self.variable = ''

    def __repr__(self):
        return self.name

class StaticRoute(object):
    def __init__(self):
        self.prefix = ''
        self.next_hop = ''
        self.name = ''
        self.vrf = ''

    def __repr__(self):
        return self.prefix

    @property
    def show_route(self):
        string = 'ip route'
        if self.vrf:
            string += ' vrf {}'.format(self.vrf)
        string += ' {}'.format(self.convert_prefix_to_ios)
        string += ' {}'.format(self.next_hop)
        if self.name:
            string += ' name {}'.format(self.name)
        return string

    @property
    def convert_prefix_to_ios(self):
        ipnetwork = netaddr.IPNetwork(self.prefix)
        string = '{} {}'.format(ipnetwork.ip, ipnetwork.netmask)
        return string

class Interface(object):
    def __init__(self):
        self.name = ''
        self.enabled = False
        self.speed = 'Auto'
        self.duplex = 'Auto'
        self.mtu = 1500
        self.description = ''
        self.variable1 = ''
        self.variable2 = ''
        self.data_vlan = ''
        self.voice_vlan = ''
        self.native_vlan = ''
        self.trunk_vlans = []
        self.interface_type = ''
        self.pc_group = ''
        self.pc_mode = ''
        self.pc_type = ''
        self.pc_members = []
        self.pc_parent = ''
        self.ipaddress = ''
        self.vrf = ''
        self.type = ''
        self.comment = ''

    def __repr__(self):
        return self.name

    @property
    def is_logical(self):
        logical_types = ["po","tu","lo","vl"]
        for type in logical_types:
            if type in self.name:
                return True
        return False

    @property
    def get_type(self):
        if self.pc_group and self.ipaddress:
            return 'layer3_pc'
        elif self.pc_group and not self.ipaddress:
            return 'layer2_pc'
        elif self.data_vlan:
            return 'layer2_access'
        elif self.trunk_vlans:
            return 'layer2_trunk'
        elif self.ipaddress:
            return 'layer3_routed'
        else:
            return 'unknown'

    @property
    def get_trunk_vlans(self):
        if not self.trunk_vlans:
            return None
        vlans = self.trunk_vlans
        vlans = vlans.replace (' ','').strip().split(',')
        vlan_list = []
        for vlan in vlans:
            if "-" not in vlan:
                vlan_list.append(str(vlan))
            elif "-" in vlan:
                start_range = int(vlan.split('-')[0])
                end_range = int(vlan.split('-')[1])
                for num in range(start_range,end_range+1):
                    vlan_list.append(str(num))
        return ','.join(vlan_list)

    @property
    def is_pc_member(self):
        if self.pc_parent:
            return True

    @property
    def is_pc_parent(self):
        if self.pc_members:
            return True

    @property
    def is_valid_ip(self):
        ipnetwork = netaddr.IPNetwork(self.ipaddress)
        ipaddress = netaddr.IPAddress(ipnetwork.ip)
        if ipaddress in ipnetwork.iter_hosts():
            return True
        return False

    @property
    def show_ipaddress(self):
        ipnet = netaddr.IPNetwork(self.ipaddress)
        string = '{} {}'.format(ipnet.ip,ipnet.netmask)
        return string

def valid_row(worksheet_name, row):
    if not worksheet_name:
        return False
    WORKSHEET_NAME = worksheet_name
    REQUIRED_FIELDS = {
        'config_templates' : [],
        'variables'        : ['Variable','Variable Value',],
        'device_types'     : ['Device Name','Device Type'],
        'device_templates' : ['Device Name','Config Template','Position (Default: Start)',],
        'vlans'            : ['Device Name','VLAN No','VLAN Name',],
        'l2_interfaces'    : ['Device Name','Interface',],
        'l3_interfaces'    : ['Device Name','Interface',],
        'vrf'              : ['Device Name','VRF',],
        'static_routes'    : ['Device Name','Route (x.x.x.x/x)','Next Hop',],
    }
    for field in REQUIRED_FIELDS[WORKSHEET_NAME]:
        if not DATA[WORKSHEET_NAME][row][field]:
            return False
    return True


#----------------------------------------------------------------
# Read information from the database.xlsx
# Information will be stored by worksheet name, row, column name
#-----------------------------------------------------------------
def read_database_from_file(filename):
    try:
        wb = xlrd.open_workbook(filename)
    except:
        print ('Cannot read data from: \'{}\''.format(filename))
        print ('Script failed.')
        exit()

    worksheet_data = []
    for i, worksheet in enumerate(wb.sheets()):
        if worksheet.name == 'Instructions': continue
        header_cells = worksheet.row(0)
        num_rows = worksheet.nrows - 1
        curr_row = 0
        header = [each.value for each in header_cells]
        while curr_row < num_rows:
            curr_row += 1
            row = [int(each.value) if isinstance(each.value, float)
                   else each.value
                   for each in worksheet.row(curr_row)]
            value_dict = dict(zip(header, row))
            worksheet_data.append(value_dict)
        else:
            DATA[worksheet.name] = worksheet_data
            worksheet_data = []

def initalise_devices():
    WORKSHEETS_TO_SEARCH = [
    'device_templates',
    'l2_interfaces',
    'l3_interfaces',
    'vlans',
    'vrf',
    'static_routes',
    ]
    for worksheet in WORKSHEETS_TO_SEARCH:
        for row_no,devices in enumerate(DATA[worksheet]):
            device = str(DATA[worksheet][row_no]['Device Name'].lower().strip())
            if device and device not in d.devices:
                new_device = Device()
                new_device.name = device
                d.devices[new_device.name] = new_device

def initalise_variables():
    WORKSHEET_NAME = 'variables'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            variable = Variable()
            variable.name  = str(DATA[WORKSHEET_NAME][row_no]['Variable'].strip())
            variable.value = str(DATA[WORKSHEET_NAME][row_no]['Variable Value'].strip())
            variable.comments = str(DATA[WORKSHEET_NAME][row_no]['Comments'])
            d.variables[variable.name] = variable

def initalise_config_templates():
    WORKSHEET_NAME = 'config-templates'
    templates = {}
    # Find all the unique templates
    for row_no in range(len(DATA[WORKSHEET_NAME])):
        line = str(DATA[WORKSHEET_NAME][row_no]["Enter config templates below this line:"])
        if not line:
            continue
        new_template = re.search(r'Config Template: \[(.*?)\]', line, re.IGNORECASE)
        if new_template:
            template_name = new_template.group(1)
            templates[template_name] = []
        else:
            templates[template_name].append(line)
    # Update the dynamic variable in each template
    for template in templates:
        updated_config = []
        for line in templates[template]:
            for variable in d.variables.keys():
                search_phrase = '[' + variable + ']'
                if search_phrase in line:
                    line = line.replace(search_phrase, str(d.variables[variable]))
            updated_config.append(line)
            d.templates[template] = updated_config

def initilise_device_templates():
    WORKSHEET_NAME = 'device_templates'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            template = Template()
            device_name       = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            template.name     = str(DATA[WORKSHEET_NAME][row_no]['Config Template'].strip())
            template.position = str(DATA[WORKSHEET_NAME][row_no]['Position (Default: Start)'].lower().strip())
            device = get_device(device_name)
            device.templates[template.name] = template

def initalise_vlans():
    WORKSHEET_NAME = 'vlans'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            vlan = Vlan()
            device_name   = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            vlan.number   = str(DATA[WORKSHEET_NAME][row_no]['VLAN No'])
            vlan.name     = str(DATA[WORKSHEET_NAME][row_no]['VLAN Name'].strip())
            device = get_device(device_name)
            device.vlans[vlan.number] = vlan

def initalise_vrfs():
    WORKSHEET_NAME = 'vrf'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            vrf = Vrf()
            device_name   = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            vrf.name      = str(DATA[WORKSHEET_NAME][row_no]['VRF'].strip())
            vrf.rd        = str(DATA[WORKSHEET_NAME][row_no]['RD'].strip())
            vrf.import_rt = str(DATA[WORKSHEET_NAME][row_no]['Import RT (separated by commas)'].strip())
            vrf.export_rt = str(DATA[WORKSHEET_NAME][row_no]['Export RT (separated by commas)'].strip())
            vrf.variable  = str(DATA[WORKSHEET_NAME][row_no]['Variable'].strip())
            if vrf.import_rt:
                vrf.import_rt = vrf.import_rt.replace(' ','').split(',')
            if vrf.export_rt:
                vrf.export_rt = vrf.export_rt.replace(' ','').split(',')
            device = get_device(device_name)
            device.vrfs[vrf.name] = vrf

def initalise_static_routes():
    WORKSHEET_NAME = 'static_routes'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            route = StaticRoute()
            device_name    = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            route.vrf      = str(DATA[WORKSHEET_NAME][row_no]['VRF (if applicable)'].strip())
            route.prefix   = str(DATA[WORKSHEET_NAME][row_no]['Route (x.x.x.x/x)'].strip())
            route.next_hop = str(DATA[WORKSHEET_NAME][row_no]['Next Hop'].strip())
            route.name     = str(DATA[WORKSHEET_NAME][row_no]['Route Name (no spaces)'].strip())
            device = get_device(device_name)
            device.static_routes[route.prefix] = route

def initalise_l2_interfaces():
    WORKSHEET_NAME = 'l2_interfaces'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            interface  = Interface()
            interface.type = 'layer2'
            device_name           = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            interface.name        = str(DATA[WORKSHEET_NAME][row_no]['Interface'].lower().strip())
            interface.enabled     = str(DATA[WORKSHEET_NAME][row_no]['Interface Enabled (yes/no)'].strip())
            interface.speed       = str(DATA[WORKSHEET_NAME][row_no]['Speed'].strip())
            interface.duplex      = str(DATA[WORKSHEET_NAME][row_no]['Duplex'].strip())
            interface.mtu         = str(DATA[WORKSHEET_NAME][row_no]['MTU'].strip())
            interface.description = str(DATA[WORKSHEET_NAME][row_no]['Description'])
            interface.variable1   = str(DATA[WORKSHEET_NAME][row_no]['Variable 1'].strip())
            interface.variable2   = str(DATA[WORKSHEET_NAME][row_no]['Variable 2'].strip())
            interface.pc_group    = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Group No'])
            interface.pc_mode     = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Mode (active/on/etc)'].strip())
            interface.pc_members  = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Members (separated by commas)'].strip())
            interface.data_vlan   = str(DATA[WORKSHEET_NAME][row_no]['Data VLAN'].strip())
            interface.voice_vlan  = str(DATA[WORKSHEET_NAME][row_no]['Voice VLAN'].strip())
            interface.native_vlan = str(DATA[WORKSHEET_NAME][row_no]['Trunk Native VLAN'])
            interface.trunk_vlans  = DATA[WORKSHEET_NAME][row_no]['Trunk Allowed VLANs (separated by commas)']
            if interface.pc_members:
                interface.pc_members = interface.pc_members.replace(' ','').lower().split(',')
            device = get_device(device_name)
            device.interfaces[interface.name] = interface

def initalise_l3_interfaces():
    WORKSHEET_NAME = 'l3_interfaces'
    for row_no, row in enumerate(DATA[WORKSHEET_NAME]):
        if valid_row(WORKSHEET_NAME,row_no):
            interface  = Interface()
            interface.type = 'layer3'
            device_name           = str(DATA[WORKSHEET_NAME][row_no]['Device Name'].lower().strip())
            interface.name        = str(DATA[WORKSHEET_NAME][row_no]['Interface'].lower().strip())
            interface.enabled     = str(DATA[WORKSHEET_NAME][row_no]['Interface Enabled (yes/no)'].strip())
            interface.speed       = str(DATA[WORKSHEET_NAME][row_no]['Speed'].strip())
            interface.duplex      = str(DATA[WORKSHEET_NAME][row_no]['Duplex'].strip())
            interface.mtu         = str(DATA[WORKSHEET_NAME][row_no]['MTU'].strip())
            interface.description = str(DATA[WORKSHEET_NAME][row_no]['Description'])
            interface.variable1   = str(DATA[WORKSHEET_NAME][row_no]['Variable 1'].strip())
            interface.variable2   = str(DATA[WORKSHEET_NAME][row_no]['Variable 2'].strip())
            interface.pc_group    = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Group No'])
            interface.pc_mode     = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Mode (active/on/etc)'].strip())
            interface.pc_members  = str(DATA[WORKSHEET_NAME][row_no]['Port-Channel Members (separated by commas)'].strip())
            interface.vrf         = str(DATA[WORKSHEET_NAME][row_no]['VRF (leave blank if global)'].strip())
            interface.ipaddress   = str(DATA[WORKSHEET_NAME][row_no]['IP Address (x.x.x.x/x)'].strip())
            if interface.pc_members:
                interface.pc_members = interface.pc_members.replace(' ','').lower().split(',')
            device = get_device(device_name)
            device.interfaces[interface.name] = interface


def initalise_portchannels():
    new_interfaces = {}
    for device in sorted(d.devices):
        for interface in sorted(d.devices[device].interfaces):
            intf = get_interface(device,interface)
            if intf.is_pc_parent:
                intf.comment = '!- member interfaces: {}'.format(','.join(intf.pc_members))
                for member in intf.pc_members:
                    member_intf = get_interface(device,member)
                    # member interface does not exist, create it
                    if not member_intf:
                        new_interface = Interface()
                        new_interface.name = member
                        new_interface.enabled = intf.enabled
                        new_interface.speed = ''
                        new_interface.duplex = ''
                        new_interface.pc_type = intf.type
                        new_interface.type = intf.type
                        new_interface.mtu = intf.mtu
                        new_interface.pc_parent = intf.name
                        new_interface.pc_group = intf.pc_group
                        new_interface.pc_mode = intf.pc_mode
                        if not new_interfaces.get(device):
                            new_interfaces[device] = {}
                        new_interfaces[device][new_interface.name] = new_interface
                    # member interface exist, update attributes
                    else:
                        member_intf.pc_parent = intf.name
                        member_intf.pc_type = intf.type
                        member_intf.pc_group = intf.pc_group
                        member_intf.pc_mode = intf.pc_mode
    # go ahead and create the new pc members that have not been created manually by the user
    for device in new_interfaces:
        for new in new_interfaces[device]:
            d.devices[device].interfaces[new] = new_interfaces[device][new]

#------------------------------------------
# Useful functions
#------------------------------------------
def get_device(device_name):
    return d.devices.get(device_name)

def get_interface(device_name,interface_name):
    return d.devices[device_name].interfaces.get(interface_name)

def get_variable(variable_name):
    return d.variables.get(variable_name)

def get_vrf(device_name,vrf_name):
    return d.devices[device_name].vrfs.get(vrf_name)

def get_route(device_name,route_prefix):
    return d.devices[device_name].static_routes.get(route_prefix)

def get_template(device_name,template_name):
    return d.devices[device_name].templates[template_name]

def get_template_list(device_name,config_position):
    device = get_device(device_name)
    template_list = []
    for template in sorted(device.templates):
        t = get_template(device_name,template)
        if config_position in t.position:
            template_list.append(t.name)
    return template_list

#------------------------------------
# Initalise the primary database
#------------------------------------
d = Database()
#------------------------------------
# Read the data from the spreadsheet
#------------------------------------
def initalise_data(filename):
    read_database_from_file(filename)
    initalise_devices()
    initalise_variables()
    initalise_config_templates()
    initilise_device_templates()
    initalise_vlans()
    initalise_vrfs()
    initalise_l2_interfaces()
    initalise_l3_interfaces()
    initalise_portchannels()
    initalise_static_routes()
