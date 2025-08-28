'''Convert CSAM HW inventory files to flat CSVs'''

import argparse
import re

from collections import defaultdict
from pathlib import Path

import pandas as pd


SYSTEM_ID_REGEX = re.compile(r'(\d+)')
HOSTNAME_REGEX = re.compile(r'^[A-Za-z0-9_-]*$')
IP_ADDR_REGEX = re.compile(r'^((\d+)\.){3}(\d+)$')


class HostExtractor:
    def __init__(self, org_mapping_path):
        self._org_mapping = self._load_org_data(org_mapping_path)

    @staticmethod
    def _load_org_data(org_mapping_path):
        mapping = {}
        data = pd.read_excel(org_mapping_path, engine='openpyxl')

        for _, row in data.iterrows():
            mapping[str(row['CSAM ID'])] = {
                'org': row['Org'],
                'acronym': row['Acronym']
            }

        return mapping

    @staticmethod
    def _load_workbook(file_path):
        xlsx = pd.ExcelFile(file_path, engine='openpyxl')
        data = pd.read_excel(xlsx, sheet_name=None)
        
        for dataframe in data.values():
            dataframe.dropna(how='all', inplace=True)
            dataframe.fillna('', inplace=True)
            
        return data

    @staticmethod
    def _clean_hostname(hostname):
        hostname = hostname.replace(' ', '').strip().lower()
        
        if IP_ADDR_REGEX.match(hostname):
            # print("IP address in hostname")
            # print(str(hostname))
            return str(hostname)
        elif '.' in hostname:
            hostname, _ = hostname.split('.', 1)
        
        return str(hostname)

    def _extract_hosts(self, workbook_data):
        systems = defaultdict(set)
        bad_hostnames = []
        
        for system, system_data in workbook_data.items():
            try: 
                hosts = system_data['Identifier or Host Name']
                ips = system_data['IP Address (Internal)']
                ips_ext = system_data['IP Address (External)']
                nat_ips_1 = system_data['NAT IPs']
                ad_domain_1 = system_data['AD Domain']
                cpu_core_1 = system_data['CPU Core']
                memory_1 = system_data['Memory (GB)']
                drive_space_1 = system_data['Drive Space (GB)']
                hva = system_data['High Value Asset']
                serial_num = system_data['Manufacturer Serial Number']
                mac_add = system_data['MAC Address(es)']
                bios_guid = system_data['BIOS UUID/GUID']
                asset_cat = system_data['Asset Category']
                asset_ty = system_data['Asset Type']
                virtual_1 = system_data['Virtual']
                public_1 = system_data['Public']
                gfe_1 = system_data['GFE']
                hw_make = system_data['Hardware Make']
                hw_model = system_data['Hardware Model']
                os_n = system_data['OS Name']
                os_v = system_data['OS Version']
                lifecycle_1 = system_data['Lifecycle']
                location_1 = system_data['Location']
                hosting_csp_contract_1 = system_data['Hosting/CSP Contract']
                date_device_added = system_data['Date Device Added to System Boundary']
                device_op = system_data['Device Operator']
                systems_supported_1 = system_data['Systems Supported CSAM Acronym']
                device_man = system_data['System Owner / Device Manager']
                primary_system_boundary_csam_id_1 = system_data['Primary System Boundary CSAM ID']
                primary_system_boundary_csam_acro_1 = system_data['Primary System Boundary CSAM Acronym']
                systems_supported_id_1 = system_data['System Supported CSAM ID']
                first_tier_supplier_1 = system_data['First Tier Supplier']

            except:
                print('System with incorrect sheet format: ')
                print(system, system_data)
            count=0
            for value in hosts:
                try:
                    hostname = self._clean_hostname(value)
                    ipaddress = ips[count]
                    ipaddress_external = ips_ext[count]
                    nat_ip_address = nat_ips_1[count]
                    ad_domain = ad_domain_1[count]
                    cpu_core = cpu_core_1[count]
                    memory = memory_1[count]
                    drive_space = drive_space_1[count]
                    high_value_asset = hva[count]
                    manufacturer_serial_number = serial_num[count]
                    mac_address = mac_add[count]
                    bios_uuid_guid = bios_guid[count]
                    asset_category = asset_cat[count]
                    asset_type = asset_ty[count]
                    virtual = virtual_1[count]
                    public = str(public_1[count])
                    gfe = gfe_1[count]
                    hardware_make = hw_make[count]
                    hardware_model = hw_model[count]
                    os_name = os_n[count]
                    os_version = os_v[count]
                    lifecycle = lifecycle_1[count]
                    location = location_1[count]
                    hosting_csp_contract = hosting_csp_contract_1[count]
                    date_device_added_to_system_boundary = date_device_added[count]
                    device_operator = device_op[count]
                    systems_supported = systems_supported_1[count]
                    device_manager = device_man[count]
                    primary_system_boundary_csam_id = primary_system_boundary_csam_id_1[count]
                    primary_system_boundary_csam_acro = primary_system_boundary_csam_acro_1[count]
                    systems_supported_id = systems_supported_id_1[count]
                    first_tier_supplier = first_tier_supplier_1[count]

                except:
                    print('Bad hostname check for upper and lower case and splitting at period: ')
                    print(system, value)
                
                if not hostname:
                    continue
                    
                if HOSTNAME_REGEX.match(hostname) or IP_ADDR_REGEX.match(hostname):
                    host_ip = tuple([hostname, ipaddress, ipaddress_external, nat_ip_address, ad_domain, cpu_core, memory, drive_space, high_value_asset, manufacturer_serial_number, mac_address,bios_uuid_guid, asset_category, asset_type, virtual, public, gfe, hardware_make, hardware_model, os_name, os_version, lifecycle, location, hosting_csp_contract, date_device_added_to_system_boundary, device_operator, systems_supported, device_manager, primary_system_boundary_csam_id, primary_system_boundary_csam_acro, systems_supported_id, first_tier_supplier])
                    # systems[system].add(hostname)
                    systems[system].add(host_ip)
                    host_ip = tuple()
                else:
                    bad_hostnames.append((system, hostname))
                count=count+1
        return systems, bad_hostnames

    def process_inventory(self, inventory_xlsx_path, output_directory):
        output_path = Path(output_directory)
        
        inventory_data = self._load_workbook(inventory_xlsx_path)
        host_ip_data, bad_hostnames = self._extract_hosts(inventory_data)

        output_data = []
        for system, hosts in host_ip_data.items():
            system_id = system.split('-')[-1]
            print('System ID is: ')
            print(system_id)
            org_acronym = self._org_mapping[system_id]

            for host in hosts:
                output_data.append({
                    'csam_id': system_id,
                    'org': org_acronym['org'],
                    'acronym': org_acronym['acronym'],
                    'id_acronym': f"{system_id}-{org_acronym['acronym']}",
                    'hostname': host[0],
                    'ip_address_internal': host[1],
                    'ip_address_external': host[2],
                    'nat_ips': host[3],
                    'ad_domain': host[4],
                    'cpu_core': host[5],
                    'memory': host[6],
                    'drive_space': host[7],
                    'high_value_asset': host[8],
                    'manufacturer_serial_number': host[9],
                    'mac_address': host[10],
                    'bios_uuid_guid': host[11],
                    'asset_category': host[12],
                    'asset_type': host[13],
                    'virtual': host[14],
                    'public': host[15],
                    'gfe': host[16],
                    'hardware_make': host[17],
                    'hardware_model': host[18],
                    'os_name': host[19],
                    'os_version': host[20],
                    'lifecycle': host[21],
                    'location': host[22],
                    'hosting_csp_contract': host[23],
                    'date_device_added_to_system_boundary': host[24],
                    'device_operator': host[25],
                    'systems_supported_csam_acronym': host[26],
                    'device_manager': host[27],
                    'primary_system_boundary_csam_id': host[28],
                    'primary_system_boundary_csam_acro': host[29],
                    'systems_supported_id': host[30],
                    'first_tier_supplier': host[31]
                })
        output_df = pd.DataFrame(output_data)
        output_df.to_csv(output_path/'generate_hosts.csv', index=False)
        
        # with open(output_path/'generate_hosts_bad_hostnames.csv', 'w') as outfile:
        with open(output_path / 'generate_hosts_bad_hostnames.csv', 'w', encoding='utf-8') as outfile:
            outfile.writelines([f"{','.join(x)}\n" for x in bad_hostnames])


def generate_hosts_file(org_mapping_path, inventory_xlsx_path, 
        output_directory):
    extractor = HostExtractor(org_mapping_path)
    extractor.process_inventory(inventory_xlsx_path, output_directory)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Process CSAM inventory files")
    
    parser.add_argument(
        'org_mapping_path', 
        help=(
            "path to an XLSX file that lists organization and acronym for each "
            "CSAM ID"
        )
    )
    
    parser.add_argument(
        'inventory_xlsx_path', 
        help=(
            "path to an XLSX file that lists host inventory for each system "
            "in separate worksheets"
        )
    )
    
    parser.add_argument(
        'output_directory', 
        help="directory to which output files will be written"
    )
    
    args = parser.parse_args()
    
    generate_hosts_file(
        args.org_mapping_path, 
        args.inventory_xlsx_path, 
        args.output_directory
    )