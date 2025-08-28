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
            except:
                print('System with incorrect sheet format: ')
                print(system, system_data)
            for value in hosts:
                try:
                    hostname = self._clean_hostname(value)
                except:
                    print ('Bad hostname check for upper and lower case and spliting at period: ')
                    print(system, value)
                
                if not hostname:
                    continue
                    
                if HOSTNAME_REGEX.match(hostname) or IP_ADDR_REGEX.match(hostname):
                    systems[system].add(hostname)
                else:
                    bad_hostnames.append((system, hostname))

        return systems, bad_hostnames

    def process_inventory(self, inventory_xlsx_path, output_directory):
        output_path = Path(output_directory)
        
        inventory_data = self._load_workbook(inventory_xlsx_path)
        host_data, bad_hostnames = self._extract_hosts(inventory_data)

        output_data = []
        for system, hosts in host_data.items():
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
                    'hostname': host
                })

        output_df = pd.DataFrame(output_data)
        output_df.to_csv(output_path/'host_data.csv', index=False)
        
        with open(output_path/'bad_hostnames.csv', 'w') as outfile:
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