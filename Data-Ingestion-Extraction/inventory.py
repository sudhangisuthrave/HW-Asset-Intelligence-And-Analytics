"""Collect and process CSAM hardware inventory

This script serves as the entry point into the program when running
python inventory.py.
"""

import argparse
import logging
import logging.config

from csv import reader
from pathlib import Path
from typing import Dict, List

import yaml

from csam_inventory import csam, extract, utils

def download_csam_inventories(config: Dict) -> None:
    """Scrape the CSAM site to download hardware inventory files

    Parameters
    ----------
    config: dict
        dictionary with configuration data, usually loaded from config.yml
    """
    scraper = csam.CsamScraper(config)
    scraper.login()
    system_ids = scraper.retrieve_system_list()
    scraper.collect_hardware_inventories(system_ids)


def process_inventories(config: Dict) -> Dict[int, List[str]]:
    """Process hardware inventory files to extract hostnames

    Parameters
    ----------
    config: dict
        dictionary with configuration data, usually loaded from config.yml

    Returns
    -------
    dict[int, list[str]]
        dictionary with system IDs as keys and a list of system hostnames as
        the corresponding values
    """
    results = {}
    output_path = Path(config['scraping']['download_path'])

    for item in output_path.iterdir():
        if (not item.is_file()
                or item.name.lower() == csam.ID_ORG_ACRONYM_FILE_NAME):
            continue

        *_, system_id = item.stem.split('-')
        hostnames = extract.extract_hostnames(str(item.resolve()))
        results[system_id] = hostnames

    return results


def export_inventories(config: Dict, inventories: Dict[int, List[str]],
                       output_name: str) -> None:
    """Export system inventories to a single file

    Parameters
    ----------
    config: dict
        dictionary with configuration data, usually loaded from config.yml

    inventories: dict[int, list[str]]
        dictionary with system IDs as keys and a list of system hostnames as
        the corresponding values

    output_name: str
        name of CSV file to be created and containing inventory data
    """
    id_org_acronym_path = (Path(config['scraping']['download_path'])
                           / csam.ID_ORG_ACRONYM_FILE_NAME)

    id_org_acronym = {}

    with open(id_org_acronym_path) as csv_file:
        csv_reader = reader(csv_file)

        for row in csv_reader:
            id_, org, acronym = row  # type: (str, str, str)

            if not id_.isnumeric():
                continue

            id_org_acronym[int(id_)] = (org, acronym)

    with open(output_name, 'w') as out_file:
        out_file.write("csam_id,org,acronym,id_acronym,hostname\n")

        for system_id, hostnames in inventories.items():
            org, acronym = id_org_acronym.get(int(system_id), ("", ""))

            for host in sorted(hostnames):
                out_file.write(f"{system_id},{org},{acronym},"
                               f"{system_id}-{acronym},{host}\n")


def main(skip_download: bool = False,
         config_path: str = "./config.yml") -> None:
    """Main function for inventory collection; calls other functions for each
    step of the process

    Parameters
    ----------
    skip_download: bool
        if true, skip download of hardware inventories and use what is already
        in the download path; if false, download hardware inventories;
        default: False

    config_path: str
        path to the configuration file; default: './config.yml'
    """
    config_path = Path(config_path).resolve()
    with open(config_path, 'r') as config_file:
        config = yaml.load(config_file, Loader=yaml.SafeLoader)

    logging.config.dictConfig(config['logging'])
    logging.info("Config file loaded from %s.", str(config_path.resolve()))

    utils.update_executable_path(config)
    utils.prepare_download_path(config)

    if not skip_download:
        logging.info("Beginning download of hardware inventory.")
        download_csam_inventories(config)

    logging.info("Beginning processing of hardware inventory.")
    inventories = process_inventories(config)

    logging.info("Exporting consolidated hardware inventory.")
    export_inventories(config, inventories, 'hostnames.csv')
    logging.info("Consolidated inventory exported to hostnames.csv.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser("Collect CSAM hardware inventory")

    parser.add_argument(
        '--skip-download',
        action="store_true",
        help="skip downloading new hardware inventory files"
    )

    parser.add_argument(
        "--config",
        help="path to configuration file, default is ./config.yml",
        default="./config.yml"
    )

    args = parser.parse_args()
    main(args.skip_download, args.config)
