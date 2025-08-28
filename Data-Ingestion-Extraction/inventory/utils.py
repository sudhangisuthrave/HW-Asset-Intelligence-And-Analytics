"""utility functions"""
import os
from pathlib import Path
from typing import Dict


def update_executable_path(config: Dict) -> None:
    """
    Update the system PATH environment variable with paths to various
    versions of the Chrome webdriver for use with Selenium.

    Parameters
    ----------
    config:
    """
    driver_path = Path(config['scraping']['webdriver_path'])

    for directory in driver_path.iterdir():
        if directory.is_dir():
            new_path = str(directory.resolve())
            os.environ['PATH'] += f";{new_path}"


def prepare_download_path(config: Dict) -> None:
    """Create the download path if necessary

    Parameters
    ----------
    config: Dict
        dictionary with configuration data
    """
    download_path = Path(config['scraping']['download_path'])
    download_path.mkdir(parents=True, exist_ok=True)
