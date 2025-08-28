"""Primary CSAM-related functionality for scraping"""
import logging
import os
import time

from enum import Enum
from pathlib import Path
from typing import Dict, List, Set
from urllib.parse import urljoin

import pandas as pd

from bs4 import BeautifulSoup
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver import ChromeOptions
from splinter import Browser
from splinter.driver.webdriver import WebDriverElement

from csam_inventory.log import LoggingBase

LOGIN_URL = "login.aspx"
LOGIN_USER_FIELD_NAME = "Login1$UserName"
LOGIN_PASS_FIELD_NAME = "Login1$Password"
LOGIN_BUTTON_ID = "Login1_LoginButton"
LOGIN_CONFIRM_ID = "ContentPlaceHolder1_btn_Accept"
MAIN_PAGE_SEARCH_ID = "ph_MainContentArea_lb_Search"
SYSTEM_SEARCH_URL = "System/Select.aspx"
SYSTEM_SEARCH_BUTTON_ID = "ctl00_ph_MainContentArea_btn_Run"

SYSTEM_SEARCH_QUANTITY_ID = (
    "ctl00_ph_MainContentArea_rg_Output_ctl00_ctl03_ctl01_PageSizeComboBox"
)

SYSTEM_SEARCH_QUALITY_ALL_SELECTOR = (
    "#ctl00_ph_MainContentArea_rg_Output_ctl00_ctl03_ctl01_"
    "PageSizeComboBox_DropDown > div > ul > li:nth-child(6)"
)

SYSTEM_PAGES_SELECTOR = (
    "#ctl00_ph_MainContentArea_rg_Output_ctl00 > tfoot > tr > td > "
    "div > div.rgWrap.rgInfoPart"
)

SYSTEM_TABLE_ID = "ctl00_ph_MainContentArea_rg_Output_ctl00"
SYSTEM_TABLE_HEADER_ROW = 2

SYSTEM_INPUT_ID = (
    "ctl00_ph_MainContentArea_rpb_SearchParameters_i0_rlv_"
    "Parms_ctrl1_tb_SystemID")

SYSTEM_SEARCH_RESULT_SELECTOR = (
    "#ctl00_ph_MainContentArea_rg_Output_ctl00__0 > td:nth-child(1) > a"
)

SYSTEM_APPENDICES_SELECTOR = (
    "#ctl00_ctl00_ph_MainContentArea_SystemMasterTabs1_RadPanelBar_SystemMain "
    "> ul > li:nth-child(2) > div > ul > li:nth-child(5) > a"
)

SYSTEM_APPENDIX_LINK_ID = (
    "ctl00_ctl00_ph_MainContentArea_ContentPlaceHolder1_gv_Appendices_"
    "ctl00_ctl70_hlViewArt"
)

ID_ORG_ACRONYM_FILE_NAME = "id-org-acronym-diff.csv"


class Selector(Enum):
    """Various selectors for use to find objects on a page"""
    CSS = 0
    ID = 1
    NAME = 2
    TEXT = 3


class CsamScraper(LoggingBase):
    """CSAM scraping functionality"""

    def __init__(self, config: Dict) -> None:
        """Initialize an instance of the CsamScraper class

        Parameters
        ----------
        config: Dict
            dictionary containing configuration data
        """
        self._config = config
        self._browser = self._create_browser()
        self._logged_in = False
        super().__init__()

    def _create_browser(self) -> Browser:
        """Create an instance of a Browser object which creates and controls
        a new browser instance.

        Returns
        -------
        Browser
            an instance of the Selenium/Splinter Browser instance
        """
        download_path = Path(self._config['scraping']['download_path']).resolve()

        prefs = {
            "download.default_directory": str(download_path),
            "download.directory_upgrade": "true",
            "download.prompt_for_download": "false",
            "disable-popup-blocking": "true"
        }

        options = ChromeOptions()
        options.add_experimental_option("prefs", prefs)
        options.add_argument('log-level=3')
        return Browser('chrome', options=options)

    def _wait_for_element(self, selector: Selector,
                          value: str,
                          matching_text: str = None) -> WebDriverElement:
        """Wait until a specified DOM element is found on a web page

        Parameters
        ----------
        selector: Selector
            type of identifier used to locate an object (id, name, etc.)

        value: str
            the desired value of the chosen selector used for locating an
            element (element id, element name, etc.)

        matching_text: str
            if specified, use this value when locating an element and ensure
            the element's text matches this value

        Returns
        -------
        WebDriverElement
            the desired element on the page; if no element is found within a
            timeout period (set in the configuration file), a TimeoutError will
            be raised
        """
        locators = {
            Selector.CSS: self._browser.find_by_css,
            Selector.ID: self._browser.find_by_id,
            Selector.NAME: self._browser.find_by_name,
        }

        """ Changed this on 2/8/2022 
        timeout = self._config['scraping']['load_timeout']
        timeout_time = time.time() + timeout"""
        timeout_time = time.time() + 240
        element_found = False
        results = None

        while not element_found:
            results = locators[selector](value)
            element_found = len(results) > 0

            if matching_text and element_found:
                try:
                    element_found = (matching_text in results.first.text)
                except StaleElementReferenceException:
                    # wait till page fully refreshes to get element
                    continue

            if time.time() > timeout_time:
                raise TimeoutError("Could not find element with"
                                   f" {selector}={value}.")

        return results.first

    def _wait_for_download(self, files_before_download: Set[str]) -> Path:
        """Wait for file download to complete

        Parameters
        ----------
        files_before_download: Set[str]
            set of strings representing the files in the download path before
            the new file is downloaded; a change in this file set indicates
            that a new file has been downloaded

        Returns
        -------
        Path
            a Path object representing the location of the newly downloaded
            file
        """
        download_path = Path(self._config['scraping']['download_path'])
        pre_download_set = set(files_before_download)
        current_file_set = set(os.listdir(download_path))
        """timeout = self._config['scraping']['download_timeout']
        timeout_time = time.time() + timeout"""
        timeout_time = time.time() + 60
        new_file = ""

        # wait for a new file that doesn't end in crdownload or tmp to appear
        # in the download directory
        while (pre_download_set == current_file_set
               or new_file.endswith("crdownload")
               or new_file.endswith("tmp")):
            if time.time() > timeout_time:
                raise TimeoutError("Download failed or did not complete.")

            time.sleep(0.1)
            current_file_set = set(os.listdir(download_path))
            difference = current_file_set - pre_download_set

            if len(difference) > 0:
                new_file = (current_file_set - pre_download_set).pop()

        return download_path / new_file

    def login(self) -> None:
        """Log into CSAM using credentials provided in the configuration file"""
        url = urljoin(str(self._config['csam']['base_url']), LOGIN_URL)
        logging.info("Logging into CSAM at %s.", url)

        self._browser.visit(url)

        username = self._wait_for_element(Selector.NAME, LOGIN_USER_FIELD_NAME)
        username.fill(self._config['csam']['username'])

        password = self._wait_for_element(Selector.NAME, LOGIN_PASS_FIELD_NAME)
        password.fill(self._config['csam']['password'])

        login_button = self._wait_for_element(Selector.ID, LOGIN_BUTTON_ID)
        login_button.click()

        confirm_button = self._wait_for_element(Selector.ID, LOGIN_CONFIRM_ID)
        confirm_button.click()

        self._wait_for_element(Selector.ID, MAIN_PAGE_SEARCH_ID)
        self._logged_in = True
        logging.info("Login successful.")

    def retrieve_system_list(self) -> List[int]:
        """Collect a list of systems in CSAM and save that data to a CSV that
        will be used during export to match systems with orgs and acronyms

        Returns
        -------
        List[int]
            list of system IDs as integers
        """
        logging.info("Retrieving list of systems from CSAM.")

        if not self._logged_in:
            self.login()

        url = urljoin(str(self._config['csam']['base_url']), SYSTEM_SEARCH_URL)
        self._browser.visit(url)

        search = self._wait_for_element(Selector.ID, SYSTEM_SEARCH_BUTTON_ID)
        search.click()

        quantity = self._wait_for_element(Selector.ID,
                                          SYSTEM_SEARCH_QUANTITY_ID)

        quantity.click()

        option_all = self._wait_for_element(Selector.CSS,
                                            SYSTEM_SEARCH_QUALITY_ALL_SELECTOR)

        option_all.click()

        self._wait_for_element(
            Selector.CSS,
            SYSTEM_PAGES_SELECTOR,
            matching_text="items in 1 pages"
        )

        soup = BeautifulSoup(
            self._browser.driver.page_source,
            features="html5lib"
        )

        table = str(soup.find(id=SYSTEM_TABLE_ID))

        data, *_ = pd.read_html(
            table,
            flavor="html5lib",
            header=SYSTEM_TABLE_HEADER_ROW
        )

        ids = sorted([int(x) for x in data.ID.to_list() if x.isnumeric()])
        id_org_acronym = data[['ID', 'Org', 'Acronym']]  # type: pd.DataFrame

        id_org_acronym_path = (Path(self._config['scraping']['download_path'])
                               / ID_ORG_ACRONYM_FILE_NAME)

        id_org_acronym.to_csv(id_org_acronym_path, index=False)

        logging.info("Retrieved %d system IDs.", len(ids))

        return ids

    def collect_hardware_inventories(self, system_list: List[int]) -> None:
        """Download hardware inventories for multiple systems

        Parameters
        ----------
        system_list: List[int]
            list of system IDs
        """
        total = len(system_list)

        for i, system in enumerate(system_list):
            logging.info(
                "Downloading CSAM hardware inventory for system %s (%d/%d).",
                system, i+1, total
            )

            self._collect_hardware_inventory(system)

    def _collect_hardware_inventory(self, system_id: int) -> None:
        """Download the hardware inventory for a single system

        Parameters
        ----------
        system_id: int
            ID of the system whose hardware inventory should be downloaded
        """
        logging.debug("Retrieving CSAM hardware inventory for system %d.",
                      system_id)

        if not self._logged_in:
            self.login()

        url = urljoin(str(self._config['csam']['base_url']), SYSTEM_SEARCH_URL)
        self._browser.visit(url)

        system_input = self._wait_for_element(Selector.ID, SYSTEM_INPUT_ID)
        system_input.fill(system_id)
        search = self._wait_for_element(Selector.ID, SYSTEM_SEARCH_BUTTON_ID)
        search.click()

        result = self._wait_for_element(Selector.CSS,
                                        SYSTEM_SEARCH_RESULT_SELECTOR)

        result.click()

        appendix_list_link = self._wait_for_element(
            Selector.CSS,
            SYSTEM_APPENDICES_SELECTOR
        )

        appendix_list_link.click()

        download_path = Path(self._config['scraping']['download_path'])
        pre_download_file_set = set(os.listdir(download_path))

        try:
            hw_link = self._wait_for_element(Selector.ID,
                                             SYSTEM_APPENDIX_LINK_ID)
        except TimeoutError:
            # if HW inventory link is not found, the system likely doesn't have
            # an inventory file in CSAM
            logging.warning("No inventory found for system %d.", system_id)
            return

        hw_link.click()

        hw_file = self._wait_for_download(pre_download_file_set)

        # wait to rename file to avoid chrome download failure messages
        time.sleep(0.5)

        extension = hw_file.suffix
        hw_file.rename(download_path / f"hw-inventory-{system_id}{extension}")

        logging.debug("Hardware inventory for system %d downloaded.", system_id)

    def cleanup(self):
        """Close the Browser instance"""
        self._browser.quit()
