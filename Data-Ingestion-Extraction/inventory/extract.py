"""Simplified interface for data extraction modules"""

import logging

from pathlib import Path
from typing import List

from .data_extraction import word, zip_archive


# data extractors by file extension
EXTRACTORS = {
    '.doc': word.extract_hostnames,
    '.docx': word.extract_hostnames,
    '.pdf': word.extract_hostnames,
    '.xls': lambda x: [],
    '.xlsx': lambda x: [],
    '.zip': zip_archive.extract_hostnames
}


def extract_hostnames(file_path: str) -> List[str]:
    """Extract hostnames from a file; this function is a simplified interface
    to the classes and functions found in the data extraction folder/package

    Parameters
    ----------
    file_path: str
        path to the file from which hostnames should be extracted

    Returns
    -------
    List[str]
        a list of hostnames, possibly an empty list if no hostnames were found
    """
    path = Path(file_path)
    extension = path.suffix.lower()

    if extension not in EXTRACTORS.keys():
        logging.warning("Unknown file type: %s (%s).", extension, file_path)
        return []

    results = EXTRACTORS[extension](file_path)

    # zip extraction will not return hostnames but a list of files that must
    # be processed further
    if extension == ".zip":
        hostnames = []

        for result_path in results:
            names = extract_hostnames(result_path)
            hostnames.extend(names)

        results = hostnames

    return results
