"""Extract files from ZIP files for further processing """

import logging

from pathlib import Path
from typing import List
from zipfile import ZipFile


def extract_hostnames(zip_path: str) -> List[Path]:
    """Extract files from a ZIP file for further processing

    Parameters
    ----------
    zip_path: str
        path to a ZIP file from which other files will be extracted

    Returns
    -------
    List[Path]
        a list of Path objects representing the location of the newly
        extracted files
    """
    logging.info("Extracting files from %s.", zip_path)

    extracted_files = []
    file_path = Path(zip_path)
    file_stem = file_path.stem
    out_dir = file_path.parent

    with ZipFile(file_path) as zip_data:
        for i, file in enumerate(zip_data.filelist):
            orig_path = out_dir / file.filename
            zip_data.extract(file, out_dir)
            prefix, system_id = file_stem.rsplit('-', 1)

            # extracted files will have names in the form of
            # hw-inventory-X-ID.SUFFIX
            # where X is the files index within the zip file,
            # ID is the system ID extracted from the zip file name,
            # and SUFFIX is the original suffix of the extracted file.
            renamed_path = (
                    out_dir / f"{prefix}-{i}-{system_id}{orig_path.suffix}"
            )

            orig_path.replace(renamed_path)
            extracted_files.append(renamed_path)

    logging.info("Extracted %d {len(extracted_files)} files from %s.",
                 len(extracted_files), zip_path)

    return extracted_files
