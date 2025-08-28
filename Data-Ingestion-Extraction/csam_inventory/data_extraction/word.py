"""Extract data from word and pdf files"""
from typing import List, Tuple, Union

import pywintypes
import win32com.client as win32

from ..log import LoggingBase
from .utils import clean_hostname

# list of possible headers for column containing hostnames, 'hostname' should
# be the last element so other can be matched first
HEADERS = ['cname', 'hostname', 'identifier']


class WordProcessor(LoggingBase):
    """Extract data from Word and PDF documents"""

    def __init__(self) -> None:
        """Initialize an instance of the WordProcessor class"""
        self._word = win32.gencache.EnsureDispatch('Word.Application')
        self._word.Visible = 0
        super().__init__()

    def extract_hosts(self, doc_path: str) -> List[str]:
        """Extract hostnames from a Word or PDF document

        Parameters
        ----------
        doc_path: str
            path to Word or PDF document

        Returns
        -------
        List[str]
            a list of hostnames as strings
        """
        hostnames = self._extract_hosts_from_tables(doc_path)

        if not hostnames:
            hostnames = self._extract_hosts_from_text(doc_path)

        return hostnames

    def _extract_hosts_from_tables(self, doc_path: str) -> List[str]:
        """Extract hostnames from tables within a Word or PDF document

        Parameters
        ----------
        doc_path: str
            path to a Word or PDF document

        Returns
        -------
        List[str]
            a list of hostnames as strings
        """
        self.logger.info("Extracting hosts from tables in %s.", doc_path)
        hostnames = []  # type: List[str]
        previous_host_column = 0
        previous_hostname_header = None
        doc = None

        try:
            doc = self._word.Documents.Open(doc_path)
            self._word.Visible = 0

            # iterate through tables
            for table_index in range(1, doc.Tables.Count + 1):
                table = doc.Tables(table_index)

                header_row, host_column, hostname_header = (
                    self._find_hostname_column(table)
                )

                # no results for hostname column and no previous results
                if not host_column and not previous_host_column:
                    continue

                if hostnames and previous_host_column:
                    # if next table doesn't have enough columns, continue
                    if (table.Columns.Count <= previous_host_column
                            or table.Rows.Count == 0):
                        continue

                    # if the first row isn't as long as the rest of the table
                    # or contains merged cells, look at the second row
                    try:
                        test_text = (
                            table.Cell(Row=1, Column=previous_host_column)
                            .Range.Text
                        )

                    except pywintypes.com_error:
                        test_text = (
                            table.Cell(Row=2, Column=previous_host_column)
                            .Range.Text
                        )

                    test_hostnames = clean_hostname(test_text,
                                                    previous_hostname_header)

                    if not test_hostnames:
                        continue

                    test_hostname, *_ = test_hostnames

                    # compare the beginning of the first element of new table
                    # with the beginning of the last hostname collected, if they
                    # are the same, treat this table as a continuation of the
                    # previous table
                    if test_hostname[:4] == hostnames[-1][:4]:
                        header_row = 0
                        host_column = previous_host_column
                        hostname_header = previous_hostname_header
                    else:
                        previous_host_column = 0
                        previous_hostname_header = None
                        continue

                # iterate through rows
                for row_index in range(header_row + 1, table.Rows.Count + 1):
                    try:
                        text = (
                            table.Cell(Row=row_index, Column=host_column)
                            .Range.Text
                        )

                        found_hostnames = clean_hostname(text, hostname_header)
                        hostnames.extend(found_hostnames)

                    except pywintypes.com_error:
                        # merged cells across rows will cause an exception
                        continue

                # keep these values in case we need to read from the next table
                previous_host_column = host_column
                previous_hostname_header = hostname_header

        except Exception as exp:
            self.logger.exception(exp)
            raise exp

        finally:
            try:
                if hasattr(doc, 'Close'):
                    doc.Close()
            except Exception:
                self.logger.error("Unable to close document at %s.", doc_path)

        hostnames = list(set(hostnames))
        self.logger.info("Found %d hostnames in tables in %s.",
                         len(hostnames), doc_path)

        return hostnames

    def _find_hostname_column(self, table: object) -> \
            Union[Tuple[int, int, str], Tuple[None, None, None]]:
        """Locate the column in a table with a header in the HEADERS list

        Parameters
        ----------
        table: object
            table extracted from a Word or PDF document

        Returns
        -------
        int, int, int or None, None, None
            the row index, column index, and header name for the hostname
            header; if the header cannot be found, None, None, None is returned
        """
        # return row and column index of hostname header,
        # as well as the name of the hostname header (or an alternative)
        for row_index in range(1, table.Rows.Count + 1):
            for column_index in range(1, table.Columns.Count + 1):
                try:
                    cell_text = (table.Cell(Row=row_index, Column=column_index)
                                 .Range.Text.lower().replace(" ", ""))

                    # system 602 contains both cname and hostname columns
                    # use the cname column and split entries by tab characters
                    # or spaces
                    for header in HEADERS:
                        if header in cell_text:
                            self.logger.debug(
                                "Found header %s at (%d, %d).",
                                header,
                                row_index,
                                column_index
                            )

                            return row_index, column_index, header

                except pywintypes.com_error:
                    # iterating through tables with merged cells will
                    # throw exceptions
                    continue

        return None, None, None

    def _extract_hosts_from_text(self, doc_path: str) -> List[str]:
        """Extract hostnames from text within a Word or PDF document; this
        should be used if hostnames could not be extracted from tabular data

        Parameters
        ----------
        doc_path: str
            path to a Word or PDF document

        Returns
        -------
        List[str]
            a list of hostnames as strings, possibly an empty list
        """
        self.logger.info("Extracting hosts from text in %s.", doc_path)
        hostnames = []
        doc = None

        try:
            doc = self._word.Documents.Open(doc_path)
            self._word.Visible = 0

            for paragraph_index in range(1, doc.Paragraphs.Count):
                text = (
                    doc.Paragraphs(paragraph_index).Range.Text.lower().strip()
                )

                if "device name:" in text:
                    names = clean_hostname(text.split()[-1])
                    hostnames.extend(names)

        except Exception as exp:
            self.logger.exception(exp)
            raise exp

        finally:
            if doc is not None:
                doc.Close()

        hostnames = list(set(hostnames))
        self.logger.info("Found %d hostnames in text in %s.",
                         len(hostnames), doc_path)

        return hostnames


def extract_hostnames(file_path: str) -> List[str]:
    """Helper function used to extract hostnames from a Word or PDF document;
    this can be used in place of code that creates an instance of WordProcessor
    and then works with that instance

    Parameters
    ----------
    file_path: str
        path to a Word or PDF document

    Returns
    -------

    """
    processor = WordProcessor()
    return processor.extract_hosts(file_path)
