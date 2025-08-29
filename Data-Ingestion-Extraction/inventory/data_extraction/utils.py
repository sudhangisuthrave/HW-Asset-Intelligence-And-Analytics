"""Utility functions used by multiple files"""

import re
import string

from typing import List, Union

EXCLUDED_HOSTNAMES = [
    "bigiploadbalancer",
    "ciscoswitch",
    "ciscol2aggregator",
    "ciscoraccessrouter",
    "windowsserver",
    "nasaas",
    "na"
]

HOSTNAME_REGEX = re.compile(r'^[A-Za-z0-9_-]*$')
IP_ADDR_REGEX = re.compile(r'^((\d+)\.){3}(\d+)$')

PRINTABLE = set(string.printable)


def remove_parens(host: str) -> str:
    """Remove parentheses, brackets, and anything between them

    Parameters
    ----------
    host: str
        hostname

    Returns
    -------
    str
        hostname without parentheses or brackets
    """
    hostname = re.sub(r"[(\[].*?[)\]]", "", host)
    return hostname.strip()


def split(names: List[str], separator: str) -> List[str]:
    """

    Parameters
    ----------
    names: List[str]
        list of hostnames to be separated

    separator: str
        separator to use

    Returns
    -------
    List
        list of hostnames after splitting previous hostnames by separator
    """
    new_names = []

    for name in names:
        new_names.extend(name.split(separator))

    return new_names


def clean_hostname(hostname: str,
                   header_field: Union[str, None] = "hostname") -> List[str]:
    """Clean up hostnames by stripping white space, by splitting on tabs and
    carriage returns (and spaces, in certain cases), by ensuring IP addresses
    and certain characters are excluded

    Parameters
    ----------
    hostname: str
        candidate hostname to be cleaned

    header_field: str
        name of the field in tabular data from which the hostname was extracted,
        when header_field == 'cname', hostnames will be split on spaces; this
        is based on the format of certain inventory files

    Returns
    -------
    List[str]
        list of hostnames as strings, possibly an empty list

    """
    # ensure hostname is a string and
    # remove leading and trailing white space
    hostname = str(hostname).strip()

    # remove url components
    hostname = hostname.replace('https://', '')
    hostname = hostname.replace('http://', '')
    hostname = hostname.strip('/')

    # split  line returns
    hostnames = hostname.splitlines()

    # split on commas
    hostnames = split(hostnames, ',')

    # split on tabs
    hostnames = split(hostnames, '\t')

    # if header_field is 'cname' split on spaces
    if header_field == "cname":
        hostnames = split(hostnames, ' ')

    # remove text in parentheses/brackets
    hostnames = [remove_parens(x) for x in hostnames]

    # remove non-ascii characters
    hostnames = [''.join(filter(lambda x: x in PRINTABLE, y))
                 for y in hostnames]

    # replace spaces and strip whitespace from beginning and end
    hostnames = [x.replace(' ', '').strip().lower() for x in hostnames]

    # if hostname is an IP address, ignore
    hostnames = [x for x in hostnames if not IP_ADDR_REGEX.match(x)]

    # if hostname is given as a FQDN, collect only the first part
    hostnames = [x.split('.', 1)[0] if '.' in x else x for x in hostnames]

    # exclude certain hostnames
    hostnames = [x for x in hostnames
                 if x not in EXCLUDED_HOSTNAMES and HOSTNAME_REGEX.match(x)]

    # remove empty entries
    hostnames = [x for x in hostnames if x]

    return list(set(hostnames))
