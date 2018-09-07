# *-* coding: utf-8 *-*

"""
    Download current NIFTY indexes documents
"""

import os
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(1, ROOT_DIR)
for path, dirs, files in os.walk(ROOT_DIR):
    if '__pycache__' in path:
        continue
    sys.path.insert(1, path)

from etc.config import broad_indices, path_files
from lib.scrap import page_links, download, url_basename

links = page_links(url=broad_indices)
csv_links = list()

for link in links:
    csv_link = page_links(url=link)
    csv_link = [_ for _ in csv_link if _.strip() and _.endswith('.csv') and 'nifty' in _]
    if csv_link:
        csv_link = csv_link[0]
        print('Downloading ...', csv_link)
        path = os.path.join(path_files, url_basename(csv_link))
        download(csv_link, path)
        csv_links.append(csv_link[0])

