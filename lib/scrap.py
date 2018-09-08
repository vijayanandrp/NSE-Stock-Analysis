# *-* coding: utf-8 *-*

from requests_html import HTMLSession
import requests
from urllib.parse import urlparse, quote
import os

headers = {'User-Agent': 'Mozilla/5.0'}


def _request(url):
    url = quote(url.strip(), safe='/:_.?=')
    session = HTMLSession(mock_browser=True)
    r = session.get(url, headers=headers)
    return r


def page_links(url):
    r = _request(url)
    return r.html.absolute_links


def page_html(url):
    r = _request(url)
    return r.html.raw_html


def _url_extract(url):
    return urlparse(url)


def url_basename(url):
    a = _url_extract(url)
    return os.path.basename(a.path)


def download(url, path=None):
    r = requests.get(url, headers=headers)
    with open(path, 'wb') as f:
        f.write(r.content)

    # Retrieve HTTP meta-data
    print(r.status_code)
    print(r.headers['content-type'])
    print(r.encoding)
    if os.path.isfile(path):
        print('Download success')
    else:
        print('Download failed')

