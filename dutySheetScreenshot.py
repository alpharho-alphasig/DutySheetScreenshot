#!/usr/bin/env python3

from urllib.parse import quote as urlencode
from xml.etree import ElementTree as ET
import requests
from tempfile import NamedTemporaryFile
from subprocess import call
from sys import exit
from discord_webhook import DiscordWebhook

dutiesPath = "/Shared/NewDrive/ALPHA SIG GENERAL/02_COMMITTEES/07_HOUSING/WEEKLY DUTIES/2023_SPRING/"

with open("/opt/bots/password", "r") as passwordFile:
    password = passwordFile.readline().strip()

with open("/opt/bots/DutySheetScreenshot/URLs/nextcloudURL.txt", "r") as URLFile:
    ncURL = URLFile.readline().strip()

dutiesPath = urlencode(dutiesPath)

# See https://docs.nextcloud.com/server/19/developer_manual/client_apis/WebDAV/basic.html for info on the NC API.
dutiesResponse = requests.request(
    method="PROPFIND",
    url=ncURL + "remote.php/dav/files/bot" + dutiesPath,
    auth=("bot", password),
    data="""<?xml version="1.0" encoding="UTF-8"?>
  <d:propfind xmlns:d="DAV:">
    <d:prop xmlns:oc="http://owncloud.org/ns">
      <oc:fileid/>
    </d:prop>
  </d:propfind>"""
)
response = dutiesResponse.text
responseXML = ET.fromstring(response)
fileIDs = {}
files = responseXML.findall("{DAV:}response")
for file in files:
    path = file.find("{DAV:}href").text
    for child in file.find("{DAV:}propstat").iter():
        if child.tag.endswith("fileid"):
            fileID = child.text
            break
    fileIDs[fileID] = path[25:]

fileID = max(fileIDs.keys())
newestDutySheet = fileIDs[fileID]

# Check if the file is the same as last week, if so, exit, otherwise, update the file.
with open("/opt/bots/DutySheetScreenshot/lastFileID", "w+") as lastIDFile:
    lastID = lastIDFile.read().strip()
    if fileID == lastID:
        print("Same as last week. Closing.")
        exit(0)

    lastIDFile.write(fileID)

xlsxFile = NamedTemporaryFile()
pdfFile = xlsxFile.name + ".pdf"
pngFile = NamedTemporaryFile()
content = requests.get(
    url=ncURL + "remote.php/dav/files/bot" + newestDutySheet,
    auth=("bot", password)
).content
xlsxFile.write(content)

call("libreoffice --headless --convert-to pdf {} --outdir /tmp".format(xlsxFile.name), shell=True)
xlsxFile.close()
call("pdftoppm {} -png > {}".format(pdfFile, pngFile.name), shell=True)

with open("/opt/bots/DutySheetScreenshot/URLs/dutySheetDiscordURL.txt", "r") as discordURLFile:
    discordURL = discordURLFile.readline().strip()


def lastSubstringAfter(s: str):
    i = s.rfind("/")
    return s[i + 1:] if i != -1 else s


webhook = DiscordWebhook(url=discordURL, content="**This week's Duty Sheet:** {}f/{}".format(ncURL, fileID))
webhook.add_file(pngFile.read(), filename=lastSubstringAfter(newestDutySheet)[:-4] + ".png")
webhook.execute()
