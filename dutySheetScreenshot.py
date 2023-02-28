#!/usr/bin/env python3

from urllib.parse import quote as urlencode
from xml.etree import ElementTree
import requests
from tempfile import NamedTemporaryFile
from subprocess import call
from sys import exit
from discord_webhook import DiscordWebhook
from os import remove

# Update this every semester!
dutiesPath = "/Shared/NewDrive/ALPHA SIG GENERAL/02_COMMITTEES/07_HOUSING/WEEKLY DUTIES/2023_SPRING/"


def readFile(path):
    with open(path, "r") as f:
        return f.readline().strip()


def lastSubstringAfter(s: str, delimiter: str):
    i = s.rfind(delimiter)
    return s[i + 1:] if i != -1 else s


# Grab the password and nextcloud URL from files.
password = readFile("/opt/bots/password")
ncURL = readFile("/opt/bots/DutySheetScreenshot/URLs/nextcloudURL.txt")
discordURL = readFile("/opt/bots/DutySheetScreenshot/URLs/dutySheetDiscordURL.txt")

# The NextCloud API requires the filepaths to be URL encoded.
dutiesPath = urlencode(dutiesPath)

# See https://docs.nextcloud.com/server/19/developer_manual/client_apis/WebDAV/basic.html for info on the NC API.
# Request a list of all the files in the duty sheet folder with all of their file IDs.
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
responseXML = ElementTree.fromstring(response)
# Pull out the file IDs from the response and associate them with their matching file name.
fileIDs = {}
files = responseXML.findall("{DAV:}response")
for file in files:
    fileID = None
    filePath = file.find("{DAV:}href").text
    for child in file.find("{DAV:}propstat").iter():
        if child.tag.endswith("fileid"):
            fileID = child.text
            break
    if fileID is not None:
        fileIDs[fileID] = filePath[25:]

fileID = max(fileIDs.keys())
newestDutySheet = fileIDs[fileID]

# Check if the file is the same as last week, if so, exit, otherwise, update the file.
with open("/opt/bots/DutySheetScreenshot/lastFileID", "w+") as lastIDFile:
    lastID = lastIDFile.read().strip()
    if fileID == lastID:
        print("Same as last week. Closing.")
        exit(0)

    lastIDFile.write(fileID)

docxFile = NamedTemporaryFile()
pdfFile = docxFile.name + ".pdf"
pngFile = NamedTemporaryFile()
# Grab the actual file and write it to a temp file
content = requests.get(
    url=ncURL + "remote.php/dav/files/bot" + newestDutySheet,
    auth=("bot", password)
).content
docxFile.write(content)
# Convert the docx to a PDF
call("libreoffice --headless --convert-to pdf {} --outdir /tmp".format(docxFile.name), shell=True)
docxFile.close()
# Convert the PDF to a PNG
call("pdftoppm {} -png > {}".format(pdfFile, pngFile.name), shell=True)

# Send a message in #house-announcements with the duty sheet attached as an image, and a link to the docx.
webhook = DiscordWebhook(url=discordURL, content="**This week's Duty Sheet:** {}f/{}".format(ncURL, fileID))
webhook.add_file(pngFile.read(), filename=lastSubstringAfter(newestDutySheet, "/")[:-4] + ".png")
webhook.execute()
remove(pdfFile)  # The other temp files will be removed automatically, but not this one.
