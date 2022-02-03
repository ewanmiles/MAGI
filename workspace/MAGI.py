import json, re, time
from datetime import date
import os, sys
import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import Combobox
from threading import Event, Thread

EVENT_TIMEOUT = 0.01
POLLING_DELAY = 1000
DESKTOP_PATH = desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

#Package install if necessary
print("Checking packages...")
with open('package.json', 'r') as f:
    packages = json.load(f)
    for (k,v) in packages.items():
        if v == False:
            print(f'{k} may not be installed. Downloading it to your system. Installer: pip')
            os.system(f'pip install {k}')
            packages[k] = True

with open('package.json', 'w') as f:
    json.dump(packages, f, indent=2)

import requests
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

#Unpack configs for monday, sharepoint and modules
#Load request details for Monday API (key, url, headers)
print('Fetching Monday API details...')
with open('mondayConfig.json', 'r') as f:
    mondayConfig = json.load(f)

#Load request details for SharePoint API (login, url)
print('Fetching SharePoint API details...')
with open('sharepointConfig.json', 'r') as f:
    spConfig = json.load(f)
    spConfig = spConfig['share_point']

#Fetch module names
print('\nFetching module structures...')
with open('moduleConfig.json', 'r') as f:
    module_dict = json.load(f)
    module_dict = module_dict['modules']


###################################################################


def windowLog(windowLabel: tk.Label, text):
    """
    Really basic. Just a print statement to the console that will also update the tkinter window text.
    text is Stringtype only.
    """
    if windowLabel.winfo_height() > 340: #If the frame height has been exceeded by text, clear all previous text
        windowLabel['text'] = ''

    #Update console and window
    print(f'\n{text}')
    windowLabel['text'] += f'\n{text}'


###################################################################


def writeApiLogins():
    global loginDict

    loginDict = {
        'spName': spNameBox.get(),
        'spPass': spPassBox.get(),
        'mKey': mKeyBox.get()
    }

    mondayConfig['monday']['login']['apiKey'] = loginDict['mKey']
    spConfig['user'] = loginDict['spName']
    spConfig['password'] = loginDict['spPass']

    newSpConfig = {
        "share_point": {
            "user": spConfig['user'],
            "password": spConfig['password'],
            "site": spConfig['site']
        }
    }

    #Save all the details to the configs
    with open('mondayConfig.json', 'w') as f:
        json.dump(mondayConfig, f, indent=2)
    with open('sharepointConfig.json', 'w') as f:
        json.dump(newSpConfig, f, indent=2)

    startup.destroy()

startup = tk.Tk()
startup.geometry("400x500")
acronym = tk.Label(text="MAGI", pady=5, font=('MS Serif', 45))
title = tk.Label(text="Madcap Automatic Graphic Implementation", pady=5, font=('MS Serif', 12))
button = tk.Button(
    text="Save details",
    width=25,
    height=2,
    bg="azure",
    command=writeApiLogins,
)

spName = tk.Label(text="SharePoint username", pady=10)
spPass = tk.Label(text="SharePoint password", pady=10)
mKey = tk.Label(text="Monday API Key", pady=10)

spNameBox = tk.Entry(bg="white", width=50)
spPassBox = tk.Entry(bg="white", show="*", width=50)
mKeyBox = tk.Entry(bg="white", width=50)

spNameBox.insert(0,spConfig['user'])
spPassBox.insert(0,spConfig['password'])
mKeyBox.insert(0,mondayConfig['monday']['login']['apiKey'])

for item in [acronym,title,spName,spNameBox,spPass,spPassBox,mKey,mKeyBox]:
    item.pack()
button.pack(pady=40)

startup.mainloop()

(mondayKey, mondayUrl) = (mondayConfig['monday']['login']['apiKey'], mondayConfig['monday']['login']['apiUrl'])
mondayHeaders = {"Authorization" : mondayKey}


def fetchEntries():
    global optDict
    optDict = {
        'module': moduleChoice.get(),
        'submodule': submoduleChoice.get(),
        'board': boardChoice.get()
    }

    window.destroy()

def browseFiles(label: tk.Label):
    """
    Opens a file dialog through tkinter to browse for the module folder; input
        - label: Label to change on the dialog box to show the chosen path.
    """
    global MODULE_PATH

    MODULE_PATH = filedialog.askdirectory(initialdir = "/",
                                          title = "Select your Module folder (e.g. containing Content, etc.)")
    try:
        MODULE_PATH = MODULE_PATH.replace('C:','')
    except:
        pass
    
    updateSubmoduleChoices()
    label['text'] = MODULE_PATH

def updateSubmoduleChoices():

    #Dict for dir tree to dump into json
    tree = {
        'modules': {}
    }

    #Centre brackets check for group X.X.X where X is any single integer
    #.* either side check for full string of literally any other characters (Watch out for this)
    #There will never be a perfect regex :(
    subSubRegex = "^(.*?([0-9]\.[0-9]*\.[a-zA-Z0-9]).*)$"
    #Now checking for X.X in centre
    subRegex = "^(.*?([0-9]\.[0-9]*).*)$"

    for root, dirs, files in os.walk(MODULE_PATH):
        for name in dirs:

            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                continue

            subSubCheck = re.findall(subSubRegex, name) #Same again for subsubmodule (X.X.X)
            if len(subSubCheck) > 0:
                tree['modules'][subSubCheck[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'dir'
                }
            else:
                subCheck = re.findall(subRegex, name) #Use regex to find submodule number as key from dir name
                if len(subCheck) > 0:
                    tree['modules'][subCheck[0][1]] = {
                        'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                        'type': 'dir'
                    }

        for name in files:
            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                continue

            subSubCheck = re.findall(subSubRegex, name) #Same again for subsubmodule (X.X.X)
            if len(subSubCheck) > 0:
                tree['modules'][subSubCheck[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'file'
                }
            else:
                subCheck = re.findall(subRegex, name) #Use regex to find submodule number as key from dir name
                if len(subCheck) > 0:
                    tree['modules'][subCheck[0][1]] = {
                        'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                        'type': 'file'
                    }

    submoduleList = list(tree['modules'].keys())
    submoduleList = [i for i in submoduleList if i[-1] != '.']
    submoduleChoice.set(submoduleList[0])
    submoduleOptions['values'] = ['All']+submoduleList

if 'loginDict' not in list(globals().keys()):
    sys.exit()

window = tk.Tk()
window.geometry("400x500")
acronym = tk.Label(text="MAGI", pady=5, font=('MS Serif', 45))
title = tk.Label(text="Madcap Automatic Graphic Implementation", pady=5, font=('MS Serif', 12))
button = tk.Button(
    text="Add graphics",
    width=25,
    height=2,
    bg="azure",
    command=fetchEntries,
)

mTitle = tk.Label(text="Module", pady=10)
submTitle = tk.Label(text="Submodule", pady=10)
pathTitle = tk.Label(text="Select a file path", pady=10)
pathSubtitle = tk.Label(text="Not chosen", pady=10, fg='blue')
bTitle = tk.Label(text="Monday Board", pady=10)

pathButton = tk.Button(window, text = "Browse Files", command=lambda: browseFiles(pathSubtitle))

moduleChoice = tk.StringVar(window)
moduleChoice.set('M1')
modules = [f'M{i}' for i in range(18)]
moduleOptions = Combobox(window, textvariable=moduleChoice, values=modules)

submoduleChoice = tk.StringVar(window)
submoduleChoice.set('1.1')
submodules = ['All',0,1,2]
submoduleOptions = Combobox(window, textvariable=submoduleChoice, values=submodules)

boards = modules + ['Upload to Flare', 'Media Requests', 'Learning Designers', 'Sort into Modules', 'Generic Content Archive']
boardChoice = tk.StringVar(window)
boardChoice.set('Upload to Flare')
boardOptions = Combobox(window, textvariable=boardChoice, values=boards)

#moduleOptions.bind("<<ComboboxSelected>>", updateSubmoduleChoices)

for item in [acronym,title,mTitle,moduleOptions,pathTitle,pathSubtitle,pathButton,submTitle,submoduleOptions,bTitle,boardOptions]:
    item.pack()
button.pack(pady=40)

window.mainloop()


###################################################################

#MOVING TO FUNCTIONS FOR RUNTIME AND MAIN BODY OF CODE

###################################################################


def queryBoardID(boardName):
    """
    Sends a query to monday for a board ID that is not in the monday config; inputs
        - boardName (str): Name of the Monday board to find the ID for
    PERMANENTLY ADDS OR REPLACES the board ID in the monday config
    """

    query = """{ boards {
                    id
                    name
                    }
                }"""

    data = {'query' : query}

    #Post request, parse JSON response
    r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)
    returned = r.json()['data']['boards']
    Id = None

    #Iterate through returned boards to find matching board name and its ID
    for board in returned:
        if board['name'] == boardName: #Board name matches given name
            Id = board["id"]
            mondayConfig['monday']['boardIDs'][boardName] = Id #Return board ID

            #Save new ID to config for future
            with open('mondayConfig.json', 'w') as f:
                json.dump(mondayConfig, f, indent=2)

            windowLog(info, f'Board {boardName} (ID {board["id"]}) saved to the config.')
    
    return Id

def downloadFiles(fileList):
    """
    Logs in to SharePoint API and downloads list of files given, saving them to a specific directory; inputs
        - fileList (arr): List of file names to search for (no extension), here scraped from Monday
        - outDir (str/path): Directory to save the downloaded files to
    Note: REQUIRES SharePoint API config external with valid credentials, unpacked as global variable
    """
    #URL to RG SharePoint site from config
    site_url = spConfig['site']

    #Create API context for queries
    ctx = ClientContext(site_url).with_credentials(UserCredential(spConfig['user'], spConfig['password']))
    
    #Access web API
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    extensions = ['.png', '.jpg', '.jpeg', '.gif'] #Possible file extensions (will check against all for given file name)

    for fi in fileList:
        module = parseSubmodule(fi)[0] #Get module from file name

        for ext in extensions:
            try:
                #Get file from API
                response = web.get_file_by_server_relative_path(f"/sites/Part-66Project/Shared Documents/General/Modules/{module_dict[module]['name']}/Development/Graphics/New {module} Graphics/{fi}{ext}")
                download_path = os.path.join(buildContentPaths(fi, module)[1], f'{fi}{ext}')

                with open(download_path, "wb") as local_file:
                    response.download(local_file).execute_query() #Execute download query (THIS LINE WILL THROW THE HTTP ERRORS)

                    graphicReplaceMap[fi]['ext'] = ext #Add graphic's extension to dict map for reference

            except:
                #print(f'{fi}{ext} not found.') #Try/except should catch files that don't exist in SharePoint (wrong extension, wrong name, etc)
                try:
                    os.remove(download_path) #File is already written in local ready to receive download, delete as file not found
                except (FileNotFoundError, UnboundLocalError) as e:
                    pass
                continue

            #This is by far the best way to check if it has been downloaded. By this point, if unsuccessful, path should have been removed above
            if os.path.exists(download_path):
                reportContent[fi] = 1
                windowLog(info, f'{fi}{ext} downloaded to {download_path}.')


def buildModuleConfig(path, outfile, debug=False):
    """
    Uses regexes to build a module path config from given root path; inputs
    - path (str): Top down point to start building dir tree from
    - outfile (str): Name of/path to json dump file
    - debug (bool): OPTIONAL kwarg. Print statements to show how it builds json for support
    Dumps a JSON for each submodule path into outfile
    """
    #Dict for dir tree to dump into json
    tree = {
        'modules': {}
    }

    #Centre brackets check for group X.X.X where X is any single integer
    #.* either side check for full string of literally any other characters (Watch out for this)
    #There will never be a perfect regex :(
    subSubRegex = "^(.*?([0-9]\.[0-9]*\.[a-zA-Z0-9]).*)$"
    #Now checking for X.X in centre
    subRegex = "^(.*?([0-9]\.[0-9]*).*)$"

    for root, dirs, files in os.walk(path):
        for name in dirs:
            if debug:
                print(f'DIR: {name}')

            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                if debug:
                    print('SKIPPED', os.path.join(root,name))
                continue

            subSubCheck = re.findall(subSubRegex, name) #Same again for subsubmodule (X.X.X)
            if len(subSubCheck) > 0:
                if debug:
                    print(f'CHECKING SUBSUBMODULES (X.X.X): {subSubCheck}')
                tree['modules'][subSubCheck[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'dir'
                }
            else:
                subCheck = re.findall(subRegex, name) #Use regex to find submodule number as key from dir name
                if len(subCheck) > 0:
                    if debug:
                        print(f'CHECKING SUBMODULES (X.X): {subCheck}')
                    tree['modules'][subCheck[0][1]] = {
                        'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                        'type': 'dir'
                    }

        for name in files:
            if debug:
                print(f'FILE: {name}')

            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                if debug:
                    print('SKIPPED', os.path.join(root,name))
                continue

            subSubCheck = re.findall(subSubRegex, name) #Same again for subsubmodule (X.X.X)
            if len(subSubCheck) > 0:
                if debug:
                    print(f'CHECKING SUBSUBMODULES (X.X.X): {subSubCheck}')
                tree['modules'][subSubCheck[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'file'
                }
            else:
                subCheck = re.findall(subRegex, name) #Use regex to find submodule number as key from dir name
                if len(subCheck) > 0:
                    if debug:
                        print(f'CHECKING SUBMODULES (X.X): {subCheck}')
                    tree['modules'][subCheck[0][1]] = {
                        'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                        'type': 'file'
                    }

    with open(outfile, 'w') as fp:
        json.dump(tree, fp, indent=2) #Save external copy to json

    return tree['modules']

def fetchDataIndex(jsonArr, name):
    """
    Cycles through json list to return first object with given name field, otherwise returns none; inputs
        - jsonArr (arr): List of json objects
        - name (str): Value of 'name' field to check for
    DO NOT USE ON LARGE OBJECT LISTS. This is very basic code used for smaller operations.
    """

    index = 0

    while index < len(jsonArr):
        if jsonArr[index]['name'] == name:
            return index
        
    return None

def buildItemList(jsonArr, name):
    """
    Cycles through json object list, finds first object with matching name field, builds list of items from that board; inputs
        - jsonArr (arr): List of json objects
        - name (str): Value of 'name' field to check for
    DO NOT USE ON LARGE OBJECT LISTS. This is very basic code used for smaller operations.
    """

    nameList = [i['name'] for i in jsonArr[fetchDataIndex(jsonArr, name)]['items']]
    return nameList

def deleteMultiple(text, substrs):
    """
    Uses replace string method to delete multiple substrings within given string; inputs
        - text (str): Given string to delete substrings from
        - substrs (arr): List of substrings to remove
    Returns the string with substrings deleted, note: will remove ALL instances of each given substring
    """

    for sub in substrs:
        if type(sub) == str:
            text = text.replace(sub, "")
        else:
            print(f'Cannot remove {sub} from {text} as {sub} is not a stringtype')

    return text

def parseSubmodule(fileName):
    """
    Parses module and full submodule path from a given file name; inputs
        - fileName (str): File to parse module and submodule from
    Returns two-item list [module, submodule], module will be "M_"
    """

    #Centre brackets check for group X.X.X where X is any single integer
    #\D either side check for full string of any non numeric characters either side
    subSubRegex = "^(.*?([0-9]\.[0-9]*\.[a-zA-Z0-9]).*)$"
    #Now checking for X.X in centre
    subRegex = "^(.*?([0-9]\.[0-9]*).*)$"

    submodule = ''

    subCheck = re.findall(subRegex, fileName) #Use regex to find submodule number as key from dir name
    if len(subCheck) > 0:
        submodule = subCheck[0][1]

    subSubCheck = re.findall(subSubRegex, fileName) #Same again for subsubmodule (X.X.X)
    if len(subSubCheck) > 0:
        submodule = subSubCheck[0][1]

    module = 'M' + submodule[:submodule.find('.')]

    return [module, submodule]


def buildContentPaths(fileName, module):
    """
    """

    try:
        sub = module_dict[module]['submodules'][graphicReplaceMap[fileName]['submodule']]
        return [os.path.join(f'{MODULE_PATH}/Content', sub), os.path.join(f'{MODULE_PATH}/Content/Resources/Images/New/')]
    except KeyError:
        path = submodule_dict[parseSubmodule(fileName)[1]]['path']
        return [path, os.path.join(f'{MODULE_PATH}/Content/Resources/Images/New/')]

def replaceExistingFigures(files):
    for filePath in files:
        #Read lines of origin doc (no embedded images here)
        if os.path.isfile(filePath.replace('.htm', '-embedded.htm')):
            with open(filePath.replace('.htm', '-embedded.htm'), "r", encoding='utf8') as f:
                line_list = f.readlines()
        else:
            with open(filePath, 'r', encoding="utf8") as f:
                line_list = f.readlines()

        indexList = []
        for ind, l in enumerate(line_list):
            if '<figure' in l or '</figure>' in l:
                indexList.append(ind)

        for ind in sorted(indexList, reverse=True):
            del line_list[ind]

        print(line_list)

        newDoc = ""
        for l in line_list:
            newDoc += l

        #Create duplicate file
        try:
            with open(f"{filePath.replace('.htm', '-embedded.htm')}", "w", encoding="utf8") as f:
                f.write(newDoc)
                f.close()
        except PermissionError:
            info['text'] += f"\nERROR! MAGI can't access {filePath.replace('.htm', '-embedded.htm')}. This is likely because you have it open in Flare, or somewhere else. Please close the editor and try again."

def addFigureTags(l):
    """
    Adds figure tags back into documents where they are absent around <img><figcaption>; input
        -l (arr): Unpacked lines(mainly from readlines()) from the target document
    Returns the same list of lines back with figure tags added in
    """

    #While loop is one of the easiest ways
    ind = 0
    while ind < len(l)-1:

        if ('<img' in l[ind]) and ('<figure' not in l[ind-1]): #Line with img tag but no figure tag before, add figure previous
            l = l[:ind] + ['\t<figure class="fortyPercent">\n'] + l[ind:]
            ind += 2

        elif ('<figcaption' in l[ind]) and ('</figure' not in l[ind+1]): #Line with figcaption tag but no closing figure tag after, add /figure next
            l = l[:ind+1] + ['\t</figure>\n'] + l[ind+1:]
            ind += 2

        #Regular line, nothing interesting  
        else:
            ind += 1

    return l

def embedImages(graphic, filePath):
    #Read lines of origin doc (no embedded images here)
    if os.path.isfile(filePath.replace('.htm', '-embedded.htm')):
        with open(filePath.replace('.htm', '-embedded.htm'), "r", encoding='utf8') as f:
            line_list = f.readlines()
    else:
        with open(filePath, 'r', encoding="utf8") as f:
            line_list = f.readlines()

    for index, l in enumerate(line_list):
        if graphic in l:
            graphicInd = index

    #New empty HTML to write the file
    newDoc = ""
    for ind, l in enumerate(line_list): #Go through lines of origin doc with image names and captions
        try:
            if (ind == graphicInd):
                continue #Skip lines with graphic names

            elif ind-1 == graphicInd: #This line has the graphic caption
                try:
                    graphicExt = graphicReplaceMap[graphic]['ext'] #Get graphic extension from dict
                    gPathPrefix = graphicReplaceMap[graphic]['pathPrefix']
                except KeyError:
                    windowLog(info, 'Graphic extension not successfully found. It is likely the file was not downloaded. Skipping embed.')
                    return None

                #Build caption from line
                caption = deleteMultiple(line_list[ind], ['<p>','</p>','  ','\n','\t','<figcaption>','</figcaption>'])
                
                #Build snippet from external HTML file
                with open('baseSnippet.html', 'r') as f:
                    snip = f.read()
                    snip = snip.replace('src=""', f'src="{gPathPrefix}Resources/Images/New/{graphic}{graphicExt}"')
                    snip = snip.replace("><", f">{caption}<")
                
                windowLog(info, f"{graphic}{graphicExt} successfully embedded in {filePath.replace('.htm','-embedded.htm')}\n")
                newDoc += f'\n{snip}\n\n' #Add snippet to new document (KEEP THE NEWLINES HERE)
                reportContent[graphic] = 2

            else:
                newDoc += l #No graphics in line, just add the next line as normal unless a figure tag

        except UnboundLocalError:
            newDoc += l

    #Create duplicate file
    try:
        with open(f"{filePath.replace('.htm', '-embedded.htm')}", "w", encoding="utf8") as f:
            f.write(newDoc)
            f.close()
    except PermissionError:
        info['text'] += f"\nERROR! MAGI can't access {filePath.replace('.htm', '-embedded.htm')}. This is likely because you have it open in Flare, or somewhere else. Please close the editor and try again."

def clearEmbeddedFiles(debug=False):
    """
    Checks over the directory tree built in the submodule config and removes any existing files with -embedded in the name
    Optional argument, debug (bool): Prints the name of each removed file
    """

    windowLog(info, "Removing previously embedded files...")
    fileCount = 0

    for submodule in list(submodule_dict.values()):
        if submodule['type'] == 'file':
            try:
                #Try and remove file with -embedded
                os.remove(submodule['path'].replace('.htm', '-embedded.htm'))
                fileCount += 1
                if debug:
                    print(f"Removed {submodule['path'].replace('.htm', '-embedded.htm')}")

            except FileNotFoundError: #No embedded file, move on
                continue

        elif submodule['type'] == 'dir':
            for file in os.listdir(submodule['path']): #Go through files in submodule dir

                if '-embedded' in file: #If embedded file, remove
                    os.remove(os.path.join(submodule['path'],file))
                    fileCount += 1
                    if debug:
                        print(f"Removed {os.path.join(submodule['path'],file)}")

    windowLog(info, f"\nRemoved {fileCount} previous files.")


###################################################################


def main(event):

    global graphicReplaceMap, submodule_dict, reportContent

    #Settings
    targetModule = optDict['module']
    targetSubmodule = None if optDict['submodule'] == 'All' else optDict['submodule']
    targetBoard = optDict['board']

    #Build module config and unpack
    windowLog(info, f'Project found at {MODULE_PATH}.')
    windowLog(info, 'Building submodule config...')
    submodule_dict = buildModuleConfig(f'/{targetModule}/Content', 'submoduleConfig.json')

    #Clear any previously embedded files
    clearEmbeddedFiles()

    #Build file list to remove figure tags
    windowLog(info, 'Removing figure tags...')
    if targetSubmodule:
        dir_list = [submodule_dict[k]['path'] for k in list(submodule_dict.keys()) if targetSubmodule in submodule_dict[k]['path']]
    else:
        dir_list = [submodule_dict[k]['path'] for k in list(submodule_dict.keys())]
    files = []
    for directory in dir_list:
        for page in os.listdir(directory):
            files.append(os.path.join(directory, page))

    replaceExistingFigures(files)

    #Build image dir if it doesn't exist
    if os.path.exists(f'{MODULE_PATH}/Content/Resources/Images/New/') == False:
        os.mkdir(f'{MODULE_PATH}/Content/Resources/Images/New/')

    try:
        boardID = mondayConfig['monday']['boardIDs'][targetBoard]
    except KeyError:
        windowLog(info, f"\nBoard {targetBoard} not found in config. Sending query to Monday for board ID...")
        boardID = queryBoardID(targetBoard)
        if boardID == None:
            raise ValueError(f'Unable to find the board {targetBoard} on Monday. Please try another name, or check Monday.')

    #Query to send to Monday API
    #Get target board ID from monday config
    queryFirstLine = "{{ boards (ids:[{0}]) {{".format(boardID)
    query = queryFirstLine + """name
            items {

                name
                } 
            } 
        }"""
    data = {'query' : query}

    tempStr = '\n##############################################################'
    tempStr += f'\nSending Monday query to board {targetBoard}, ID {boardID}...'
    tempStr += '\n##############################################################\n'
    windowLog(info, tempStr)

    #Post request, parse JSON response
    r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)
    returned = r.json()['data']['boards']

    #Build a list of item names from the Monday board
    targetGraphics = buildItemList(returned, targetBoard)
    windowLog(info, f'{spConfig["user"]} logging in to SharePoint...')

    #Slice for targeted submodule if set
    if targetSubmodule:
        targetGraphics = [i for i in targetGraphics if targetSubmodule in i]

    #This is essentially the god map, saving relevant info about each graphic, with graphic name being the key
    graphicReplaceMap = {}
    for graphic in targetGraphics:
        parsed = parseSubmodule(graphic)
        graphicReplaceMap[graphic] = {
            'module': parsed[0],
            'submodule': parsed[1]
        }

    # Save file names to dict that will write a report at end
    reportContent = {}
    for i in targetGraphics:
        reportContent[i] = 0

    #Time file download and embed
    start = time.time()

    info['text'] += '\n' #Aesthetic for logger
    downloadFiles(targetGraphics)

    #Cycle through graphics, if submodule not correctly parsed, change report content to parser error
    unparsedGraphics = []
    for k,v in graphicReplaceMap.items():
        if not v['submodule']:
            reportContent[k] = -1
            unparsedGraphics.append(k)

    #Remove empty graphic map entries (random Monday scrapes that aren't graphics or don't have module information)
    graphicReplaceMap = {k: v for k,v in graphicReplaceMap.items() if v['submodule']}
    targetGraphics = list(graphicReplaceMap.keys())
    fileCount = len(targetGraphics) #For report


    for graphic in targetGraphics:
        parsed = parseSubmodule(graphic) #Get module/submodule

        #Find if file or dir from submodule config, add path prefix (../ or ../../ typically) accordingly
        graphicReplaceMap[graphic]['pathPrefix'] = parsed[1].count('.')*'../' if submodule_dict[parsed[1]]['type'] == 'dir' else (parsed[1].count('.')-1)*'../'

    
    if os.path.exists(f'{DESKTOP_PATH}/FlareResults') == False:
        os.mkdir(f'{DESKTOP_PATH}/FlareResults')

    with open(f'{DESKTOP_PATH}/FlareResults/grm.json', 'w') as fp:
            json.dump(graphicReplaceMap, fp, indent=2) #Save external copy of graphicReplaceMap to grm.json
    windowLog(info, 'The Graphic Replace Map has been written to grm.json for checking.')

    for ind, fi in enumerate(targetGraphics):
        tempPath = submodule_dict[graphicReplaceMap[fi]['submodule']]['path'] #Get path to submodule for graphic

        if os.path.isdir(tempPath):
            for page in os.listdir(tempPath): #If submodule is folder, go through each page
                pagePath = os.path.join(tempPath, page)
                if '-embedded' in pagePath: #Ignore files that have already embedded images (this is checked again in embedImages())
                    continue

                #Read the currently observed page
                with open(pagePath, 'r', encoding='utf8') as f:
                    content = f.read()
                    if fi in content:
                        windowLog(info, f'Embedding {fi} to {pagePath}...') #Find graphic name in page, embed
                        embedImages(fi, pagePath)
                    else:
                        f.close() #Graphic not found in page
        
        #Case that submodule is file (.htm page) not folder
        else:
            pagePath = tempPath
            if '-embedded' in pagePath: #Again ignore embedded pages, caught in embedImages()
                continue
            
            with open(pagePath, 'r', encoding='utf8') as f:
                content = f.read()
                if fi in content:
                    windowLog(info, f'Embedding {fi} to {pagePath}...') #Find graphic name in page, embed
                    embedImages(fi, pagePath)
                else:
                    f.close() #Graphic not found in page

    end = time.time()

    #Build file list to remove figure tags
    windowLog(info, 'Replacing broken figure tags...')
    if targetSubmodule:
        dir_list = [submodule_dict[k]['path'] for k in list(submodule_dict.keys()) if targetSubmodule in submodule_dict[k]['path']]
    else:
        dir_list = [submodule_dict[k]['path'] for k in list(submodule_dict.keys())]
    files = []
    for directory in dir_list:
        for page in os.listdir(directory):
            if '-embedded.htm' in page:
                files.append(os.path.join(directory, page))

    for fi in files:
        with open(fi, 'r') as f:
            lineList = f.readlines()
            
        lineList = addFigureTags(lineList)
        with open(fi, 'w') as f:
            for line in lineList:
                f.write(line)

    #Statuses are saved as 0/1/2, this is the key value map for the report
    reportRef = {
        -1: 'Parser Error',
        0: 'Not Found',
        1: 'Downloaded',
        2: 'Embedded'
    }

    windowLog(info, 'Writing graphic statuses to a CSV...')
    time.sleep(3)

    #Writing graphics and their status to csv
    col1 = targetGraphics + unparsedGraphics
    col2 = [reportRef[reportContent[i]] for i in col1]
    excelData = pd.DataFrame([col1, col2]).transpose()
    excelData.to_csv(f'{DESKTOP_PATH}/FlareResults/embeddingReport.csv', index=None, header=None)
    windowLog(info, f'The report data is available at {DESKTOP_PATH}/FlareResults/embeddingReport.csv.')

    #Plot bar chart of embedded/downloaded/not found
    plt.figure()
    plt.bar([0.5, 1.5, 2.5, 3.5], [list(reportContent.values()).count(-1), 
                                    list(reportContent.values()).count(0), 
                                    list(reportContent.values()).count(1), 
                                    list(reportContent.values()).count(2)])
    plt.title('Flare Import Results')
    plt.xlabel('Graphic Status')
    plt.ylabel('Frequency')
    plt.xticks(ticks=[0.5,1.5,2.5,3.5],labels=['Parser Error','Not Found','Downloaded','Embedded'])
    plt.savefig(f'{DESKTOP_PATH}/FlareResults/importResults.pdf')
    windowLog(info, f'A bar graph for the graphics report is available at {DESKTOP_PATH}/FlareResults/importResults.pdf.')

    windowLog(info, 'Writing report...')
    time.sleep(3)

    #Write report to txt file
    with open(f'{DESKTOP_PATH}/FlareResults/embeddingReport.txt', 'w') as f:

        #Count of file statuses
        errorCount = ('Parser Error', list(reportContent.values()).count(-1))
        unfoundCount = ('Not Found', list(reportContent.values()).count(0))
        downCount = ('Downloaded', list(reportContent.values()).count(1))
        embedCount = ('Embedded', list(reportContent.values()).count(2))

        f.write('Images have been embedded. Here is a status list.\n\n')
        f.write(f'{date.today()}\n')
        f.write(f'Time from start of download to end of embed: {(end - start):.2f} s\n\n')
        
        f.write(f'{errorCount[0]}: {errorCount[1]} (Not included in the % below)\n')
        for count in [unfoundCount, downCount, embedCount]:
            f.write(f'{count[0]}: {count[1]} ({(100*count[1])/fileCount:.1f}%)\n')
            
        f.write('\n')

        for name, status in zip(col1,col2):
            f.write(f'{name}: {status}\n')

        f.close() #Just to make sure

    windowLog(info, f'A full report of the graphic embed is available at {DESKTOP_PATH}/FlareResults/embeddingReport.txt.')
    time.sleep(5)

    event.set() #Signal to threader that process is complete

##########################################################################


def checkStatus(logger, event):
    event_is_set = event.wait(EVENT_TIMEOUT)
    if event_is_set:
        logger.destroy()
    else:
        logger.after(POLLING_DELAY, checkStatus, logger, event)

if 'optDict' not in list(globals().keys()):
    sys.exit()

logger = tk.Tk()
logger.geometry("600x600")
acronym = tk.Label(text="MAGI", pady=5, font=('MS Serif', 45))
title = tk.Label(text="Madcap Automatic Graphic Implementation", pady=5, font=('MS Serif', 12))
button = tk.Button(
    text="Cancel",
    width=25,
    height=2,
    bg="azure",
    command=sys.exit,
)
log = tk.Frame(master=logger, width=550, height=360, bg="white")
info = tk.Label(master=log, text="Starting the process...", bg="white", justify='left', wraplength=540)
for item in [acronym,title,log]:
    item.pack()
info.place(x=0, y=0)

button.pack(pady=40)

event = Event()
thread = Thread(target=main, args=(event,))
checkStatus(logger, event)  # Starts the polling of main() status
thread.start()

logger.mainloop()

final = tk.Tk()
final.geometry("540x210")
title = tk.Label(text="Graphics complete!", pady=5, font=('MS Serif', 12))
subtitle = tk.Label(text="Make sure to check your project in Flare to see which files have been successfully embedded.", pady=5)
moreInfo = tk.Label(text="Report: 'Desktop/FlareResults/embeddingReport.txt'\nSpreadsheet: 'Desktop/FlareResults/embeddingReport.csv'\nGraph: 'Desktop/FlareResults/importResults.pdf'")
button = tk.Button(
    text="Ok",
    width=20,
    height=2,
    bg="azure",
    command=final.destroy
)

title.pack()
subtitle.pack()
moreInfo.pack()
button.pack(pady=30)

final.mainloop()

sys.exit()