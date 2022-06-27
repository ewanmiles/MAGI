import json, re, os

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

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

import requests
import eel

eel.init("web") #Initialise front end files from web dir
DESKTOP_PATH = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

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

def parseSubmodule(text):
    """
    Parses module and full submodule path from a given file name; input
        
        - text (str): File to parse module and submodule from
    
    Returns two-item list [module, submodule], module will be "M_"
    """

    #This regex can essentially pick up two, three and four digit submodules, e.g. 1.1 - 1.1.1.1
    reg = "^(.*?((([0-9]{1,}\.[0-9]{1,}[a-zA-Z]?)(\.[0-9]{1,}[a-zA-Z]?)?)(\.[0-9]{1,}[a-zA-Z]?)?).*)$"

    submodule = ''

    check = re.findall(reg, text) #Find regex occurrences in text (e.g. 6.11)
    if len(check) > 0:
        submodule = check[0][1]

    module = 'M' + submodule[:submodule.find('.')]

    return [module, submodule]

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

def generateFigure(img, caption, conds):
    """
    Simple function to avoid hardcoding figure snippet. Takes image name, caption and conditions and turns out necessary figure to add to new doc; inputs
        
        - img (str): Image name to embed (typically relative path from location to New/ folder)
        - caption (str): Caption to embed (just caption, tags etc. removed)
        - conds (str): MC Conditions for the figure (MAKE SURE THE FIRST CHAR IS A SPACE, e.g. ' {conds}')
    
    Returns a figure in four lines as an array to add to the line list of the new document
    """

    figure = []

    figure += [f'\n<figure{conds} class="fortyPercent">']
    figure += [f'\n\t<img src="{img}"/>']
    figure += [f'\n\t<figcaption>{caption}</figcaption>']
    figure += ['\n</figure>\n\n']

    return figure

@eel.expose
def py_setModulePath(newPath):
    """
    Basic setter for global variable MODULE_PATH that can be accessed from the JS
    """

    global MODULE_PATH

    MODULE_PATH = newPath
    print(f'Module path set to {MODULE_PATH}')

@eel.expose
def fillLoginDetails():
    """
    Sends the login details unpacked from the Monday/SharePoint jsons to the front end for unpacking into the UI.
    """

    return {
        'spUser': spConfig['user'],
        'spPass': spConfig['password'],
        'mKey': mondayConfig['monday']['login']['apiKey']
    }

@eel.expose
def py_writeLoginDetails(dict):
    """
    Saves the login details entered in the UI to the Monday/SharePoint jsons for loading in future.
    """

    print('Writing new login details...')

    mondayConfig['monday']['login']['apiKey'] = dict['mKey']
    spConfig['user'] = dict['spName']
    spConfig['password'] = dict['spPass']

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

@eel.expose
def createLoginContext(loginObj):
    """
    Creates a SharePoint login context as a global variable for the Python to make API calls using the given credentials, with input

        - loginObj (dict): Contains login credentials to create context with, input in the front end

    The main intention is to create the global context variable. Can either return a success message or an error trigger for the front end errorHandling.js.
    """
    global spCtx

    #URL to RG SharePoint site from config
    site_url = spConfig['site']

    #Create API context for queries
    spCtx = ClientContext(site_url).with_credentials(UserCredential(loginObj['spName'], loginObj['spPass']))
    
    #Access web API
    web = spCtx.web
    spCtx.load(web)
    try:
        spCtx.execute_query()
        return 'Successfully logged in to SharePoint.'

    except IndexError:
        return ['error', 'password']
    except ValueError:
        return ['error', 'username']

@eel.expose
def py_getActions(actionSet):
    """
    Small function to get action set from js_getActions() in the front end and choose functions to run from the details.

    Runs other functions, think of this as a mapping from the front end details to run the correct programs in Python.
    """

    #Map of functions to run according to received actions
    actionMap2 = {
        'embedGraphics': [downloadFiles, embedGraphics] if actionSet['imageLocation'] == 'sharepoint' else [embedGraphics],
        'updateMonday': [updateMonday],
        'outputReport': [outputReport],
    }

    #Escape if no actions selected, log in JS
    if len(actionSet['actions']) < 1:
        eel.log('You haven\'t selected any actions.')
        return None

    #Build function list to iterate from included actions
    fList = []
    for action in actionSet['actions']:
        fList += actionMap2[action]

    for f in fList:
        f(actionSet) #Execute each function

@eel.expose
def compileSubmoduleChoices(modulePath):
    """
    Checks submodule Regex to build list of submodule folders in the target module, with input

        - modulePath (str, pathlike): Path to target module to build submodule list for

    Returns the list of submodule folders found in the target location
    """

    #This regex can essentially pick up two, three and four digit submodules, e.g. 1.1 - 1.1.1.1
    reg = "^(.*?((([0-9]{1,}\.[0-9]{1,}[a-zA-Z]?)(\.[0-9]{1,}[a-zA-Z]?)?)(\.[0-9]{1,}[a-zA-Z]?)?).*)$"

    tree = {
        'modules': {}
    }

    for root, dirs, files in os.walk(modulePath):
        for name in dirs:

            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                continue

            check = re.findall(reg, name) #Check for subsubmodule (X.X - X.X.X.X)
            if len(check) > 0:
                tree['modules'][check[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'dir'
                }

        for name in files:
            if 'Resources' in root: #Skip files under the resources folder (just need document structure)
                continue

            check = re.findall(reg, name) #Same again for subsubmodule (X.X - X.X.X.X)
            if len(check) > 0:
                tree['modules'][check[0][1]] = {
                    'path': os.path.join(root,name), #Key is submodule number, val is path to submodule from root
                    'type': 'file'
                }

    submoduleList = list(tree['modules'].keys())
    submoduleList = [i for i in submoduleList if i[-1] != '.']
    submoduleList.sort()

    return submoduleList

@eel.expose
def queryBoardID(boardName):
    """
    Sends a query to monday for a board ID that is not in the monday config; inputs
        
        - boardName (str): Name of the Monday board to find the ID for
    
    PERMANENTLY ADDS OR REPLACES the board ID in the monday config
    """

    (mondayKey, mondayUrl) = (mondayConfig['monday']['login']['apiKey'], mondayConfig['monday']['login']['apiUrl'])
    mondayHeaders = {"Authorization" : mondayKey}

    query = """{ boards {
                    id
                    name
                    }
                }"""

    data = {'query' : query}

    try:
        r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)
        returned = r.json()['data']['boards']
        Id = None
    except:
        eel.log('Sorry, you do not appear to be connected to the internet. Please check your connection.')
        return None

    #Iterate through returned boards to find matching board name and its ID
    for board in returned:
        if board['name'] == boardName: #Board name matches given name
            Id = board["id"]
            mondayConfig['monday']['boardIDs'][boardName] = Id #Return board ID

            #Save new ID to config for future
            with open('mondayConfig.json', 'w') as f:
                json.dump(mondayConfig, f, indent=2)

            return f'Board {boardName} (ID {board["id"]}) saved to the config.'
    
    return [boardName, 'Not Found']

def checkProjectImages(images):
    #Full scan of all docs in project
    htmlCollection = []
    for root, dirs, files in os.walk(f'{MODULE_PATH}/Content'):
        for f in files:
            if f.endswith('.htm'):
                path = os.path.join(root, f)
                if 'Welcome' not in path:
                    htmlCollection.append(path)

    #Builds dict of doc: images; e.g. {HTM FILE: [IMG 1, IMG 2, ...], ...}
    initialDict = {}
    for doc in htmlCollection:
        with open(doc, 'r', encoding='utf-8') as d: #Read each htm file as one string
            asString = d.read()

        found = [i for i in images if os.path.splitext(i)[0] in asString] #Images found in this htm
        temp = {}

        if len(found) > 0:
            for img in found:
                temp[img] = asString[:asString.find(os.path.splitext(img)[0])].count('\n')
                
            initialDict[doc] = temp

    #Builds dict of img: {doc, line, embedded?}; e.g. {IMG: {'doc': HTM FILE, 'line': LINE, 'emb': True}, ...}
    resultDict = {}
    for doc in list(initialDict.keys()): #Get htm file names from previous dict
        graphics = list(initialDict[doc].keys())
        inds = list(initialDict[doc].values())

        with open(doc, 'r', encoding='utf-8') as f: #Now unpack files as list of lines
            lineList = f.readlines()

        for g, i in zip(graphics, inds):
            resultDict[g] = {
                'doc': doc,
                'ind': i,
                'emb': False
            }

            if '<img' in lineList[i]: #Signifies already properly embedded in img tag
                resultDict[g]['emb'] = True

    return resultDict

@eel.expose
def downloadFiles(actionSet):
    """
    Logs in to SharePoint API and downloads list of files given, saving them to a specific directory; input

        - actionSet (dict): Set of information about what actions to take from script.js, including image locations and login details. See js_getActions().

    Note: REQUIRES SharePoint API config external with valid credentials, unpacked as global variable. No return.
    """
    
    (targetSubmodule, targetBoard) = (actionSet['actionDetails']['submoduleSelect'], actionSet['actionDetails']['boardSelect'])

    #Monday request header and context
    (mondayKey, mondayUrl) = (mondayConfig['monday']['login']['apiKey'], mondayConfig['monday']['login']['apiUrl'])
    mondayHeaders = {"Authorization" : mondayKey}

    try:
        boardID = mondayConfig['monday']['boardIDs'][targetBoard]
    except KeyError:
        eel.log(f"Board {targetBoard} not found in config. Sending query to Monday for board ID...")
        boardID = queryBoardID(targetBoard)
        if type(boardID) == list:
            eel.raiseError('Monday', boardID)
            return None

    boardID = mondayConfig['monday']['boardIDs'][targetBoard] #Try again; if code gets to this stage board ID should be saved to config

    #Actual Monday query
    queryFirstLine = "{{ boards (ids:[{0}]) {{".format(boardID)
    query = queryFirstLine + """name
            items {

                name
                } 
            } 
        }"""
    data = {'query' : query}

    #Post request, parse JSON response
    try:
        r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)
        returned = r.json()['data']['boards']
    except:
        eel.log('Sorry, you do not appear to be connected to the internet. Please check your connection.')
        return None

    #Build a list of item names from the Monday board
    targetGraphics = buildItemList(returned, targetBoard)

    #Slice for targeted submodule if set
    if targetSubmodule:
        targetGraphics = [i for i in targetGraphics if targetSubmodule in i]

    #Access web API
    web = spCtx.web
    spCtx.load(web).execute_query()

    extensions = ['.png', '.jpg', '.jpeg'] #Possible file extensions (will check against all for given file name)
    
    #Create New image folder if doesn't exists
    if os.path.exists(f'{MODULE_PATH}/Content/Resources/Images/New/') == False:
        os.mkdir(f'{MODULE_PATH}/Content/Resources/Images/New/')

    #Build UI for file downloads
    eel.buildDownloadList(targetGraphics)

    if len(targetGraphics) < 1:
        eel.log(f"Couldn't find any graphics for {targetSubmodule} on the {targetBoard} board.")

    for fi in targetGraphics:
        success = False #Simple boolean to track download success
        module = parseSubmodule(fi)[0] #Get module from file name

        eel.updateDownloadUI(fi, 'downloading') #Update UI (this is a JS function)
        for ext in extensions:
            try:
                #Get file from API
                response = web.get_file_by_server_relative_path(f"/sites/MediaLibrary/Projects/Part 66/{module_dict[module]['name']}/{fi}{ext}")
                download_path = os.path.join(f'{MODULE_PATH}/Content/Resources/Images/New/', f'{fi}{ext}')

                with open(download_path, "wb") as local_file:
                    response.download(local_file).execute_query() #Execute download query (THIS LINE WILL THROW THE HTTP ERRORS)

                success = True

            except:
                #eel.updateDownloadUI(fi, 'failed')
                try:
                    os.remove(download_path) #File is already written in local ready to receive download, delete as file not found
                except (FileNotFoundError, UnboundLocalError) as e:
                    pass
                continue

        if success:
            eel.updateDownloadUI(fi, 'downloaded')
        else:
            eel.updateDownloadUI(fi, 'failed')

def outputReport(actionSet):
    """
    Function that builds data profile of images already in the New folder, without API contact

        - Prints to logger the percentage of successfully embedded images;
        - Generates txt report for images that are not used;
        - Generates json for more in depth info.

    Input:

        - actionSet (dict): Set of information about what actions to take from script.js, including image locations and login details. See js_getActions().
    
    No return.
    """

    eel.log(f'Building report for {MODULE_PATH}...')

    #Break if module path doesn't exist/Get module folder from path selected in UI
    if (os.path.exists(MODULE_PATH) == False) or (not any(f.endswith('.flprj') for f in os.listdir(MODULE_PATH))):
        eel.log(f'Unable to find a project at {MODULE_PATH}. Please try again, selecting a folder with a valid ".flprj" file inside.')
        return None

    #Get all images with png/jpe?g extensions from Images/New
    images = [i for i in os.listdir(f'{MODULE_PATH}/Content/Resources/Images/New') if i.lower().endswith(('.png', '.jpg', '.jpeg'))]

    # #Full scan of all docs in project
    resultDict = checkProjectImages(images)
 
    eel.log(f'There are a total of {len(images)} images (ending PNG, JPG, JPEG) in the New folder.')

    dictImages = list(resultDict.keys())

    if len(images) > 0:
        percent = 100*(len(dictImages)/len(images)) #Images found in project (embedded or not)
        eel.log(f'{percent:.2f}% of the images were detected in the project!')
    else:
        eel.log('No images were found in the New folder.')

    #Images not found in project but ARE in New
    notFound = [i for i in images if i not in dictImages]
    eel.log(f'{len(notFound)} images were not found in the project.')

    #Count images found but not embedded (failCount)
    failCount = 0
    fails = []
    for i in resultDict.items():
        if i[1]['emb'] == False:
            failCount += 1
            fails.append(f"{i[0]}:\n\t\tPAGE:{i[1]['doc']}\n\t\tLINE:{i[1]['ind']}")

    eel.log(f'Images failed to embed: {failCount}')
    eel.log(f'<b>TOTAL SUCCESS RATE: {100*((len(dictImages) - failCount)/len(images)):.2f}%</b>')
    eel.log('See the <b>embedStats</b> JSON and <b>Unembedded</b> TXT file for the full information, both under FlareResults on your Desktop.')

    if os.path.exists(f'{DESKTOP_PATH}/FlareResults') == False:
        os.mkdir(f'{DESKTOP_PATH}/FlareResults')

    #Stats JSON saved to FlareResults
    with open(f'{DESKTOP_PATH}/FlareResults/embedStats.json', 'w') as f:
        json.dump(resultDict, f, indent=2)

    #Unembedded image TXT report
    with open(f'{DESKTOP_PATH}/FlareResults/Unembedded.txt', 'w', encoding='utf-8') as f:
        f.write(f'NOT FOUND IN {actionSet["actionDetails"]["moduleSelect"]} PROJECT:\n')
        if len(notFound) > 0:
            for i in notFound:
                f.write(f'\t{i}\n')
        else:
            f.write('\tNone\n')
        
        f.write(f'\n\nFAILED EMBED IN {actionSet["actionDetails"]["moduleSelect"]}:\n')
        if len(fails) > 0:
            for i in fails:
                f.write(f'\t{i}\n')
        else:
            f.write('\tNone\n')

def updateMonday(actionSet):
    """
    Matches any images in the New folder to images already embedded in the project by checking the HTML files.
    If any are found to match, tries to update the Upload To Flare board by changing the image item status with an API call.

    The input 'actionSet' is not used, but is needed because of how py_getActions triggers the function.
    """
    #Monday request header and context
    (mondayKey, mondayUrl) = (mondayConfig['monday']['login']['apiKey'], mondayConfig['monday']['login']['apiUrl'])
    mondayHeaders = {"Authorization" : mondayKey}

    #Break if module path doesn't exist/Get module folder from path selected in UI
    if (os.path.exists(MODULE_PATH) == False) or (not any(f.endswith('.flprj') for f in os.listdir(MODULE_PATH))):
        eel.log(f'Unable to find a project at {MODULE_PATH}. Please try again, selecting a folder with a valid ".flprj" file inside.')
        return None

    eel.log('Re-gathering image names from <b>Upload to Flare</b>...')
    boardID = mondayConfig['monday']['boardIDs']['Upload to Flare']

    #Get item names AND IDs from the Upload to Flare board (Only ever check this board)
    queryFirstLine = "{{ boards (ids:[{0}]) {{".format(boardID)
    query = queryFirstLine + """name
            items {

                name
                id
                } 
            } 
        }"""
    data = {'query' : query}

    #Post request, parse JSON response
    try:
        r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)
        returned = r.json()['data']['boards']
    except:
        eel.log('Sorry, you do not appear to be connected to the internet. Please check your connection.')
        return None

    #Build a list of item names from the Monday board
    targetGraphics = [(i['name'], i['id']) for i in returned[fetchDataIndex(returned, 'Upload to Flare')]['items']]
    graphicNames = [i[0] for i in targetGraphics]

    #Get all images with png/jpe?g extensions from Images/New
    images = [i for i in os.listdir(f'{MODULE_PATH}/Content/Resources/Images/New') if i.lower().endswith(('.png', '.jpg', '.jpeg'))]

    #Generate embed dict
    resultDict = checkProjectImages(images)

    #Get all definitely embedded images from dict
    embeddedImages = [os.path.splitext(i[0])[0] for i in list(resultDict.items()) if i[1]['emb'] == True]

    #Match them to graphics on Upload to Flare board
    matching = [i for i in embeddedImages if i in graphicNames]
    newStatus = "Media Inserted"

    if len(matching) < 1:
        eel.log('The Upload to Flare board looks up to date! No images were found on the board that have been embedded already.')

    for i in graphicNames:
        if i in embeddedImages:
            print(f'FOUND: {i}')

    for name in matching:
        graphicID = [i[1] for i in targetGraphics if i[0] == name][0]  #Get monday item ID, Lazy and intensive way to do it but pretty harmless here
        
        eel.updateMondayUI(name)

        #Request to change status of each item
        query = 'mutation {{change_simple_column_value (board_id: {0}, item_id: {1}, column_id: "status", value: "{2}") {{id}} }}'.format(boardID, graphicID, newStatus)

        #Post request!
        data = {'query' : query}
        r = requests.post(url=mondayUrl, json=data, headers=mondayHeaders)

def embedGraphics(actionSet):
    """
    First builds list of images from the New folder with the target submodule in the name.
    Second builds list of HTML files in the target submodule folder in the project.
    Finally, builds a graphic replace map for each HTML doc and embeds the correct images, with input
        
        - actionSet (dict): Set of information about what actions to take from script.js, including image locations and login details. See js_getActions().
    
    Returns nothing, but writes new embedded files to {fileName}-embedded.htm for a sanity check; also updates frontend UI with info on embedding.
    """
    
    targetSubmodule = actionSet['actionDetails']['submoduleSelect']

    #Path to New folder, images in folder, images with correct submodule indicator
    folderPath = f'{MODULE_PATH}/Content/Resources/Images/New/'
    imageFolder = os.listdir(folderPath)
    images = [i for i in imageFolder if parseSubmodule(i)[1].startswith(targetSubmodule)]

    #Unfortunately we have to walk() to get the path
    for root, dirs, files in os.walk(f'{MODULE_PATH}/Content/'):
        for d in dirs:
            if d == targetSubmodule: #NOTE: Flare folder MUST be named the submodule number ONLY, e.g. 7.5
                targetFolder = os.path.join(root, d) #Our path to target submodule folder

    #Build the GRM. Has structure {file: {line: image, ...}, ...} for all html files under target folder EXCLUDING ones that have '-embedded' in the name
    #Available at grm.json (dumped, indent 2)
    graphicReplaceMap = {}
    try:
        for root, dirs, files in os.walk(targetFolder):
            for f in files:
                if '-embedded' in f: #Skip files already embedded
                    continue

                filePath = os.path.join(root, f)

                with open(filePath, 'r', encoding="utf8") as f:
                    line_list = f.readlines()

                graphicReplaceMap[f.name] = {} #Add file to GRM

                for index, l in enumerate(line_list):
                    if any(os.path.splitext(graphic)[0] in l for graphic in images):
                        if '<img' in l: #Skip images already in figures
                            continue
                        
                        im = [i for i in images if os.path.splitext(i)[0] in l][0]

                        dirPath = os.path.dirname(os.path.realpath(f.name)) #Absolute path to current file's dir
                        relPath = os.path.relpath(folderPath, dirPath) #Relative path from dir to images folder

                        graphicReplaceMap[f.name][index] = os.path.join(relPath,im)

    except UnboundLocalError: #No folder specifically matching chosen submodule

        #Use regex to pick up next module selector up (e.g. 12.4 for 12.4.1 given)
        reg = "^(.*?((([0-9]{1,}\.[0-9]{1,}[a-zA-Z]?)(\.[0-9]{1,}[a-zA-Z]?)?)(\.[0-9]{1,}[a-zA-Z]?)?).*)$"
        found = re.findall(reg, targetSubmodule)
        nextUp = next(i for i in found[0] if i != targetSubmodule) #First regex match that is not the target submodule (next one up)

        eel.log(f'Unable to find a folder called <b>{targetSubmodule}</b> in the project. Please select the correct folder for your target submodule (<i>Hint: it might be <i><b>{nextUp}</b></i>).')

    #Save the GRM for debug and info
    with open('grm.json', 'w') as j:
        json.dump(graphicReplaceMap, j, indent=2)

    #Iterate through files in the GRM
    for file in list(graphicReplaceMap.keys()):
        if len(graphicReplaceMap[file].keys()) < 1: #No graphics to embed in this file, skip
            continue

        with open(file, 'r', encoding="utf8") as f: #Not skipped, read the file line by line
            line_list = f.readlines()

        graphicInds = list(graphicReplaceMap[file].keys()) #Line locations of graphics in this file

        #New empty HTML to write the embedded file
        newDoc = []
        newDoc += line_list[:graphicInds[0]] #Add text up to first graphic location

        if len(graphicInds) == 1:
            #For UI logging
            grmContent = [graphicInds[0], graphicReplaceMap[file][graphicInds[0]].rsplit('\\',1)[-1]]

            #Check for MC conditions
            conditions = ''
            condRegex = 'MadCap:conditions=".*"'
            found = re.findall(condRegex, line_list[graphicInds[0]])
            
            #Only worry about one condition set found                
            if len(found) == 1:
                conditions = f' {found[0]}' #PLEASE LEAVE SPACE HERE, separates attr from figure tag
                caption = deleteMultiple(line_list[graphicInds[0]+1], [f'<p{conditions}>', '</p>','  ','\n','\t','<figcaption>','</figcaption>'])
            else:
                caption = deleteMultiple(line_list[graphicInds[0]+1], [f'<p>', '</p>','  ','\n','\t','<figcaption>','</figcaption>'])
            
            newDoc += generateFigure(graphicReplaceMap[file][graphicInds[0]], caption, conditions)
            
            filePathEnd = "{0} - {1}".format(file.rsplit('\\', 2)[-2], file.rsplit('\\', 2)[-1])
            eel.updateEmbedUI(filePathEnd, [grmContent]) #Make grmContent a list here for the way the JS function works (uses forEach)
            
            newDoc += line_list[graphicInds[0]+2:] # +2 removes the lines where it detected the graphic (no longer needed)

        else:
            #For UI logging
            grmContent = []

            for i, g in enumerate(graphicInds):

                grmContent.append([g, graphicReplaceMap[file][g].rsplit('\\',1)[-1]])
                
                #Check for MC conditions
                conditions = ''
                condRegex = 'MadCap:conditions=".*"'
                found = re.findall(condRegex, line_list[g])
                
                #Only worry about one condition set found                
                if len(found) == 1:
                    conditions = f' {found[0]}' #PLEASE LEAVE SPACE HERE, separates attr from figure tag
                    caption = deleteMultiple(line_list[g+1], [f'<p{conditions}>', '</p>','  ','\n','\t','<figcaption>','</figcaption>'])
                else:
                    caption = deleteMultiple(line_list[g+1], [f'<p>', '</p>','  ','\n','\t','<figcaption>','</figcaption>'])
                
                newDoc += generateFigure(graphicReplaceMap[file][g], caption, conditions)

                if i == len(graphicInds)-1:
                    newDoc += line_list[g+2:]
                else:
                    newDoc += line_list[g+2:graphicInds[i+1]] # +2 removes the lines where it detected the graphic (no longer needed)

            filePathEnd = "{0} - {1}".format(file.rsplit('\\', 2)[-2], file.rsplit('\\', 2)[-1])
            eel.updateEmbedUI(filePathEnd, grmContent)

        with open(file.replace('.htm', '-embedded.htm'), 'w', encoding='utf8') as f:
            for l in newDoc:
                f.write(l)

eel.start("index.html", size=(800, 850))