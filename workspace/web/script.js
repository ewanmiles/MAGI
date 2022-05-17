let MODULE_PATH = '/';
const LOGGER = document.getElementById('logWindow');

function js_writeLoginDetails() {
    eel.py_writeLoginDetails(getLogins());
}

/**
 * Setter for the python AND JS global variable MODULE_PATH
 * @param {str} modulePath System path to the project folder to set as the global module path
 */
function js_setModulePath(modulePath) {
    MODULE_PATH = `/${modulePath}`

    eel.py_setModulePath(`/${modulePath}`); //Set it in the Py backend
    eel.compileSubmoduleChoices(`/${modulePath}`)(updateSubmodules); //Get the built submodules
}

/**
 * Updates the submodule select box with the submodule choices compiled by the Py backend
 * @param {str[]} list Set of submodule options to add to the select
 */
const updateSubmodules = list => {
    selectEl = document.getElementById('submoduleSelect');
    selectEl.innerHTML = ''; //Clear submodule select dropdown options
    
    list.forEach(sm => {
        selectEl.innerHTML += `<option value="${sm}">${sm}</option>` //Add an option for each submodule built in Python
    });
}

/**
 * Gets the set of actions chosen in the UI by checking the radio/select/checkbox inputs, sends them to Py backend
 * 
 * Action set is sent as a dictionary, unpacked by a sort of mapping function in Python, including login details, 
 * selected actions from the checkboxes, image location from the radio, action details from the selects. 
 */
function js_getActions() {

    let imgLocation = document.querySelector('input[name="imageLocation"]:checked').id;
    let actions = getActions();
    let actionDetails = getActionDetails();

    //Create login context before executing actions if necessary
    if (imgLocation === 'sharepoint') {
        eel.createLoginContext(getLogins())(handleContext);
    };

    eel.py_setModulePath(`/${actionDetails['moduleSelect']}`); //Double sure that the right module path set in py backend

    LOGGER.innerHTML = '';

    eel.py_getActions({
        'loginDetails': getLogins(),
        'actions': actions,
        'imageLocation': imgLocation,
        'actionDetails': actionDetails
    });
}

/**
 * Callback function for when a sharepoint login context is created in the backend, raises errors to the logger if fail.
 * @param {str} res Whatever is returned from the login context creation call, whether it triggers an error or is successful. See createLoginContext() in Python.
 */
const handleContext = res => {
    if (res === 'Connection') {
        e = new ConnectionError();
        LOGGER.innerHTML = `<p>${e.msg}</p>`;
    } else if (typeof res === 'object') {
        e = new LoginError(res[1]);
        LOGGER.innerHTML = `<p>${e.msg}</p>`;
    } else {
        LOGGER.innerHTML = `<p>${res}</p>`;
    }
}

/**
 * General function to update the value of any target with given id.
 * 
 * NOTE: Obviously not error proof. Will raise an error for setting the value of something that lacks a value attr.
 * 
 * @param {str} id ID of target grabbed from DOM using getElementById()
 * @param {str} value New value to set for target
 */
const updateValue = (id, value) => {
    let target = document.getElementById(id);
    target.value = value;
}


eel.expose(buildDownloadList)
/**
 * Purely aesthetic. Builds UI items for each of the graphics in the given download list.
 * @param {str[]} graphics List of graphics that will be/have been downloaded
 */
function buildDownloadList(graphics) {
    graphics.forEach(g => {
        LOGGER.innerHTML += `<div name="${g}" class="pending"><span>${g}</span><p>Pending</p></div>`;
    });
}

eel.expose(updateDownloadUI)
/**
 * Purely aesthetic. Updates graphic item UI to new download status of given graphic by matching the graphic name with the correct div name attr.
 * @param {str} graphic Graphic name to update - will match given querySelector(div[name=graphic])
 * @param {str} status Status of the download, e.g. Pending, Failed, Downloaded
 */
function updateDownloadUI(graphic, status) {
    let el = document.querySelector(`div[name="${graphic}"]`);

    //Give it new class according to graphic download status and update item subtitle
    el.className = status;
    el.lastChild.innerText = status;
}

eel.expose(updateMondayUI)
/**
 * Purely aesthetic. Builds UI items for each graphic that is successfully updated on the Monday board.
 * @param {str} graphic Graphic name to build item for
 */
function updateMondayUI(graphic) {
    LOGGER.innerHTML += `<div name="${graphic}" class="mondayUpdate"><span>${graphic}</span><p>Media Inserted >>> Sort into Modules</p></div>`
}

eel.expose(updateEmbedUI)
/**
 * Purely aesthetic. Updates the UI with information about where the graphics have been successfully embedded.
 * @param {str} file HTML file name where graphics have been embedded
 * @param {str[]} list List of graphics that have been embedded. List entries are two element lists - [0] is HTML file line of embed, [1] is graphic name
 */
function updateEmbedUI(file, list) {
    let el = document.createElement('div');
    el.className = 'embeds';
    el.innerHTML = `<span>${file}</span>`;

    list.forEach(graphic => {
       el.innerHTML += `<p><b>${graphic[0]}</b>: ${graphic[1]}</p>`;
    });

    LOGGER.innerHTML += el.outerHTML;
}

eel.expose(log);
/**
 * Basic logging function for the log window. Adds <p> tag to the window with given text.
 * @param {str} text String to log to the window
 */
function log(text) {
    if (text === 'You haven\'t selected any actions.') {
        LOGGER.innerHTML = `<p>${text}</p>`;
    } else {
        LOGGER.innerHTML += `<p>${text}</p>`;
    }
}

eel.expose(raiseError);
/**
 * The beginnings of an obviously unfinished and possibly unnecessary error handling system.
 * @param {str} ref Reference type of error to raise, e.g. Monday, Login, Sharepoint
 * @param {str} content Info to pass to errorHandling.js to produce correct error msg
 */
function raiseError(ref, content) {
    switch(ref) {
        case 'Monday':
            e = new MondayError(content);
    };

    log(e.msg);
}

/**
 * Receives login fields from JSON unpacked in the Python backend and updates the GUI
 * @param {} dict Dictionary containing JSON fields, spName, spPass, mKey as keys with string values
 */
const unpackLogins = dict => {
    //Cycle through login dict keys, update inputs on GUI
    Object.keys(dict).forEach(k => {
        updateValue(k, dict[k]);
    });
}

/**
 * Getter for the login details box
 * @returns object (=> dict) containing spName, spPass, mKey keys with string values
 */
const getLogins = () => {
    return {
        'spName': document.getElementById('spUser').value,
        'spPass': document.getElementById('spPass').value,
        'mKey': document.getElementById('mKey').value
    };
}

const getActions = () => {
    let out = [];

    //Returns DOM elements
    document.querySelectorAll('input[name="actions"]:checked').forEach(el => {
        out.push(el.id); //Get ID from each el
    });

    return out;
}

const getActionDetails = () => {
    let out = {};
    const IDs = ['moduleSelect', 'submoduleSelect', 'boardSelect'];

    IDs.forEach(el => {
        out[el] = document.getElementById(el).value;
    });

    return out;
}

//Fill in the GUI login details from JSONs unpacked in Python
eel.fillLoginDetails()(unpackLogins);