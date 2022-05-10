let MODULE_PATH = '/';
const LOGGER = document.getElementById('logWindow');

function js_writeLoginDetails() {
    eel.py_writeLoginDetails(getLogins());
}

/**
 * Setter for the python AND JS global variable MODULE_PATH
 * @param {} modulePath System path to the project folder to set as the global module path
 */
function js_setModulePath(modulePath) {
    MODULE_PATH = `/${modulePath}`

    eel.py_setModulePath(`/${modulePath}`); //Set it in the Py backend
    eel.compileSubmoduleChoices(`/${modulePath}`)(updateSubmodules); //Get the built submodules
}

const updateSubmodules = list => {
    selectEl = document.getElementById('submoduleSelect');
    selectEl.innerHTML = ''; //Clear submodule select dropdown options
    
    list.forEach(sm => {
        selectEl.innerHTML += `<option value="${sm}">${sm}</option>` //Add an option for each submodule built in Python
    });
}

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

    // eel.queryBoardID(actionDetails['boardSelect'])(handleQuery);
}

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

const handleQuery = res => {
    if (typeof res === 'object') {
        e = new MondayError(res);
        log(e.msg);
    } else {
        log(res);
    }
}

const updateValue = (id, value) => {
    let target = document.getElementById(id);
    target.value = value;
}

eel.expose(buildDownloadList)
function buildDownloadList(graphics) {
    graphics.forEach(g => {
        LOGGER.innerHTML += `<div name="${g}" class="pending"><span>${g}</span><p>Pending</p></div>`;
    });
}

eel.expose(updateDownloadUI)
function updateDownloadUI(graphic, status) {
    let el = document.querySelector(`div[name="${graphic}"]`);

    //Give it new class according to graphic download status and update item subtitle
    el.className = status;
    el.lastChild.innerText = status;
}

eel.expose(updateMondayUI)
function updateMondayUI(graphic) {
    LOGGER.innerHTML += `<div name="${graphic}" class="mondayUpdate"><span>${graphic}</span><p>Media Inserted >>> Sort into Modules</p></div>`
}

eel.expose(updateEmbedUI)
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
function log(text) {
    if (text === 'You haven\'t selected any actions.') {
        LOGGER.innerHTML = `<p>${text}</p>`;
    } else {
        LOGGER.innerHTML += `<p>${text}</p>`;
    }
}

eel.expose(raiseError);
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