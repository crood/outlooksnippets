// taskpane.js

// Ensure Office is ready before doing anything
Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        initializeAddin();
    } else {
        console.error("This add-in only works in Outlook.");
        // Optionally display a message to the user in the task pane
        document.getElementById('snippet-list-container').innerHTML = '<p class="message error">This add-in requires Outlook.</p>';
    }
});

const SNIPPETS_KEY = 'cw_snippets_v1'; // Prefix to avoid potential collisions

// DOM References (declared globally within the module scope)
let snippetForm, commandInput, descriptionInput, textInput, saveButton, snippetListContainer, formMessage, listMessage, formTitle, editOriginalCommandInput, cancelEditButton, btnAddNewline, btnAddPlaceholder;

let currentSnippets = []; // Cache snippets locally

function initializeAddin() {
    // Get DOM references after Office is ready and DOM is loaded
    snippetForm = document.getElementById('snippet-form');
    commandInput = document.getElementById('command');
    descriptionInput = document.getElementById('description');
    textInput = document.getElementById('text');
    saveButton = document.getElementById('save-button');
    snippetListContainer = document.getElementById('snippet-list-container');
    formMessage = document.getElementById('form-message');
    listMessage = document.getElementById('list-message');
    formTitle = document.getElementById('form-title');
    editOriginalCommandInput = document.getElementById('edit-original-command');
    cancelEditButton = document.getElementById('cancel-edit-button');
    btnAddNewline = document.getElementById('btn-add-newline');
    btnAddPlaceholder = document.getElementById('btn-add-placeholder');

    // Attach Event Listeners
    attachEventListeners();

    // Load initial snippets
    loadSnippetsFromSettings();
}

// --- Storage Functions (using Office Roaming Settings) ---

async function loadDefaultSnippets() {
    try {
        const response = await fetch('../assets/snippets.json'); // Adjust path if needed
        if (!response.ok) {
            throw new Error(`HTTP error loading defaults: ${response.status}`);
        }
        const defaults = await response.json();
        console.log("Loaded default snippets:", defaults);
        return defaults;
    } catch (error) {
        console.error("Failed to load default snippets:", error);
        showMessage(listMessage, "Error loading default snippets.", true);
        return []; // Return empty array on failure
    }
}

function loadSnippetsFromSettings() {
    setMessage(snippetListContainer, '<p class="status-message">Loading snippets...</p>');
    Office.context.roamingSettings.remove(SNIPPETS_KEY + '_error_flag'); // Clear previous errors

    const storedSnippets = Office.context.roamingSettings.get(SNIPPETS_KEY);

    if (storedSnippets) {
        try {
            currentSnippets = JSON.parse(storedSnippets);
            console.log("Snippets loaded from roaming settings:", currentSnippets);
            renderSnippetList(currentSnippets);
        } catch (e) {
            console.error("Error parsing snippets from settings:", e);
            showMessage(listMessage, "Error loading snippets from storage. Loading defaults.", true);
            // Mark error and load defaults on next load
             Office.context.roamingSettings.set(SNIPPETS_KEY + '_error_flag', 'true');
             Office.context.roamingSettings.saveAsync( () => loadDefaultsAndSave());
        }
    } else {
        console.log("No snippets in settings, loading defaults.");
        // Check if we previously failed, to avoid infinite loop if defaults also fail
         if (Office.context.roamingSettings.get(SNIPPETS_KEY + '_error_flag')) {
             showMessage(listMessage, "Failed to load stored snippets previously. Manual reset might be needed.", true);
             renderSnippetList([]); // Show empty list
         } else {
             loadDefaultsAndSave();
         }
    }
}

async function loadDefaultsAndSave() {
     const defaults = await loadDefaultSnippets();
     currentSnippets = defaults;
     saveSnippetsToSettings(defaults, (success) => {
         if (success) {
             renderSnippetList(defaults);
             showMessage(listMessage, "Loaded default snippets.", false);
         } else {
             // Error message already shown by saveSnippetsToSettings
             renderSnippetList([]); // Show empty list if save failed
         }
     });
}


function saveSnippetsToSettings(snippets, callback) {
    try {
        const snippetsString = JSON.stringify(snippets);
        // Check size limit (approx 32KB for roaming settings value)
        if (snippetsString.length > 30000) { // Leave some buffer
             showMessage(listMessage, "Error: Snippets data too large to save.", true);
             if (callback) callback(false);
             return;
        }

        Office.context.roamingSettings.set(SNIPPETS_KEY, snippetsString);
        Office.context.roamingSettings.saveAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Snippets saved successfully.");
                 Office.context.roamingSettings.remove(SNIPPETS_KEY + '_error_flag'); // Clear error flag on success
                if (callback) callback(true);
            } else {
                console.error("Error saving snippets to roaming settings:", asyncResult.error.message);
                showMessage(listMessage, `Error saving snippets: ${asyncResult.error.message}`, true);
                if (callback) callback(false);
            }
        });
    } catch (e) {
         console.error("Error stringifying snippets:", e);
         showMessage(listMessage, "Error preparing snippets for saving.", true);
         if (callback) callback(false);
    }
}

// --- Utility Functions ---
function setMessage(element, htmlContent) {
    element.innerHTML = htmlContent;
}

function showMessage(element, message, isError = false) {
    element.textContent = message;
    element.className = `message ${isError ? 'error' : 'success'}`;
    element.classList.remove('hidden');
    // Auto-hide after a few seconds
    setTimeout(() => {
        element.classList.add('hidden');
        element.textContent = '';
    }, isError ? 5000 : 3000);
}

function clearForm() {
    snippetForm.reset();
    editOriginalCommandInput.value = ''; // Clear edit tracking
    formTitle.textContent = 'Create New Snippet';
    saveButton.textContent = 'Save Snippet';
    cancelEditButton.classList.add('hidden');
    formMessage.classList.add('hidden');
    commandInput.disabled = false; // Re-enable command input
}

function extractPlaceholders(text) {
    if (!text) return [];
    const regex = /\{([^}]+)\}/g;
    const matches = text.match(regex);
    if (!matches) return [];
    // Return unique, non-empty placeholder names
    return [...new Set(matches.map(p => p.slice(1, -1).trim()).filter(name => name))];
}

// --- Text Area Helpers ---
function insertAtCursor(textarea, textToInsert) {
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    textarea.value = text.substring(0, start) + textToInsert + text.substring(end);
    // Place cursor after inserted text
    textarea.selectionStart = textarea.selectionEnd = start + textToInsert.length;
    textarea.focus(); // Keep focus on textarea
    // Trigger input event in case anything listens for it
    textarea.dispatchEvent(new Event('input', { bubbles: true }));
}

// --- Snippet List Rendering ---
function renderSnippetList(snippets) {
    snippetListContainer.innerHTML = ''; // Clear previous list

    if (!snippets || snippets.length === 0) {
        setMessage(snippetListContainer, '<p class="status-message">No snippets found. Create one above!</p>');
        return;
    }

    snippets.forEach((snippet, index) => {
        const item = document.createElement('div');
        item.className = 'snippet-item';
        item.setAttribute('data-index', index); // Use index for easy lookup

        const info = document.createElement('div');
        info.className = 'snippet-info';
        info.innerHTML = `<strong>${snippet.command}</strong><span>${snippet.description || '(No description)'}</span>`;

        const actions = document.createElement('div');
        actions.className = 'snippet-actions';

        const insertBtn = document.createElement('button');
        insertBtn.textContent = 'Insert';
        insertBtn.className = 'insert-btn';
        insertBtn.type = 'button';

        const editBtn = document.createElement('button');
        editBtn.textContent = 'Edit';
        editBtn.className = 'edit-btn';
        editBtn.type = 'button';

        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Delete';
        deleteBtn.className = 'delete-btn';
        deleteBtn.type = 'button';

        actions.appendChild(insertBtn);
        actions.appendChild(editBtn);
        actions.appendChild(deleteBtn);

        // Placeholder for dynamic inputs
        const placeholderDiv = document.createElement('div');
        placeholderDiv.className = 'placeholder-inputs hidden'; // Hidden by default

        item.appendChild(info);
        item.appendChild(actions);
        item.appendChild(placeholderDiv); // Add the placeholder container
        snippetListContainer.appendChild(item);
    });
}


// --- Insertion Logic ---

function handleInsertClick(event) {
    const itemElement = event.target.closest('.snippet-item');
    if (!itemElement) return;

    const index = parseInt(itemElement.getAttribute('data-index'), 10);
    const snippet = currentSnippets[index];
    if (!snippet) {
        showMessage(listMessage, "Error: Could not find snippet data.", true);
        return;
    }

    const placeholders = extractPlaceholders(snippet.text);
    const placeholderDiv = itemElement.querySelector('.placeholder-inputs');

    // Hide any other open placeholder divs
    document.querySelectorAll('.placeholder-inputs').forEach(div => {
        if (div !== placeholderDiv) div.classList.add('hidden');
    });


    if (placeholders.length > 0) {
        // Show inputs for placeholders
        showPlaceholderInputs(snippet, placeholderDiv);
    } else {
        // Insert directly
        placeholderDiv.classList.add('hidden'); // Ensure it's hidden
        insertTextIntoEmail(snippet.text);
    }
}

function showPlaceholderInputs(snippet, placeholderDiv) {
    placeholderDiv.innerHTML = ''; // Clear previous inputs
    placeholderDiv.classList.remove('hidden');

    const placeholders = extractPlaceholders(snippet.text);
    const inputs = {}; // To store references

    placeholders.forEach(name => {
        const label = document.createElement('label');
        label.textContent = `${name}:`;
        label.htmlFor = `placeholder-input-${snippet.command}-${name}`; // Unique ID

        const input = document.createElement('input');
        input.type = 'text';
        input.id = label.htmlFor;
        input.name = name;
        input.className = 'placeholder-input-field';
        input.placeholder = `Enter value for {${name}}`;

        inputs[name] = input; // Store reference

        placeholderDiv.appendChild(label);
        placeholderDiv.appendChild(input);
    });

    const confirmButton = document.createElement('button');
    confirmButton.textContent = 'Confirm & Insert';
    confirmButton.type = 'button';
    confirmButton.className = 'confirm-insert-btn';
    confirmButton.addEventListener('click', () => handleConfirmInsertClick(snippet, inputs, placeholderDiv));

    const cancelButton = document.createElement('button');
    cancelButton.textContent = 'Cancel';
    cancelButton.type = 'button';
    cancelButton.className = 'cancel-insert-btn';
     cancelButton.addEventListener('click', () => placeholderDiv.classList.add('hidden'));


    const buttonWrapper = document.createElement('div');
    buttonWrapper.className = 'placeholder-button-wrapper';
    buttonWrapper.appendChild(cancelButton);
    buttonWrapper.appendChild(confirmButton);
    placeholderDiv.appendChild(buttonWrapper);


    // Focus the first input
    const firstInput = placeholderDiv.querySelector('input');
    if (firstInput) firstInput.focus();
}

function handleConfirmInsertClick(snippet, inputRefs, placeholderDiv) {
    const placeholderValues = {};
    let allFilled = true;
    for (const name in inputRefs) {
        placeholderValues[name] = inputRefs[name].value;
         // Basic check: make sure required placeholders aren't empty (optional)
         // if (!placeholderValues[name]) {
         //     allFilled = false;
         //     inputRefs[name].classList.add('input-error'); // Add visual cue
         // } else {
         //     inputRefs[name].classList.remove('input-error');
         // }
    }

    // if (!allFilled) {
    //     showMessage(listMessage, "Please fill in all placeholder values.", true);
    //     return;
    // }

    const compiledText = compileSnippet(snippet, placeholderValues);
    insertTextIntoEmail(compiledText);
    placeholderDiv.classList.add('hidden'); // Hide after insertion
}

function compileSnippet(snippet, placeholderValues) {
    let compiled = snippet.text;
    for (const name in placeholderValues) {
        // Regex to replace {placeholder_name}, handling potential whitespace
        const regex = new RegExp(`\\{\\s*${escapeRegExp(name)}\\s*\\}`, 'g');
        compiled = compiled.replace(regex, placeholderValues[name] || ''); // Replace with value or empty string
    }
    return compiled;
}

// Helper function to escape special characters for regex
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}


function insertTextIntoEmail(text) {
    if (!Office.context.mailbox.item) {
        showMessage(listMessage, "Error: Cannot insert text. No email item context found (are you in compose mode?).", true);
        return;
    }

    // Outlook body typically expects HTML. Convert newlines to <br>.
    // Preserve existing HTML structure if any - simple replace is basic.
    // Consider a more robust HTML sanitizer/parser if snippets contain complex HTML.
    const htmlText = text.replace(/\n/g, '<br/>');

    Office.context.mailbox.item.body.setSelectedDataAsync(
        htmlText,
        { coercionType: Office.CoercionType.Html }, // Insert as HTML
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(`Error inserting text: ${asyncResult.error.message}`);
                showMessage(listMessage, `Error inserting snippet: ${asyncResult.error.message}`, true);
            } else {
                console.log("Snippet inserted successfully.");
                showMessage(listMessage, "Snippet inserted!", false);
            }
        }
    );
}


// --- Event Handlers ---

function attachEventListeners() {
    // Text Area Helpers
    btnAddNewline.addEventListener('click', () => insertAtCursor(textInput, '\n\n'));
    btnAddPlaceholder.addEventListener('click', () => {
        const start = textInput.selectionStart;
        insertAtCursor(textInput, '{}');
        textInput.selectionStart = textInput.selectionEnd = start + 1; // Place cursor inside {}
    });

    // Form Submission (Create/Update)
    snippetForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const command = commandInput.value.trim();
        const description = descriptionInput.value.trim();
        const text = textInput.value; // Keep whitespace
        const originalCommand = editOriginalCommandInput.value; // Check if editing

        if (!command || !text) {
            showMessage(formMessage, "Command Name and Snippet Text are required.", true);
            return;
        }

        // Simple validation for command name (avoid special chars that might break things)
        if (!/^[a-zA-Z0-9_\-]+$/.test(command)) {
             showMessage(formMessage, "Command Name can only contain letters, numbers, underscores, and hyphens.", true);
             return;
        }


        const snippetData = { command, description, text };
        let updatedSnippets = [...currentSnippets]; // Create a copy

        if (originalCommand) {
            // --- Update Existing Snippet ---
            const indexToUpdate = updatedSnippets.findIndex(s => s.command === originalCommand);
            if (indexToUpdate === -1) {
                showMessage(formMessage, `Error: Original snippet "${originalCommand}" not found for update.`, true);
                return;
            }
            // Check if new command name conflicts (unless it's the same item)
             if (command !== originalCommand && updatedSnippets.some((s, i) => s.command === command && i !== indexToUpdate)) {
                 showMessage(formMessage, `Error: Command Name "${command}" already exists.`, true);
                 return;
             }
            updatedSnippets[indexToUpdate] = snippetData;
            saveSnippetsToSettings(updatedSnippets, (success) => {
                if (success) {
                    showMessage(formMessage, "Snippet updated successfully!", false);
                    currentSnippets = updatedSnippets; // Update local cache
                    clearForm();
                    renderSnippetList(currentSnippets); // Refresh list
                } // Error message handled by save function
            });

        } else {
            // --- Add New Snippet ---
            // Check for duplicate command before adding
             if (updatedSnippets.some(s => s.command === command)) {
                 showMessage(formMessage, `Error: Command Name "${command}" already exists.`, true);
                 return;
             }
            updatedSnippets.push(snippetData);
            saveSnippetsToSettings(updatedSnippets, (success) => {
                 if (success) {
                     showMessage(formMessage, "Snippet added successfully!", false);
                     currentSnippets = updatedSnippets; // Update local cache
                     clearForm();
                     renderSnippetList(currentSnippets); // Refresh list
                 } // Error message handled by save function
             });
        }
    });

    // List Actions (Edit/Delete/Insert - Event Delegation)
    snippetListContainer.addEventListener('click', (event) => {
        const target = event.target;
        const snippetItem = target.closest('.snippet-item');
        if (!snippetItem) return;

        const snippetIndex = parseInt(snippetItem.getAttribute('data-index'), 10);
        const snippet = currentSnippets[snippetIndex];

        if (!snippet) {
            console.error("Could not find snippet data for index:", snippetIndex);
            showMessage(listMessage, "Error: Could not find snippet data.", true);
            return;
        }

        if (target.classList.contains('edit-btn')) {
            // --- Populate form for editing ---
            formTitle.textContent = 'Edit Snippet';
            commandInput.value = snippet.command;
            descriptionInput.value = snippet.description;
            textInput.value = snippet.text;
            editOriginalCommandInput.value = snippet.command; // Track original command for update logic
            // commandInput.disabled = true; // Optionally disable editing command name directly
            saveButton.textContent = 'Update Snippet';
            cancelEditButton.classList.remove('hidden');
            formMessage.classList.add('hidden'); // Clear previous form messages
            snippetItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); // Scroll list item into view
            window.scrollTo(0, 0); // Scroll task pane to top
            commandInput.focus();
             // Hide any open placeholder inputs
            document.querySelectorAll('.placeholder-inputs').forEach(div => div.classList.add('hidden'));
        }
        else if (target.classList.contains('delete-btn')) {
            // --- Delete Snippet ---
            // Use Office UI dialog for confirmation if possible, otherwise browser confirm
            if (confirm(`Are you sure you want to delete the snippet "${snippet.command}"?`)) {
                let updatedSnippets = currentSnippets.filter((s, i) => i !== snippetIndex);
                saveSnippetsToSettings(updatedSnippets, (success) => {
                    if (success) {
                        showMessage(listMessage, "Snippet deleted successfully!", false);
                        currentSnippets = updatedSnippets; // Update cache
                        renderSnippetList(currentSnippets); // Refresh list
                        // If the deleted item was being edited, clear the form
                         if (editOriginalCommandInput.value === snippet.command) {
                             clearForm();
                         }
                    } // Error handled by save function
                });
            }
        }
         else if (target.classList.contains('insert-btn')) {
            // --- Handle Insert Click ---
             handleInsertClick(event);
         }
    });

    // Cancel Edit Button
    cancelEditButton.addEventListener('click', () => {
        clearForm();
    });
}