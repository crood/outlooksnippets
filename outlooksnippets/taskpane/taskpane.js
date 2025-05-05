// taskpane.js

// Ensure Office is ready before doing anything
Office.onReady(info => {
    // Log host info for debugging
    console.log("Office.onReady Info:", info);
    if (info.host === Office.HostType.Outlook) {
        console.log("Host is Outlook. Initializing...");
        initializeAddin();
    } else {
        console.warn("Host is NOT Outlook:", info.host);
        const container = document.getElementById('snippet-list-container');
        if (container) {
            container.innerHTML = '<p class="message error">This add-in requires Outlook.</p>';
        }
    }
});

const SNIPPETS_KEY = 'cw_snippets_v1'; // Prefix to avoid potential collisions

// --- DOM References ---
let snippetForm, commandInput, descriptionInput, textInput, saveButton, snippetListContainer, formMessage, listMessage, formTitle, editOriginalCommandInput, cancelEditButton, btnAddNewline, btnAddPlaceholder, toggleFormButton, searchInput, formSection;

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
    toggleFormButton = document.getElementById('toggle-form-btn');
    searchInput = document.getElementById('search-snippets');
    formSection = document.querySelector('.form-section'); // Get the section itself

    // Initial state: Hide the form section
    if (formSection) {
        formSection.classList.add('hidden'); // Start hidden
    } else {
        console.error("Form section element not found!");
    }

    // Attach Event Listeners
    attachEventListeners();

    // Load initial snippets
    loadSnippetsFromSettings();
}

// --- Storage Functions ---

async function loadDefaultSnippets() {
    try {
        // Assuming snippets.json is in an 'assets' folder relative to the taskpane folder
        const response = await fetch('../assets/snippets.json');
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
    // Ensure container exists before trying to update it
    if (!snippetListContainer) {
        console.error("Snippet list container not found during load!");
        return;
    }
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
    if (element) element.innerHTML = htmlContent;
}

function showMessage(element, message, isError = false) {
    if (!element) return; // Check if element exists
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
    if (!snippetForm) return;
    snippetForm.reset();
    if (editOriginalCommandInput) editOriginalCommandInput.value = ''; // Clear edit tracking
    if (formTitle) formTitle.textContent = 'Create New Snippet';
    if (saveButton) saveButton.textContent = 'Save Snippet';
    if (cancelEditButton) cancelEditButton.classList.add('hidden');
    if (formMessage) formMessage.classList.add('hidden');
    if (commandInput) commandInput.disabled = false; // Re-enable command input
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
     if (!textarea) return;
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
    if (!snippetListContainer) return;
    snippetListContainer.innerHTML = ''; // Clear previous list

    if (!snippets || snippets.length === 0) {
        setMessage(snippetListContainer, '<p class="status-message">No snippets found. Create one above!</p>');
        return;
    }

    // Filter based on search input BEFORE rendering
    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';

    snippets.forEach((snippet, index) => {
        // --- Filtering Logic ---
        const command = snippet.command.toLowerCase();
        const description = (snippet.description || '').toLowerCase();
        const text = snippet.text.toLowerCase(); // Also search snippet text itself? Optional.
        const isMatch = searchTerm === '' || command.includes(searchTerm) || description.includes(searchTerm) || text.includes(searchTerm);

        // Create elements only if it's a match (or no search term)
        if (isMatch) {
            const item = document.createElement('div');
            item.className = 'snippet-item';
            item.setAttribute('data-index', index); // Use index for easy lookup

            // Wrapper for flex layout
            const contentWrapper = document.createElement('div');
            contentWrapper.className = 'snippet-content-wrapper';

            const info = document.createElement('div');
            info.className = 'snippet-info';
            info.innerHTML = `<strong>${snippet.command}</strong><span>${snippet.description || '(No description)'}</span>`;

            // Wrapper for controls (actions and confirm)
            const controlsWrapper = document.createElement('div');
            controlsWrapper.className = 'snippet-controls-wrapper';

            const actions = document.createElement('div');
            actions.className = 'snippet-actions'; // Container for normal actions

            // --- Icon Button Creation ---
            const insertBtn = document.createElement('button');
            insertBtn.className = 'action-btn insert-btn'; // Common class for styling/selection
            insertBtn.type = 'button';
            insertBtn.title = 'Insert Snippet';
            insertBtn.innerHTML = `<span class="material-symbols-outlined">data_object</span>`; // Or 'add_comment', 'input'

            const editBtn = document.createElement('button');
            editBtn.className = 'action-btn edit-btn';
            editBtn.type = 'button';
            editBtn.title = 'Edit Snippet';
            editBtn.innerHTML = `<span class="material-symbols-outlined">edit</span>`;

            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'action-btn delete-btn';
            deleteBtn.type = 'button';
            deleteBtn.title = 'Delete Snippet';
            deleteBtn.innerHTML = `<span class="material-symbols-outlined">delete</span>`;
            // --- End Icon Button Creation ---

            actions.appendChild(insertBtn);
            actions.appendChild(editBtn);
            actions.appendChild(deleteBtn);

            // --- Add Confirmation UI Div (Hidden) ---
            const confirmDiv = document.createElement('div');
            confirmDiv.className = 'snippet-confirm-delete hidden'; // Hidden by default
            confirmDiv.innerHTML = `
                <span class="confirm-text">Delete snippet?</span>
                <button type="button" class="confirm-yes-btn" title="Confirm Delete">
                    <span class="material-symbols-outlined">check</span>
                </button>
                <button type="button" class="confirm-no-btn" title="Cancel">
                    <span class="material-symbols-outlined">close</span>
                </button>
            `;
            // --- End Confirmation UI Div ---

            controlsWrapper.appendChild(actions);      // Normal actions
            controlsWrapper.appendChild(confirmDiv);   // Confirmation UI (initially hidden)

            contentWrapper.appendChild(info);          // Add info to main wrapper
            contentWrapper.appendChild(controlsWrapper); // Add controls block to main wrapper

            // Placeholder for dynamic inputs (below info/actions)
            const placeholderDiv = document.createElement('div');
            placeholderDiv.className = 'placeholder-inputs hidden';

            item.appendChild(contentWrapper); // Add the content wrapper (info + controls)
            item.appendChild(placeholderDiv); // Add the placeholder container below
            snippetListContainer.appendChild(item);
        }
    });

     // Show message if search yields no results but snippets exist
     if (snippetListContainer.children.length === 0 && searchTerm !== '' && snippets.length > 0) {
         setMessage(snippetListContainer, '<p class="status-message">No snippets match your search.</p>');
     }
     // Show message if list is empty and no search term
     else if (snippetListContainer.children.length === 0 && searchTerm === '' && snippets.length === 0) {
          setMessage(snippetListContainer, '<p class="status-message">No snippets found. Create one above!</p>');
     }
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
    if (!placeholderDiv) return; // Check if div exists

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
     if (!placeholderDiv) return;
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
    confirmButton.className = 'confirm-insert-btn'; // Uses primary button style
    confirmButton.addEventListener('click', () => handleConfirmInsertClick(snippet, inputs, placeholderDiv));

    const cancelButton = document.createElement('button');
    cancelButton.textContent = 'Cancel';
    cancelButton.type = 'button';
    cancelButton.className = 'cancel-insert-btn'; // Uses secondary button style
     cancelButton.addEventListener('click', () => placeholderDiv.classList.add('hidden'));

    const buttonWrapper = document.createElement('div');
    buttonWrapper.className = 'placeholder-button-wrapper';
    buttonWrapper.appendChild(cancelButton); // Cancel first
    buttonWrapper.appendChild(confirmButton); // Then confirm
    placeholderDiv.appendChild(buttonWrapper);

    // Focus the first input
    const firstInput = placeholderDiv.querySelector('input');
    if (firstInput) firstInput.focus();
}

function handleConfirmInsertClick(snippet, inputRefs, placeholderDiv) {
    const placeholderValues = {};
    for (const name in inputRefs) {
        placeholderValues[name] = inputRefs[name].value;
    }
    const compiledText = compileSnippet(snippet, placeholderValues);
    insertTextIntoEmail(compiledText);
     if (placeholderDiv) placeholderDiv.classList.add('hidden'); // Hide after insertion
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
    // Check context exists before proceeding
    if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        showMessage(listMessage, "Error: Cannot insert text. Mailbox item context not available.", true);
        console.error("Mailbox item context is not available for insertion.");
        return;
    }

    // Outlook body typically expects HTML. Convert newlines to <br>.
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
     // Ensure elements exist before attaching listeners
     if (btnAddNewline) {
        btnAddNewline.addEventListener('click', () => insertAtCursor(textInput, '\n\n'));
     }
     if (btnAddPlaceholder) {
        btnAddPlaceholder.addEventListener('click', () => {
            if (!textInput) return;
            const start = textInput.selectionStart;
            insertAtCursor(textInput, '{}');
            textInput.selectionStart = textInput.selectionEnd = start + 1; // Place cursor inside {}
        });
     }

    if (snippetForm) {
        snippetForm.addEventListener('submit', handleFormSubmit);
    }

    if (snippetListContainer) {
        // Use event delegation for all actions within the list
        snippetListContainer.addEventListener('click', handleListActions);
    }

    if (cancelEditButton) {
        cancelEditButton.addEventListener('click', () => {
            clearForm();
             // Ensure form collapses if user cancels edit
             if (formSection && !formSection.classList.contains('hidden')) {
                 toggleFormDisplay(false); // Explicitly hide
             }
        });
    }

    // --- New Listeners ---
    if (toggleFormButton) {
        toggleFormButton.addEventListener('click', handleToggleForm);
    }

    if (searchInput) {
        searchInput.addEventListener('input', handleSearchInput); // 'input' event for real-time filtering
         // Optional: Clear search on Escape key
         searchInput.addEventListener('keydown', (event) => {
              if (event.key === 'Escape') {
                   searchInput.value = '';
                   renderSnippetList(currentSnippets); // Re-render full list
              }
         });
    }
}

// --- New Handler Functions ---
function handleToggleForm() {
    if (!formSection || !toggleFormButton) return;
    const isHidden = formSection.classList.toggle('hidden');
    toggleFormButton.classList.toggle('expanded', !isHidden);
     // Update button text/title if desired
     const icon = toggleFormButton.querySelector('.material-symbols-outlined');
     if (icon) {
          icon.textContent = isHidden ? 'expand_more' : 'expand_less';
     }
     // If showing form, clear any previous edit state only if not already editing
     if (!isHidden) {
          if (editOriginalCommandInput && editOriginalCommandInput.value === '') {
              clearForm(); // Clear only if not in edit mode
          }
     }
}
// Helper to explicitly show/hide form (used in edit/cancel/save)
function toggleFormDisplay(show) {
     if (!formSection || !toggleFormButton) return;
     const isHidden = formSection.classList.contains('hidden');
     if (show && isHidden) { // Show it if hidden
          formSection.classList.remove('hidden');
          toggleFormButton.classList.add('expanded');
          const icon = toggleFormButton.querySelector('.material-symbols-outlined');
          if (icon) icon.textContent = 'expand_less';
     } else if (!show && !isHidden) { // Hide it if shown
          formSection.classList.add('hidden');
          toggleFormButton.classList.remove('expanded');
          const icon = toggleFormButton.querySelector('.material-symbols-outlined');
           if (icon) icon.textContent = 'expand_more';
     }
}


function handleSearchInput() {
    // Re-render the list based on the current search term
    renderSnippetList(currentSnippets);
}


// --- Existing Handlers (modified for new structure/confirm) ---
function handleFormSubmit(event) {
     event.preventDefault();
     // Ensure elements exist
     if (!commandInput || !textInput || !editOriginalCommandInput) return;

    const command = commandInput.value.trim();
    const description = descriptionInput ? descriptionInput.value.trim() : '';
    const text = textInput.value;
    const originalCommand = editOriginalCommandInput.value;

    if (!command || !text) {
        showMessage(formMessage, "Command Name and Snippet Text are required.", true);
        return;
    }
    if (!/^[a-zA-Z0-9_\-]+$/.test(command)) {
         showMessage(formMessage, "Command Name can only contain letters, numbers, underscores, and hyphens.", true);
         return;
    }

    const snippetData = { command, description, text };
    let updatedSnippets = [...currentSnippets];

    if (originalCommand) {
        // Update
        const indexToUpdate = updatedSnippets.findIndex(s => s.command === originalCommand);
        if (indexToUpdate === -1) {
            showMessage(formMessage, `Error: Original snippet "${originalCommand}" not found for update.`, true);
            return;
        }
         if (command !== originalCommand && updatedSnippets.some((s, i) => s.command === command && i !== indexToUpdate)) {
             showMessage(formMessage, `Error: Command Name "${command}" already exists.`, true);
             return;
         }
        updatedSnippets[indexToUpdate] = snippetData;
        saveSnippetsToSettings(updatedSnippets, (success) => {
            if (success) {
                showMessage(formMessage, "Snippet updated successfully!", false);
                currentSnippets = updatedSnippets;
                clearForm();
                renderSnippetList(currentSnippets);
                toggleFormDisplay(false); // Hide form after successful save
            }
        });
    } else {
        // Add
         if (updatedSnippets.some(s => s.command === command)) {
             showMessage(formMessage, `Error: Command Name "${command}" already exists.`, true);
             return;
         }
        updatedSnippets.push(snippetData);
        saveSnippetsToSettings(updatedSnippets, (success) => {
             if (success) {
                 showMessage(formMessage, "Snippet added successfully!", false);
                 currentSnippets = updatedSnippets;
                 clearForm();
                 renderSnippetList(currentSnippets);
                 toggleFormDisplay(false); // Hide form after successful save
             }
         });
    }
}

function handleListActions(event) {
    const target = event.target;
    const button = target.closest('button'); // Find the actual button clicked
    const snippetItem = target.closest('.snippet-item');

    if (!button || !snippetItem) return; // Exit if click wasn't on a button within an item

    const snippetIndex = parseInt(snippetItem.getAttribute('data-index'), 10);
    const snippet = currentSnippets[snippetIndex];

    if (!snippet) {
        console.error("Could not find snippet data for index:", snippetIndex);
        showMessage(listMessage, "Error: Could not find snippet data.", true);
        return;
    }

    // Get references to the action and confirmation divs within this specific item
    const actionsDiv = snippetItem.querySelector('.snippet-actions');
    const confirmDiv = snippetItem.querySelector('.snippet-confirm-delete');
    const confirmTextSpan = confirmDiv ? confirmDiv.querySelector('.confirm-text') : null;

    // --- Hide all other confirmations first ---
    // Only do this if we are NOT clicking inside a confirmation itself
     if (!button.classList.contains('confirm-yes-btn') && !button.classList.contains('confirm-no-btn')) {
          document.querySelectorAll('.snippet-item').forEach(item => {
               if (item !== snippetItem) { // Don't hide the one we might be opening/interacting with
                    const otherConfirm = item.querySelector('.snippet-confirm-delete');
                    const otherActions = item.querySelector('.snippet-actions');
                    if (otherConfirm) otherConfirm.classList.add('hidden');
                    if (otherActions) otherActions.classList.remove('hidden'); // Ensure actions are visible
               }
          });
     }
     // --- ---

    // --- Handle Specific Button Clicks ---

    if (button.classList.contains('edit-btn')) {
        // Edit logic
        if (!formTitle || !commandInput || !descriptionInput || !textInput || !editOriginalCommandInput || !saveButton || !cancelEditButton) return; // Check elements needed

        formTitle.textContent = 'Edit Snippet';
        commandInput.value = snippet.command;
        descriptionInput.value = snippet.description;
        textInput.value = snippet.text;
        editOriginalCommandInput.value = snippet.command;
        saveButton.textContent = 'Update Snippet';
        cancelEditButton.classList.remove('hidden');
        if (formMessage) formMessage.classList.add('hidden');

        toggleFormDisplay(true); // Ensure form is visible for editing

        // Ensure correct UI state for the item being edited
        if (actionsDiv) actionsDiv.classList.remove('hidden');
        if (confirmDiv) confirmDiv.classList.add('hidden');

        window.scrollTo(0, 0); // Scroll task pane to top
        commandInput.focus();
        document.querySelectorAll('.placeholder-inputs').forEach(div => div.classList.add('hidden')); // Hide placeholder inputs too
    }
    else if (button.classList.contains('delete-btn')) {
        // Show Confirmation UI for *this* item
        if (actionsDiv) actionsDiv.classList.add('hidden');
        if (confirmDiv) {
            if (confirmTextSpan) confirmTextSpan.textContent = `Delete '${snippet.command}'?`; // Update text
            confirmDiv.classList.remove('hidden');
        }
    }
    else if (button.classList.contains('insert-btn')) {
        // Handle Insert Click
         handleInsertClick(event); // Pass the original event
         // Ensure correct UI state after potential placeholder interaction
         if (actionsDiv) actionsDiv.classList.remove('hidden');
         if (confirmDiv) confirmDiv.classList.add('hidden');
    }
    else if (button.classList.contains('confirm-yes-btn')) {
        // Execute Delete
        console.log("Confirmed delete for:", snippet.command);
        let updatedSnippets = currentSnippets.filter((s, i) => i !== snippetIndex);
        saveSnippetsToSettings(updatedSnippets, (success) => {
            if (success) {
                showMessage(listMessage, "Snippet deleted successfully!", false);
                currentSnippets = updatedSnippets; // Update cache *after* successful save
                renderSnippetList(currentSnippets); // Re-render the list
                 // If the deleted item was being edited, clear and hide the form
                 if (editOriginalCommandInput && editOriginalCommandInput.value === snippet.command) {
                     clearForm();
                     toggleFormDisplay(false);
                 }
            }
            // Hiding confirmation is now handled by re-rendering the list,
            // but we could explicitly hide here too if needed before save completes.
            // if (confirmDiv) confirmDiv.classList.add('hidden');
            // if (actionsDiv) actionsDiv.classList.remove('hidden');
        });
    }
    else if (button.classList.contains('confirm-no-btn')) {
        // Cancel Delete - Hide confirmation and show actions
        if (confirmDiv) confirmDiv.classList.add('hidden');
        if (actionsDiv) actionsDiv.classList.remove('hidden');
    }
}