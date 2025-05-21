# Snippet Manager (Outlook Add-in)

Manage and insert text snippets into your emails seamlessly within Microsoft Outlook. This add-in allows you to create, store, and quickly use predefined blocks of text, including support for dynamic placeholders.

## How to Use

1.  **Open the Snippet Manager:**
    *   In Outlook, when composing a new email or replying to one, look for the "Snippet Manager" button in the Outlook ribbon (usually under the "Message" tab or a dedicated "Snippets" group).
    *   Click the "Snippets" button to open the task pane.

2.  **Create a New Snippet:**
    *   At the top of the task pane, click the "Create New Snippet" (or similar, often an icon like `+` or `expand_more`) button to expand the form.
    *   **Command Name:** Enter a short, memorable command for your snippet (e.g., `greet_hello`, `signature_work`). This is used to identify the snippet.
    *   **Description (Optional):** Add a brief description of what the snippet is for.
    *   **Snippet Text:** Type or paste the text content of your snippet.
        *   **Placeholders:** To create a dynamic field, insert a placeholder like `{placeholder_name}` (e.g., `{client_name}`, `{date}`). When you insert the snippet, you'll be prompted to fill these in. Use the "Add Placeholder {}" button for convenience.
        *   **Newlines:** Use the "Add Newline" button or press Enter to create multi-line snippets.
    *   Click "Save Snippet".

3.  **Insert a Snippet:**
    *   Find the snippet in the list. You can use the search bar to filter snippets by command, description, or text.
    *   Click the "Insert" icon (often looks like `data_object` or similar) next to the snippet.
    *   If the snippet contains placeholders, a small form will appear below it. Fill in the values for the placeholders and click "Confirm & Insert".
    *   The snippet text (with placeholders filled) will be inserted into the currently active part of your email.

4.  **Edit a Snippet:**
    *   Find the snippet in the list.
    *   Click the "Edit" icon (pencil icon).
    *   The snippet's details will load into the form at the top. Modify as needed.
    *   Click "Update Snippet".

5.  **Delete a Snippet:**
    *   Find the snippet in the list.
    *   Click the "Delete" icon (trash can icon).
    *   A confirmation prompt will appear. Click "Yes" (or a checkmark icon) to confirm deletion.

6.  **Search for Snippets:**
    *   Use the search bar at the top of the snippet list to filter snippets by their command name, description, or even parts of their text content. The list updates as you type.

## Features

*   **Snippet Management:** Easily create, edit, and delete text snippets.
*   **Placeholder Support:** Insert dynamic fields (e.g., `{client_name}`) into your snippets that can be filled out before insertion.
*   **Rich Text Editing:** Add newlines and placeholders with dedicated buttons.
*   **Quick Insertion:** Insert snippets into your email body with a single click.
*   **Search Functionality:** Quickly find snippets by command, description, or content.
*   **Roaming Storage:** Snippets are saved to your Outlook profile's roaming settings, making them available across your Outlook instances (where supported).
*   **Default Snippets:** Comes pre-loaded with a set of example snippets from `assets/snippets.json` on first use or if storage is empty.
*   **User-Friendly Interface:** A clear and intuitive task pane for managing snippets.

## Technical Overview

*   **Type:** Microsoft Outlook Mail Add-in.
*   **Manifest:** `outlooksnippets/manifest.xml` (defines the add-in's properties and UI integration).
*   **Task Pane UI:**
    *   `outlooksnippets/taskpane/taskpane.html` (HTML structure)
    *   `outlooksnippets/taskpane/taskpane.css` (Styling)
    *   `outlooksnippets/taskpane/taskpane.js` (Core logic, event handling, communication with Outlook)
*   **Default Snippets Data:** `outlooksnippets/assets/snippets.json`
*   **Icons:** Located in `outlooksnippets/assets/`.

## Installation/Deployment

This add-in is designed to be hosted on a web server (currently configured for GitHub Pages as per the manifest).

**For End Users (Production):**
The add-in would typically be deployed by an administrator through the Microsoft 365 admin center or acquired from the Office Store if published.

**For Developers (Sideloading for Testing):**
You can sideload this add-in in Outlook for testing and development:

1.  **Using Outlook on the web:**
    *   Open Outlook on the web.
    *   Open an email or start composing a new one.
    *   Click the "Get Add-ins" button.
    *   In the "Add-Ins for Outlook" dialog, select "My add-ins".
    *   Scroll down to "Custom Addins" and click "+ Add a custom add-in".
    *   Select "Add from file..."
    *   Upload the `outlooksnippets/manifest.xml` file from this repository.
    *   Follow the prompts to install.

2.  **Using Outlook Desktop (Windows/Mac):**
    *   The process is similar. Find the "Get Add-ins" option in the Outlook ribbon.
    *   Look for "My add-ins" and the option to add a custom add-in from a file, then upload the `outlooksnippets/manifest.xml`.

**Note:** The URLs within the `manifest.xml` (e.g., `SourceLocation`, `IconUrl`) point to `https://crood.github.io/outlooksnippets/...`. If you are hosting it yourself or modifying it, these URLs will need to be updated in the `manifest.xml` file to point to your server.

## Contributing

Contributions are welcome! If you have ideas for improvements, new features, or find bugs, please feel free to:

1.  Fork the repository.
2.  Create a new branch for your feature or bug fix.
3.  Make your changes.
4.  Submit a pull request.

## License

This project is licensed under the terms of the MIT License. See the [LICENSE](LICENSE) file for details.
