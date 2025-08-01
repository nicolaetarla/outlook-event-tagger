# Outlook Event Tagger Plugin

This Outlook add-in automatically tags newly created calendar events with a user-defined or pre-defined category.

## Functionality

- **Mandatory Category Selection**: Upon first use or if no category is configured, the add-in will prompt the user to select an Outlook category before an event can be sent or saved.
- **Automatic Tagging**: Once configured, the add-in automatically applies the selected Outlook category to newly created, saved, or sent calendar events.
- **User-Defined Tag**: Users select an existing Outlook category from a list within the add-in's task pane. This selected category will then be used for all subsequent new events.

## How it Works

The plugin leverages Outlook's Event-based Activation feature, specifically the `ItemSend`, `ItemSave`, and `ItemNew` events. When a user creates, saves, or sends a new calendar event, the `onCalendarEvent` function in `src/commands/commands.js` is triggered. This function retrieves the user's preferred category (from roaming settings) and applies it to the event. If no category is configured, the event will not be sent/saved, and the user will be prompted to configure the add-in.

## Project Structure

- `manifest.xml`: The core manifest file that defines the add-in's identity, permissions, and entry points. It includes references to the task pane and command files.
- `src/taskpane/taskpane.html`: The HTML file for the add-in's task pane, providing the user interface for category selection.
- `src/taskpane/taskpane.js`: The JavaScript logic for the task pane, handling the loading of existing Outlook categories, saving the user's selection to roaming settings, and displaying messages.
- `src/commands/commands.html`: A minimal HTML file required for the event-based activation runtime.
- `src/commands/commands.js`: Contains the `onMessageSend` function, which is the entry point for the `ItemSend` event. This script reads the selected category (or uses the default) and applies it to the calendar item.
- `.gitignore`: Specifies files and directories that Git should ignore.

## Setup and Installation (Sideloading)

To use and test this add-in, you need to sideload it into your Outlook client. This typically involves serving the add-in files from a local web server and then registering the `manifest.xml` with Outlook.

1.  **Generate a Unique GUID**: Ensure the `<Id>` tag in `manifest.xml` contains a unique GUID. This has already been done for you.

2.  **Serve Files Locally (HTTPS Required)**:
    Outlook add-ins require files to be served over HTTPS. You can use a tool like `http-server` or `live-server` for this. If you have the Office Add-in Debugging Tools installed, `npx office-addin-debugging start` can handle this for you.

    Example using `http-server`:
    ```bash
    # Install http-server globally if you haven't already
    npm install -g http-server

    # Navigate to the root of this project
    cd /Users/ntarla/CODE/outlook-event-tagger

    # Generate a self-signed certificate (if you don't have one)
    # For development, you can use mkcert or openssl
    # Example using openssl (for Linux/macOS):
    # openssl req -x509 -newkey rsa:2048 -nodes -keyout key.pem -out cert.pem -days 365

    # Start the server (replace with your cert/key paths if different)
    http-server . -S -C cert.pem -K key.pem -p 3000
    ```
    Ensure the `AppDomain` and `SourceLocation` URLs in `manifest.xml` match your server's address (e.g., `https://localhost:3000`).

3.  **Sideload the Add-in in Outlook**:
    *   **Outlook on the Web**: Go to `Settings` (gear icon) > `View all Outlook settings` > `Mail` > `Customize actions` (or `General` > `Manage add-ins`). Click `Add new add-in` and choose `Add from file`. Select the `manifest.xml` file.
    *   **Outlook Desktop Client**: Go to `File` > `Manage Add-ins` (or `Get Add-ins`). This will usually open a web page. From there, look for an option to `Add a custom add-in` or `My add-ins` and then `Add from a file...`. Select the `manifest.xml` file.

## Usage

1.  **Configure the Tag**: Create a new calendar event. You should see a new button (e.g., "Configure Event Tagger") in the ribbon. Click it to open the task pane. Select your desired category from the dropdown and click "Save Setting."
2.  **Create New Events**: Create a new calendar event. When you send or save it, the add-in will automatically apply the configured category to the event.

## Customization

- **Icons**: Update the `IconUrl` and `HighResolutionIconUrl` in `manifest.xml` to your preferred icon URLs.