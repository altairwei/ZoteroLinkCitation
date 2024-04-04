# ZoteroLinkCitation

An MS Word macro that links author-date or number style citations to their bibliography entry. This project was inspired by discussions of [Word: Possibility to link references and bibliography in a document?](https://forums.zotero.org/discussion/12431/word-possibility-to-link-references-and-bibliography-in-a-document)

## Supported Citation Styles

* http://www.zotero.org/styles/molecular-plant

Other styles are still being worked on.

## How to Use ZoteroLinkCitation

**Important Warning:** Before running the `ZoteroLinkCitation` macro, **please ensure you have backed up your document**. The operations performed by this script are bulk actions that are irreversible. A backup ensures that you can restore your original document in case anything does not go as expected.

This guide is aimed at beginners and provides detailed instructions on importing and running the `ZoteroLinkCitation.bas` script in Microsoft Word. This script, which includes the `ZoteroLinkCitation` macro along with other utility functions, enhances your document with advanced citation linking capabilities.

### Prerequisites

- Microsoft Word (2016 or later recommended for compatibility).
- The `ZoteroLinkCitation.bas` file.

### Step 1: Accessing the VBA Editor

1. Open Microsoft Word.
2. Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) Editor.

### Step 2: Importing the VB Script

1. Within the VBA Editor, locate `Normal` in the Project window on the left. Right-click on `Modules` under `Normal`. If `Modules` is not visible, right-click on `Normal` and choose `Insert` > `Module`.
2. With the new module selected, navigate to `File` > `Import File...` in the VBA Editor's menu.
3. Locate and select your `ZoteroLinkCitation.bas` file, then click `Open` to import the script.

### Step 3: Saving Your Macro-Enabled Document

1. Exit the VBA Editor to return to your Word document.
2. Save your document as a Macro-Enabled Document (`.docm`):
   - Click `File` > `Save As`.
   - Select your desired location.
   - Choose `Word Macro-Enabled Document (*.docm)` from the "Save as type" dropdown.
   - Click `Save`.

### Step 4: Adjusting Macro Security Settings

Adjust Word’s macro settings to allow the macro to run:

1. Go to `File` > `Options` > `Trust Center` > `Trust Center Settings...` > `Macro Settings`.
2. Select `Disable all macros with notification` for security while enabling functionality.
3. Click `OK` to confirm.

### Step 5: Running the `ZoteroLinkCitation` Macro

1. Make the `Developer` tab visible in Word (if it’s not already):
   - Navigate to `File` > `Options` > `Customize Ribbon`.
   - Ensure `Developer` is checked on the right side, then click `OK`.
2. Click `Macros` in the `Developer` tab.
3. Find and select `ZoteroLinkCitation` from the list, then click `Run`.

### Important Tips

- **Macro Security**: Only run macros from trusted sources. Macros can contain harmful code.
- **Testing**: Consider running the macro on a non-critical document first to familiarize yourself with its effects.

### Troubleshooting

- **Macro Not Running**: Verify the macro security settings and ensure the document is saved with a `.docm` extension.
- **Developer Tab Not Visible**: Enable the Developer tab through `Word Options` > `Customize Ribbon` > check `Developer`.
