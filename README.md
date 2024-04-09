# ZoteroLinkCitation

An MS Word macro that links author-date or numeric style citations to their bibliography entry. This project was inspired by discussions of [Word: Possibility to link references and bibliography in a document?](https://forums.zotero.org/discussion/12431/word-possibility-to-link-references-and-bibliography-in-a-document)

## Supported Citation Styles

### Author-Year styles

* [American Psychological Association (APA) 7th edition](http://www.zotero.org/styles/apa)
* [American Political Science Association](http://www.zotero.org/styles/american-political-science-association)
* [American Sociological Association 6th/7th edition](http://www.zotero.org/styles/american-sociological-association)
* [Molecular Plant](http://www.zotero.org/styles/molecular-plant)
* [China National Standard GB/T 7714-2015 (author-date)](http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-author-date)
* [Chicago Manual of Style 17th edition (author-date)](http://www.zotero.org/styles/chicago-author-date)

### Numeric styles

* [IEEE](http://www.zotero.org/styles/ieee)
* [Vancouver](http://www.zotero.org/styles/vancouver)
* [American Chemical Society](http://www.zotero.org/styles/american-chemical-society)
* [China National Standard GB/T 7714-2015 (numeric)](http://www.zotero.org/styles/china-national-standard-gb-t-7714-2015-numeric)
* [American Medical Association 11th edition](http://www.zotero.org/styles/american-medical-association)
* [Nature](http://www.zotero.org/styles/nature)

Other styles are still being worked on.

## How to Use ZoteroLinkCitation

**Important Warning:** Before running the `ZoteroLinkCitationAll` macro, **please ensure you have backed up your document**. The operations performed by this script are bulk actions that are irreversible. A backup ensures that you can restore your original document in case anything does not go as expected.

This guide is aimed at beginners and provides detailed instructions on importing and running the `ZoteroLinkCitation.bas` script in Microsoft Word.

### Prerequisites

- Microsoft Word (2016 or later recommended for compatibility).
- The `ZoteroLinkCitation.bas` file.

### Step 1: Accessing the VBA Editor

1. Open Microsoft Word.
2. Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) Editor.

### Step 2: Importing the VB Script

1. Within the VBA Editor, locate `Normal` in the Project window on the left. Right-click on `Normal` choose `Import File...`.
2. Locate and select your `ZoteroLinkCitation.bas` file, then click `Open` to import the script.

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

### Step 5: Running the `ZoteroLinkCitationAll` Macro

#### Method 1: Developer Tab

1. Make the `Developer` tab visible in Word (if it’s not already):
   - Navigate to `File` > `Options` > `Customize Ribbon`.
   - Ensure `Developer` is checked on the right side, then click `OK`.
2. Click `Macros` in the `Developer` tab.
3. Find and select `ZoteroLinkCitationAll` from the list, then click `Run`.

#### Method 2: Shortcut Key

Press `Alt` + `F8`, find and select `ZoteroLinkCitationAll` from the list, then click `Run`.

#### Method 3: Add a Button

1. Click `File` > `Options` > `Quick Access Toolbar`.
2. In the `Choose commands from list`, click `Macros`.
3. Select the macro `ZoteroLinkCitationAll`.
4. Click `Add` to move the macro to the list of buttons on the `Quick Access Toolbar`.

### Important Tips

- **Macro Security**: Only run macros from trusted sources. Macros can contain harmful code.
- **Testing**: Consider running the macro on a non-critical document first to familiarize yourself with its effects.

### Troubleshooting

- **Macro Not Running**: Verify the macro security settings and ensure the document is saved with a `.docm` extension.
