# ZoteroLinkCitation

       An MS Word macro that links Zotero author-date or numeric style citations to their bibliography entry. This program has been **improved and expanded** based on https://github.com/altairwei/ZoteroLinkCitation. **Specifically, we have developed corresponding functional tabs on the Word toolbar to enable quick and permanent access**. 
      This project was inspired by discussions of [Word: Possibility to link references and bibliography in a document?](https://forums.zotero.org/discussion/12431/word-possibility-to-link-references-and-bibliography-in-a-document)

## Functionalities:

* Supports various author-year and numeric citation styles, including APA, Chicago, IEEE, Nature, Vancouver, Cite Them Right.
* Retains the Zotero fields after linking.
* Enables the formatting of links (e.g. colour, font).
* Also works with multiple references within a citation field.
* Word toolbar quick and permanent access.
## Support style
1.olecular-plant    2.ieee    3.apa    4.vancouver    5.american-chemical-society    6.erican-medical-association    7.nature    8.american-political-science-association    9.rican-sociological-association    10.chicago-author-date    11.bmc-medicine    12.china-national-standard-gb-t-7714-2015-numeric    13.china-national-standard-gb-t-7714-2015-author-date    14.harvard-cite-them-right    15.elsevier-harvard    16.modern-language-association

## How to Use

> [!CAUTION]
> Before running the `ZoteroLinkCitationAll` macro, ** suggest backed up your document to avoid irreversible mishaps**.
This guide is aimed at beginners and provides detailed instructions on importing the `ZoteroLinkCitation.bas` script and creating the corresponding Word toolbar in Microsoft Word.

### Prerequisites

- Microsoft Word (2016 or later recommended for compatibility).
- Download the `ZoteroLinkCitation.bas` file.

### Step 1: Accessing the VBA Editor

1. Open Microsoft Word.
2. Press `Alt` + `F11` (or `Alt`+`Fn`+ `F11`) to open the Visual Basic for Applications (VBA) Editor.

### Step 2: Importing the `ZoteroLinkCitation.bas` Script

1. Within the VBA Editor, locate `Normal` in the Project window on the left. Right-click on `Normal` choose `Import File...`.
2. Locate and select your `ZoteroLinkCitation.bas` file, then click `Open` to import the script.

### Step 3: Creating the subtoolbar in **Cite** toolbar

1. Exit the VBA Editor to return to your Word document.
2. Open the Word options
   - Right-click on `Cite` choose `Custom Ribbon`
3. Create `ZoteroCitationHyperlink` under the `Cite`
   - Locate `Cite`: Click `Create a new group` > `Rename`. Rename as ZoteroCitationHyperlink.
4. Add four subtoolbars in `ZoteroCitationHyperlink` group
   - Locate `Select a command from the following locations`: Click `Common Commands` choose `Macro`.
   a) **Remove Hyperlink** toolbar.
   - Choose `Normal.ZoteroLinkCitation.**ZoteroUnlinkCitations**`: click `Add` > `Rename`.
   - Write `Remove Hyperlink` and set symbol as **1-row by 2-column (or 2nd) symbol** in Symbol group.
   - Click `Enter`.
   b) **Auto Hyperlink** toolbar
   - Choose `Normal.ZoteroLinkCitation.**ZoteroLinkCitationAll**`: click `Add` > `Rename`
   - Write `Auto Hyperlink` and set symbol as **10-row by 9-column (or 90-th) symbol** in Symbol group.
   - Click `Enter`
  c) **Manual Hyperlink** toolbar
   - Choose `Normal.ZoteroLinkCitation.**ZoteroLinkCitationSelect**`: click `Add` > `Rename`
   - Write `Manual Hyperlink` and set symbol as **4-row by 5-column (or 20-th) symbol** in Symbol group.
   - Click `Enter`
  d) **Support Style Information** toolbar
   - Choose `Normal.ZoteroLinkCitation.**SupportStylelnformation**`: click `Add` > `Rename`
   - Write `Support Style Information` and set symbol as **11-row by 5-column (or 20-th) symbol** in Symbol group.
   - Click `Enter`
Finally, click `Enter` of Word options. You can see four subtoolbars in **Cite** toolbar.

### Step 4: Adjusting Macro Security Settings (if need)

Adjust Word’s macro settings to allow the macro to run:

1. Go to `File` > `Options` > `Trust Center` > `Trust Center Settings...` > `Macro Settings`.
2. Select `Disable all macros with notification` for security while enabling functionality.
3. Click `OK` to confirm.

### Step 5: Running the `ZoteroLinkCitationAll` Macro

Click new toolbars in **Cite** toolbar to achieve the desired behavior.

### Important Tips

- **Macro Security**: Only run macros from trusted sources. Macros can contain harmful code.
- **Testing**: Consider running the macro on a non-critical document first to familiarize yourself with its effects.

### Step 6: Deleting added toolbars

Open Word options > find the `ZoteroCitationHyperlink` group > click `delete`.

### Step 7: Changing citation color (generally, it remains black by default.)

1. Press `Alt` + `F11` (or `Alt`+`Fn`+ `F11`) to open the Visual Basic for Applications (VBA) Editor.
2. Within the VBA Editor, locate `Normal` in the Project window on the left. Double-click on `ZoteroLinkCitation`.
3. Press `Ctrl` + `F` searchs **Sub SetZoteroCitationColor** in right code.
4. Set desired color after `fld.result.Font.ColorIndex = ` code.  **wdAuto→black,wdRed→red、wdGreen→green、wdBlue→blue**
5. Save document.
## Known Issues

If you encounter any errors, please do not hesitate to contact me or refer to **Known Issues** in https://github.com/altairwei/ZoteroLinkCitation.
