# Snippet Tool for Visual Studio 2022 Change Log 📋

## v1.5 *(current)* 🆕
#### 🛠️ Fixes:
    • Same accelerator key was assigned for two different commands.
#### 🌟 Improvements:
    • Source-code upgraded to .NET Framework 4.8.
    • Implemented standard asynchronous package loading.
    • Command names were renamed with better descriptions.
    • Some commands now can work properly by distinguishing if the selected text is single-line or multi-line.

## v1.4 🔄
#### 🌟 Improvements:
    • Extension upgraded to support Visual Studio 2022 community edition or higher.

## v1.3 🔄
#### 🚀 New Features:
    • New Command: Collapse Xml Comments
    • New Command: Expand Xml Comments
    • New Command: Delete Xml Comments
#### 🛠️ Fixes:
    • In previous version some commands needs to be undone twice. 
          ( Fixed by the usage of 'UndoContext' DTE's member. )
#### 🌟 Improvements:
    • Paragraph command now adds an empty <para></para> section if any text is selected.
    • Remarks command now adds an empty <remarks></remarks> section if any text is selected.
    • General source-code refactor.

## v1.2 🔄
#### 🌟 Improvements:
    • "Code Ref.", "Param Ref." and "Lang. Ref." commands works on the word at where the caret is, 
      this means, no need to select the text anymore when pretending to document only one word.
    • "Code Ref.", "Param Ref." and "Lang. Ref." commands are enabled by default.

## v1.1 🔄
#### 🚀 New Features:
    • A properties page with name "Snippet Tools" inside the "Tools -> Options" menu.
    • Paragraph command ( <para></para> tag ) with hotkey: Ctrl+E+Space.
    • Separator Line command with hotkey: Ctrl+E+Tab.
#### 🛠️ Fixes:
    • Keyboard shortcuts now are only avaliable when are pressed on the text editor window.
#### 🌟 Improvements:
    • Simplified Command icons.
    • Tag enclosing behavior when a full line is selected.
    • Improved keyboard shortcuts of "Hyperlink", "Hyperlink Alter" and "Remarks Section" commands.

## v1.0 🔄
Initial Release.