# Word Add-in: DocUtility

The Word Add-in "DocUtility" is a Microsoft Word specific add-in that allows you to manipulate Word contents in specific ways. This add-in is built using C# and VSTO (Visual Studio Tools for Office) technology.

## Features

- Highlights the word "of"
- Changes the same word to uppercase on odd clicks and underlines the preceding word on even clicks
- Reverses words in a paragraph from where the caret cursor will be
- Reverses the order of the entire document's paragraphs
- Automatically identifies even-numbered paragraphs.
- Extracts the middle word from the previous paragraph.
- Replaces the middle word of the current paragraph with the extracted word.

## Requirements

- Microsoft Word 2010 or later.
- .NET Framework 4.5 or later.

## Installation

1. Clone or download the repository to your local machine.
2. Open the solution file (`BNDocument.sln`) in Visual Studio.
3. Build the solution to generate the add-in assembly (`BNDocument.dll`).
4. Copy the `BNDocument.dll` to the desired location on your computer.
5. Open Microsoft Word.
6. Click on **File -> Options -> Add-ins**.
7. In the Add-ins window, click on the **Manage** dropdown and select **Word Add-ins**. Click **Go**.
8. In the Add-ins dialog, click on **Browse** and locate the `BNDocument.dll` file. Click **OK** to add the add-in.
9. The "DocUtility" add-in will appear in the list. Make sure the checkbox is selected to enable the add-in.
10. Click **OK** to close the Add-ins dialog.
11. For computer installation, extract the setup file from the DocUtilty.zip.
12. Run it and install, then open MS Word and go to the **Test** tab.

## Usage

1. Open a Word document.
2. Click on the **Test** tab in the Word ribbon.
3. Click on the **First** button to highlight the word of and see its number of occurrences
4. Click on the **Second** button to change the word's case and underline the preceding word
5. Click on the **Third** button/combobox to select "Paragraph" or "Document" which reverses either respectively
6. Click on the **Fourth** button to perform the middle word replacement.
7. The middle word of each even-numbered paragraph will be replaced with the middle word from the previous paragraph.

## Contributing

Contributions are welcome! If you encounter any issues or have suggestions for improvements, please feel free to submit a pull request or open an issue in the GitHub repository.

## Disclaimer

The "DocUtility" add-in is provided as-is without any warranties or guarantees and is for experimental or educational purposes. Use it at your own risk.
