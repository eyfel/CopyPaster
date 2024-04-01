# CopyPaster

This Python script captures a screenshot from a specific region, reads the text in the image, and identifies a specific text pattern. It then matches rows from two specified hours in an Excel file based on this identified text. After reading text from a designated window using OCR, it retrieves the corresponding data from the matched row in the Excel file and adds it to the clipboard. Finally, it displays the copied value in a graphical user interface (GUI).

This advanced feature enhances functionality by automating data retrieval based on detected text, thus improving efficiency and usability.It enables completing tasks 10 times faster.

## Usage

- The script runs when it detects a left mouse click.
- Firstly, it captures a screenshot from a specific region.
- It processes the image, reads the text, and finds the identification number based on a specific pattern.
- It copies the value corresponding to the identification number to the clipboard.
- It displays the copied identification number in the graphical user interface (GUI).

## Requirements

This script requires the following Python libraries to be installed:
- openpyxl
- pyautogui
- pyperclip
- pynput
- pytesseract (You should download this from the internet.)

You can install these libraries by running the following command in your terminal or command prompt:

```
pip install -r requirements.txt
```

## Notes

- The script uses the `pyautogui` library to capture a screenshot from a specific region.
- It utilizes the pytesseract OCR engine to read the text. 
- It reads an Excel file containing the identification number and the corresponding value.
- It uses `tkinter` to display the copied identification number in the graphical user interface.

## License

This script is licensed under the MIT License. For more information, please see the [LICENSE](LICENSE) file.
