# CopyPaster2

Screen Text Copier
Overview
The Screen Text Copier is a Python-based automation tool designed to enhance productivity by streamlining the process of copying text from screen areas where direct selection is not feasible. Utilizing Optical Character Recognition (OCR) technology, this script intelligently captures, recognizes, and matches text within a predefined screen region against a list of identifiers. Upon finding a match, it copies the corresponding value to the clipboard, ready for use. This tool is invaluable for tasks involving repetitive data entry, document processing, or interaction with software that restricts text selection, effectively reducing manual effort and improving workflow efficiency.

Features
Automatic Text Recognition: Leverages OCR to identify text within a user-defined screen area.
Intelligent Text Matching: Uses regular expressions to match recognized text against a predefined list of identifiers, accommodating variations and special characters.
Clipboard Integration: Automatically copies the matched text's corresponding value to the clipboard.
Real-time Feedback: Updates a minimalistic GUI with the copied identifier, ensuring the user is informed of the action taken.
Customizable Identifiers: Allows for easy customization of identifiers and their corresponding values within the script.
Dependencies
Python: The script is written in Python, requiring Python 3.x.
pytesseract: A Python wrapper for Google's Tesseract-OCR Engine, used for OCR capabilities.
PyAutoGUI: Utilized for capturing screenshots of the screen region of interest.
PIL (Python Imaging Library): Required for image processing tasks before text recognition.
pyperclip: Allows for copying text to the system's clipboard.
Tkinter: Provides the graphical user interface for real-time feedback.
threading: Used to run the GUI in a background thread, keeping the application responsive.
Installation
Python Installation: Ensure Python 3.x is installed on your system.
Dependency Installation: Install the required libraries using pip:
Copy code
pip install pytesseract pyautogui Pillow pyperclip
Tesseract-OCR Setup: Install Tesseract-OCR and set the pytesseract.pytesseract.tesseract_cmd to the installation path.
Running the Script: Execute the script in your Python environment. Ensure the screen region coordinates in the script match the area you intend to capture.
Usage
Configuration: Adjust the copy_list within the script to include the identifiers and corresponding values you wish to copy.
Activation: The script listens for left mouse clicks. Upon clicking, it captures the predefined screen area, processes the image, and attempts to match and copy the recognized text.
GUI Feedback: The GUI updates in real-time to display the copied identifier, providing immediate feedback without disrupting workflow.
Contributing
Contributions are welcome! Whether it's adding new features, improving existing ones, or reporting issues, your input helps make this tool better for everyone.
