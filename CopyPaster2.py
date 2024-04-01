from pynput.mouse import Listener as MouseListener
import pyautogui
import pyperclip
import pytesseract
from PIL import Image, ImageEnhance
import re
from tkinter import Tk, Label
import threading
from openpyxl import load_workbook

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

workbook = load_workbook('data.xlsx')
sheet = workbook.active

copy_list = []

# Check for duplicate copy_id values
duplicate_check = set()

for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row - 1, values_only=True):
    copy_id = row[6]  # 7th column, considering index starts at 0
    copy_value = row[5]  # 6th column, considering index starts at 0
    
    # Check if the copy_id already exists
    if copy_id in duplicate_check:
        print(f'Duplicate copy_id found: {copy_id}')
    else:
        duplicate_check.add(copy_id)
    
    copy_list.append({"copy_id": copy_id, "copy_value": copy_value})

def update_gui_with_copy_id(copy_id):
    """
    Update the GUI with the provided copy_id.
    
    The copy_id is displayed on the GUI for a short duration,
    and the background color of the label changes to indicate
    different states.
    """
    
    # Set the window title to "Copied ID"
    window.title("Copied ID")
    
    # Set the text of the label to the provided copy_id
    label.config(text=f"{copy_id}")
    
    # Update the GUI
    window.update_idletasks()
    window.update()
    
    # Change the background color of the label to green after 200ms
    window.after(200, lambda: label.config(bg="green"))
    
    # Change the background color of the label to blue after 400ms
    window.after(400, lambda: label.config(bg="blue"))
    
    # Change the background color of the label back to green after 600ms
    window.after(600, lambda: label.config(bg="green"))
    
    # Change the background color of the label to blue after 800ms
    window.after(800, lambda: label.config(bg="blue"))
    
    # Change the background color of the label back to the system default after 1000ms
    window.after(1000, lambda: label.config(bg="SystemButtonFace"))



def keep_window_always_on_top():
    """
    Set the window attribute to keep the window always on top.
    """
    window.attributes('-topmost', True)

def start_gui():
    # Starts the GUI thread
    global window, label
    window = Tk()
    window.title("Copied ID")
    window_width = 800
    window_height = 115
    screen_width = window.winfo_screenwidth()
    x_coordinate = (screen_width / 2) - (window_width / 2)
    y_coordinate = 0 
    window.geometry(
        f"{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}")
    label = Label(window, font=("Arial", 48, "bold"), foreground="red",
                  text="No ID copied yet")
    label.pack(expand=True)
    window.attributes('-topmost', True)
    window.mainloop()

gui_thread = threading.Thread(target=start_gui, daemon=True)
gui_thread.start()

def read_and_copy(x, y, button, pressed):
    """
    Takes a screenshot on left click, reads text, and copies the corresponding value.

    Args:
        x (int): The x-coordinate of the mouse cursor.
        y (int): The y-coordinate of the mouse cursor.
        button (Button): The button that was pressed.
        pressed (bool): True if the button was pressed, False otherwise.

    Returns:
        None
    """
    # Takes a screenshot on left click, reads text, and copies the corresponding value
    if pressed:
        keep_window_always_on_top()

    if button.name == "left" and pressed:
        # This is very critical.
        # Set the coordinates to capture a specific region of the screen.
        # (top left corner x, top left corner y, width, height)
        region_to_capture = (460, 110, 490, 120)
        
        # Capture the screen screenshot within the specified region.
        screenshot = pyautogui.screenshot(region=region_to_capture)

        screenshot.save("screenshot.png")

        image = Image.open("screenshot.png")

        gray_image = image.convert('L')

        enhancer = ImageEnhance.Contrast(gray_image)
        enhanced_image = enhancer.enhance(2)  # contrast value

        text = pytesseract.image_to_string(enhanced_image)
        text = re.sub(r'[.]', '-', text)

        print("OCR Text:", text)
        
        for item in copy_list:
            pattern = re.escape(item["copy_id"]) + r"(?![ -])"
            if re.search(pattern, text, re.IGNORECASE):
                pyperclip.copy(item["copy_value"])
                print(f'{"copied value: " + item["copy_value"]}')
                update_gui_with_copy_id(item["copy_id"])
                break

with MouseListener(on_click=read_and_copy) as listener:
    listener.join()
