from pynput.mouse import Listener as MouseListener
import pyautogui
import pyperclip
import pytesseract
from PIL import Image, ImageEnhance
import re
from tkinter import Tk, Label
import threading

copy_list = [
    {"copy_id": "ARBR-1", "copy_value": "111"},
    {"copy_id": "ARBR-2", "copy_value": "222"},
    {"copy_id": "FK-3", "copy_value": "333"},
    {"copy_id": "FK-4", "copy_value": "444"}
]

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def update_gui_with_copy_id(copy_id):
    window.title("Copied ID")
    label.config(text=f"Copied ID: {copy_id}")
    window.update_idletasks()
    window.update()

def keep_window_always_on_top():
    window.attributes('-topmost', True)

def start_gui():
    #Starts the GUI thread
    global window, label
    window = Tk()
    window.title("Copied ID")
    window_width = 1100
    window_height = 120
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

def regexify_copy_id(copy_id):
#Converts Turkish characters in copy_id to dots and returns it
    copy_id = re.sub(r'[çğıöşüÇĞİÖŞÜ]', '.', copy_id)
    print(copy_id)
    return copy_id

for item in copy_list:
    item["copy_id"] = regexify_copy_id(item["copy_id"])

def read_and_copy(x, y, button, pressed):
    #Takes a screenshot on left click, reads text, and copies the corresponding value
    if pressed:
        keep_window_always_on_top()

    if button.name == "left" and pressed:
        #This is very critical.
        # Set the coordinates to capture a specific region of the screen.
        # (top left corner x, top left corner y, width, height)
        region_to_capture = (460, 110, 490, 120)
        
        # Capture the screen screenshot within the specified region.
        screenshot = pyautogui.screenshot(region=region_to_capture)

        # You can use this code to capture the full screen if necessary.
        #screen_width, screen_height = pyautogui.size()
        #screenshot = pyautogui.screenshot(region=(0, 0, screen_width, screen_height))
        
        screenshot.save("screenshot.png")

        image = Image.open("screenshot.png")

        gray_image = image.convert('L')

        enhancer = ImageEnhance.Contrast(gray_image)
        enhanced_image = enhancer.enhance(2)  # contrast value

        text = pytesseract.image_to_string(enhanced_image)

        print("OCR Metni:", text)
        
        for item in copy_list:
            regex_pattern = re.escape(item["copy_id"]) + r'\b(?!\w)'

            if re.search(regex_pattern, text, re.IGNORECASE):
                pyperclip.copy(item["copy_value"])
                print(f'Kopyalandı: {item["copy_value"]}')
                update_gui_with_copy_id(item["copy_id"])
                break

with MouseListener(on_click=read_and_copy) as listener:
    listener.join()
