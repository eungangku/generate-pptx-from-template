import os
from pathlib import Path
import pyautogui
from time import sleep


# Loop over all files in the target directory
def get_file_names(directory):
    return os.listdir(directory)


sleep(5)
files = get_file_names("source_pptx")
for file in files:
    path = f"source_pptx/{file}"
    os.startfile(Path(path))
    sleep(7)

    # select slide section
    coord = pyautogui.locateCenterOnScreen("1.png", confidence=0.9)
    pyautogui.click(coord)

    # copy all the slides
    sleep(2)
    pyautogui.hotkey("ctrl", "a")
    sleep(2)
    pyautogui.hotkey("ctrl", "c")

    # close the file
    sleep(2)
    pyautogui.hotkey("alt", "f4")

    # open the template file
    sleep(3)
    os.startfile("template.pptx")
    sleep(10)
    coord = pyautogui.locateCenterOnScreen("changeview.png", confidence=0.9)
    pyautogui.click(coord)
    sleep(2)

    # paste the slides into the template
    sleep(3)
    coord = pyautogui.locateCenterOnScreen("index.png", confidence=0.9)
    # right click
    pyautogui.click(coord, button="right")
    sleep(2)
    # paste
    coord = pyautogui.locateCenterOnScreen("paste.png", confidence=0.9)
    pyautogui.click(coord)

    # save the slides
    sleep(2)
    pyautogui.hotkey("ctrl", "shift", "s")
    
    sleep(2)
    pyautogui.write(file.replace(".pptx", ""))
    pyautogui.press("enter")
    
    sleep(2)
    try:
        # Attempt to locate the 'confirm.png' image on the screen
        confirm_button = pyautogui.locateCenterOnScreen('confirm.png', confidence=0.9)

        if confirm_button:
            # If the 'confirm.png' was found, pyautogui will click on the center of this image
            pyautogui.click(confirm_button)
            print("Clicked on the 'Confirm' button.")
        else:
            # If the 'confirm.png' was not found, inform the user
            print("Could not find the 'Confirm' button on the screen.")
    except pyautogui.ImageNotFoundException:
        # Exception handling if the image is not found at all
        print("Image not found. Please ensure 'confirm.png' exists in the current directory.")
    except Exception as e:
        # Handle other potential exceptions
        print(f"An unexpected error occurred: {e}")

    # close the new file
    sleep(12)
    pyautogui.hotkey("alt", "f4")
