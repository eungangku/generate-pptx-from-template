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
    sleep(1)
    pyautogui.hotkey("ctrl", "c")

    # close the file
    sleep(2)
    pyautogui.hotkey("alt", "f4")

    # open the template file
    sleep(3)
    os.startfile("template.pptx")
    sleep(7)
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

    # close the new file
    sleep(10)
    pyautogui.hotkey("alt", "f4")
