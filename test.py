
import os
import winshell
from win32com.client import Dispatch
from tkinter import *
from tkinter import filedialog as fd

from PIL import Image, ImageTk
import win32ui
import win32gui
import win32con
import win32api
import random
# Create a Class that stores two paths named image and shortcut


class ImagePath:
    def __init__(self, image, shortcut):
        self.image = image
        self.shortcut = shortcut

# Create an check to see if Shortcuts.xml exists


def check_shortcuts():
    try:
        f = open("Shortcuts.xml", "r")
        f.close()
        return True
    except:
        # Create a new Shortcuts.xml file intialized with the default values
        f = open("Shortcuts.xml", "w")
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write("<Shortcuts>\n")
        f.write("</Shortcuts>\n")

        return False


check_shortcuts()
# Create a prototype of the ImagePath class and intialize it with default xml values


def create_image_path(path="", shortcut=""):
    image_path = ImagePath(path, shortcut)
    return image_path

# Convert ImagePath class to xml format


def image_path_to_xml(image_path):
    xml = "<ImagePath>\n"
    xml += "\t<img>" + image_path.image + "</img>\n"
    xml += "\t<Shortcut>" + image_path.shortcut + "</Shortcut>\n"
    xml += "</ImagePath>\n"
    return xml


# Create a function that inserts xml inbetween <Shortcuts> and </Shortcuts> tags
def insert_xml(xml):
    f = open("Shortcuts.xml", "r")
    lines = f.readlines()
    f.close()
    lines.insert(2, xml)
    f = open("Shortcuts.xml", "w")
    f.writelines(lines)
    f.close()


# Parse the xml file and return a list of <ImagePath> objects from between the <ImagePath> and </ImagePath> tags

def get_image_paths(): 
    image_paths = []
    f = open("Shortcuts.xml", "r")
    lines = f.readlines()
    f.close()
    for line in lines:
        if "<img>" in line:
            image_path = create_image_path()
            image_path.image = line.split("<img>")[1].split("</img>")[0]
            print(image_path.image)
        elif "<Shortcut>" in line:
            image_path.shortcut = line.split("<Shortcut>")[1].split("</Shortcut>")[0]
            image_paths.append(image_path)
    return image_paths

# Print the list of ImagePaths to the console


def print_image_paths(image_paths):
    image_paths = get_image_paths()
    for image_path in image_paths:
        print(image_path.image, image_path.shortcut)

# Create a gallery of the images using the <img> tags in the xml file
# Create a function that deletes all the Shortcuts.xml file
def delete_shortcuts():
    f = open("Shortcuts.xml", "w")
    f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
    f.write("<Shortcuts>\n")
    f.write("</Shortcuts>\n")
    f.close()
    # Refresh the gallery
    create_gallery(get_image_paths())


# Create a function that creates a gallery of the images, recreated every time the function is called
def create_gallery(image_paths):
    #create a frame to hold the images
    frame = Frame(root)




    # Loop through the list of ImagePaths
    # Sort them in a row
    for image_path in image_paths:
        # Create a clickable image that opens the shortcut
        image = Image.open(image_path.image)
        image = image.resize((200, 200), Image.ANTIALIAS)
        image = ImageTk.PhotoImage(image)
        label = Label(frame, image=image)
        label.image = image
        label.bind("<Button-1>", lambda event, arg=image_path.shortcut: open_shortcut(arg))
        # Pack the label to the left and center the images
        label.pack(side=LEFT)    # Pack the frame to the left and center it in the root window
    frame.pack(side=LEFT, anchor=CENTER)


    
    return frame

# Create a version of the create_gallery that takes a single ImagePath object
def create_gallery_single(image_path):
    # Create a clickable image that opens the shortcut
    image = Image.open(image_path.image)
    image = image.resize((200, 200), Image.ANTIALIAS)
    image = ImageTk.PhotoImage(image)
    label = Label(frame, image=image)
    label.image = image
    label.bind("<Button-1>", lambda event, arg=image_path.shortcut: open_shortcut(arg))
    label.pack()


def open_shortcut(path):
    os.startfile(path)

# Create a button that opens the file dialog to select a shortcut file
def open_file_dialog():
    file_path = fd.askopenfilename(initialdir="C:\\", title="Select a shortcut file", filetypes=(
        ("Shortcut files", "*.lnk"), ("All files", "*.*")))
    if file_path != "":
        if os.path.isfile(file_path):
            # Get the target and icon location of the shortcut

            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(file_path)
            target = shortcut.Targetpath
            

            # if icon ends with .exe, use icoextract to get the icon
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)

            large, small = win32gui.ExtractIconEx(target, 0)
            win32gui.DestroyIcon(small[0])

            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
            hdc = hdc.CreateCompatibleDC()

            hdc.SelectObject(hbmp)
            hdc.DrawIcon((0, 0), large[0])
            #generate a random number
            random_number = random.randint(0, 100000)
            icon_name = "icon_" + str(random_number) + ".ico"
            hbmp.SaveBitmapFile(hdc, icon_name)
            icon = icon_name

            # remove \t and \n from target and icon
            target = target.replace("\t", "").replace("\n", "")

            # Create a new ImagePath object and add it to the xml file.
            image_path = create_image_path(icon, target)
            insert_xml(image_path_to_xml(image_path))
            #append the new image to the gallery
            create_gallery_single(image_path)



# ALlow the user to drag and drop an image onto the root window to add it to the gallery and xml file
def drag_and_drop(event):
    # Get the file path of the image
    file_path = event.data.decode("utf-8")

    # If the file path is an image, add it to the gallery and xml file
    if os.path.isfile(file_path):
        # Get the target and icon location of the shortcut

        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(file_path)
        target = shortcut.Targetpath
        # if icon ends with .exe, use icoextract to get the icon
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)

        large, small = win32gui.ExtractIconEx(target, 0)
        win32gui.DestroyIcon(small[0])

        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
        hdc = hdc.CreateCompatibleDC()

        hdc.SelectObject(hbmp)
        hdc.DrawIcon((0, 0), large[0])
        #generate a random number
        random_number = random.randint(0, 100000)
        icon_name = "icon_" + str(random_number) + ".ico"
        hbmp.SaveBitmapFile(hdc, icon_name)
        icon = icon_name

        # remove \t and \n from target and icon
        target = target.replace("\t", "").replace("\n", "")

        # Create a new ImagePath object and add it to the xml file.
        image_path = create_image_path(icon, target)
        insert_xml(image_path_to_xml(image_path))
        #append the new image to the gallery
        create_gallery_single(image_path)



# Calculate the minimum size of the window based on the number of images and set the window size to that size
def calculate_window_size(image_paths):
    # Calculate the minimum size of the window based on the number of images and set the window size to that size
    window_width = 500
    window_height = 500
    if len(image_paths) > 0:
        window_width = len(image_paths) * 200
        window_height = 200
    return window_width, window_height

# Set minsize of the window based on the number of images using the calculate_window_size function
def set_window_size(image_paths):
    window_width, window_height = calculate_window_size(image_paths)
    root.minsize(window_width, window_height)


root = Tk()
root.title("Image Gallery")
root.geometry("500x500")






# Create a button that deletes all the Shortcuts.xml file
delete_button = Button(root, text="Delete All Shortcuts", command=delete_shortcuts)
delete_button.pack()




set_window_size(get_image_paths())
# create a UI button that opens the file dialog
button = Button(root, text="Select a shortcut file", command=open_file_dialog)
button.pack()

create_gallery(get_image_paths())




root.mainloop()
