import cv2
from tkinter import *
from PIL import Image,ImageTk
import datetime

win = Tk()
win.geometry("1200x650")
win.title('Portable Optical Spectrometer')

L2 = Label(win,text = " Camera  ",font=("times new roman",20,"bold"),bg="white",fg="red").grid(row=0, column=0)
L2 = Label(win,text = " Spectrum Graph ",font=("times new roman",20,"bold"),bg="white",fg="red").grid(row=0, column=2)
label =Label(win)
label.grid(row=1, column=0)
cap = cv2.VideoCapture(0)

def show_frames():
   im = cap.read()[1]
   Image1= cv2.cvtColor(im,cv2.COLOR_BGR2RGB)
   img = Image.fromarray(Image1)
   imgtk = ImageTk.PhotoImage(image = img)
   label.imgtk = imgtk
   label.configure(image=imgtk)
   label.after(20, show_frames)

def capture():
    I = cap.read()[1]
    save_image(I)

def save_image(I):
   time = str(datetime.datetime.now().today()).replace(":", " ") + ".jpg"
   cv2.imwrite(time, I)

show_frames()
B1 = Button(win,text="Capture",font=("Times new roman",20,"bold"),bg="white",fg="red",command=capture()).grid(row=3, column=0, )
B1 = Button(win,text="Analysis",font=("Times new roman",20,"bold"),bg="white",fg="red").grid(row=4, column=0)

L1 = Label(win,text = "Detected Material is:  ",font=("times new roman",20,"bold"),bg="white",fg="red").grid(row=4, column=1)
Output = Text(win, height = 3,width = 25,bg = "light cyan").grid(row=4, column=2)

win.mainloop()