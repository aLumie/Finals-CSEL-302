import tkinter as tk
import threading
from PIL import ImageTk, Image
from main import vafs
import cv2

# Create a Tkinter window
window = tk.Tk()
window.title("VaFs by Dotes")
window.geometry("600x400")

window.iconbitmap("D:/python/science.ico")

# Load the image
image = Image.open("D:/python/backround.jpg")
background_image = ImageTk.PhotoImage(image)

# Create a Label widget to display the background image
background_label = tk.Label(window, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)


def display_video():
    cap = cv2.VideoCapture("D:/python/motion.mp4")

    while True:
        ret, frame = cap.read()
        if not ret:

            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            continue

        image = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        image = image.resize((350, 265))

        photo_image = ImageTk.PhotoImage(image)


        video_label.configure(image=photo_image)
        video_label.image = photo_image


        video_label.place(x=(window.winfo_width() - image.width) // 2, y=(window.winfo_height() - image.height) // 2)

        # Check if the user closed the program
        if not window.winfo_exists():
            cap.release()
            return


video_label = tk.Label(window)
video_label.place(x=0, y=0)  # Position the video label initially


video_thread = threading.Thread(target=display_video)
video_thread.start()



def run_vafs():
    t = threading.Thread(target=vafs)
    t.start()

# Button to start the voice assistant
start_button = tk.Button(window, text="OoO", command=run_vafs, font=("Helvecita", 13, "bold"))
start_button.pack(side=tk.BOTTOM, pady=10)


def on_close():
    window.destroy()

window.protocol("WM_DELETE_WINDOW", on_close)  # Call on_close when the window is closed

# Start the Tkinter event loop
window.mainloop()
