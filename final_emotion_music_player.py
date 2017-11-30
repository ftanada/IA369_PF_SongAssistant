# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 00:16:46 2017

@author: ftanada
http://www.paulvangent.com/2016/06/30/making-an-emotion-aware-music-player/
"""

import cv2, numpy as np, argparse, time, glob, os, sys, subprocess, pandas, random, Update_Model, math, ctypes, win32con, time
from pygame import mixer # Load the required library
import win32com.client  # FMT voice synthetzer
import ctypes  # FMT windows dialog
import easygui  # FMT windows 

##  Styles:
##  0 : OK
##  1 : OK | Cancel
##  2 : Abort | Retry | Ignore
##  3 : Yes | No | Cancel
##  4 : Yes | No
##  5 : Retry | No 
##  6 : Cancel | Try Again | Continue
def Mbox(text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, "Song Mood Assistant", style)
    
# FMT initialization
#Define variables and load classifier
camnumber = 0
video_capture = cv2.VideoCapture()
facecascade = cv2.CascadeClassifier("haarcascade_frontalface_default.xml")
fishface = cv2.face.FisherFaceRecognizer_create()
#model = cv2.face.FisherFaceRecognizer_create()
mixer.init()
voiceEngine = win32com.client.Dispatch("SAPI.SpVoice")  # FMT initiaiing voice

print("Welcome to your Song Mood Assistant")
voiceEngine.Speak("Welcome to your Song Mood Assistant.")
voiceEngine.Speak("Please wait while I get prepared.")
Mbox("Welcome to your Song Mood Assistant",0)
#easygui.msgbox("Welcome to your Song Mood Assistant", title="Song Mood Assistant")

try:
    fishface.read("trained_emoclassifier.xml")
    print("trained_emoclassifier.xml successfully loaded.")
except:
    print("no trained xml file found, please run program with --update flag first")
parser = argparse.ArgumentParser(description="Options for the emotion-based music player")
parser.add_argument("--update", help="Call to grab new images and update the model accordingly", action="store_true")
parser.add_argument("--retrain", help="Call to re-train the the model based on all images in training folders", action="store_true") #Add --update argument
parser.add_argument("--wallpaper", help="Call to run the program in wallpaper change mode. Input should be followed by integer for how long each change cycle should last (in seconds)", type=int) #Add --update argument
args = parser.parse_args()
facedict = {}
actions = {}
emotions = ["anger", "happy", "sadness", "neutral"]
df = pandas.read_excel("EmotionLinks.xlsx") #open Excel file
actions["anger"] = [x for x in df.anger.dropna()] #We need de dropna() when columns are uneven in length, which creates NaN values at missing places. The OS won't know what to do with these if we try to open them.
actions["happy"] = [x for x in df.happy.dropna()]
actions["sadness"] = [x for x in df.sadness.dropna()]
actions["neutral"] = [x for x in df.neutral.dropna()]
print("Loaded EmotionLinks.xlsx")

def open_stuff(filename): #Open the file, credit to user4815162342, on the stackoverflow link in the text above
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])

def crop_face(clahe_image, face):
    for (x, y, w, h) in face:
        faceslice = clahe_image[y:y+h, x:x+w]
        faceslice = cv2.resize(faceslice, (350, 350))
    facedict["face%s" %(len(facedict)+1)] = faceslice
    return faceslice

def update_model(emotions):
    print("Model update mode active")
    check_folders(emotions)
    for i in range(0, len(emotions)):
        save_face(emotions[i])
    print("collected images, looking good! Now updating model...")
    Update_Model.update(emotions)
    print("Done!")

def check_folders(emotions):
    for x in emotions:
        if os.path.exists("dataset\\%s" %x):
            pass
        else:
            os.makedirs("dataset\\%s" %x)

def save_face(emotion):
    print("\n\nplease look " + emotion + ". Press enter when you're ready to have your pictures taken")
    input() #Wait until enter is pressed with the raw_input() method
    video_capture.open(camnumber)
    while len(facedict.keys()) < 16:
        detect_face()
    video_capture.release()
    for x in facedict.keys():
        cv2.imwrite("dataset\\%s\\%s.jpg" %(emotion, len(glob.glob("dataset\\%s\\*" %emotion))), facedict[x])
    facedict.clear() 
    
def recognize_emotion():
    predictions = []
    confidence = []
    for x in facedict.keys():
        pred, conf = fishface.predict(facedict[x])
        cv2.imwrite("output\\%s.jpg" %x, facedict[x])
        predictions.append(pred)
        confidence.append(conf)
        print(x," prediction=",pred," confidence=",conf)
    recognized_emotion = emotions[max(set(predictions), key=predictions.count)]
    print("I think you're %s" %recognized_emotion)
    return recognized_emotion

def grab_webcamframe():
    ret, frame = video_capture.read()
    cv2.imwrite('output\\frame.jpg',frame)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    #cv2.imshow('frame',gray)  # FMT debug
    cv2.imwrite('output\\gray.jpg',gray)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    clahe_image = clahe.apply(gray)
    cv2.imwrite('output\\clahe2.jpg',clahe_image)
    return clahe_image

def detect_face():
    clahe_image = grab_webcamframe()
    face = facecascade.detectMultiScale(clahe_image, scaleFactor=1.1, minNeighbors=15, minSize=(10, 10), flags=cv2.CASCADE_SCALE_IMAGE)
    #print("Output from detecMultipleScale",len(face))
    if len(face) == 1: 
        faceslice = crop_face(clahe_image, face)
        cv2.imwrite('output\\crop.jpg',faceslice)
        return faceslice
    else:
        print("no/multiple faces detected, passing over frame")

def run_detection():
    while len(facedict) != 10:
        faceslice = detect_face()
    print("Generated faceslice")
    cv2.imshow('Face Slice',faceslice)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    print("Calling recognize_emotion()")
    recognized_emotion = recognize_emotion()
    return recognized_emotion

def wallpaper_timer(seconds):
    video_capture.release()
    time.sleep(int(seconds))
    video_capture.open(camnumber)
    facedict.clear()

def change_wallpaper(emotion):
    files = glob.glob("wallpapers\\%s\\*.bmp" %emotion)
    current_dir = os.getcwd()
    random.shuffle(files)
    file = "%s\%s" %(current_dir, files[0])
    setWallpaperWithCtypes(file)

def setWallpaperWithCtypes(path): #Taken from http://www.blog.pythonlibrary.org/2014/10/22/pywin32-how-to-set-desktop-background/
    cs = ctypes.c_buffer(path)
    ok = ctypes.windll.user32.SystemParametersInfoA(win32con.SPI_SETDESKWALLPAPER, 0, cs, 0)

def make_train_vector(emotions):
    training_data = []
    training_labels = []

    for emotion in emotions:
        training = glob.glob("dataset\\%s\\*" %emotion)
        for item in training:
            image = cv2.imread(item) 
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) 
            training_data.append(gray)
            training_labels.append(emotions.index(emotion))
            #print("training ",item," for ",emotion, " with index= ",emotions.index(emotion))

    return training_data, training_labels

def run_train_recognizer(emotions):
    training_data, training_labels = make_train_vector(emotions)
    print("training fisher face classifier")
    print("size of training set is: " + str(len(training_labels)) + " images")
    fishface.train(training_data, np.asarray(training_labels))
    fishface.save("trained_new.xml")
    
def play_song(songName, songIndex):
        print("Playing ",songName)
        mixer.music.load(songName)
        mixer.music.play()
        random.shuffle(actionlist) #Randomly shuffle the list
        # wiat the song to finish
        while mixer.music.get_busy(): 
            time.sleep(10)
        songName = "songs\\"+recognized_emotion+"\\"+actionlist[songIndex]+".mp3"
        print("Next song",songName)    
    
# main part
if args.update:
    update_model(emotions)
elif args.retrain:
    Update_Model.update(emotions)
elif args.wallpaper:
    cycle_time = args.wallpaper
    while True:
        wallpaper_timer(cycle_time)
        print("Running detector with wallpaper")
        recognized_emotion = run_detection()
        change_wallpaper(recognized_emotion)
else:
    video_capture.release()  # FMT code fails if webcam is open
    run_train_recognizer(emotions)  # FMT force train despite of the xml read
    
    voiceEngine.Speak("Opening webcam to capture your face so we can detect your mood.")
    voiceEngine.Speak("Please stay with a steady position.")
    easygui.msgbox("Opening webcam to capture your face so we can detect your mood", title="Song Mood Assistant")
    print("Opening webcam to capture your face so we can detect your mood")
    video_capture.open(camnumber)
    print("Running emotion detector")
    recognized_emotion = run_detection()
    print("Recognized emotion: ",recognized_emotion)
    voiceEngine.Speak("I recognized that you are"+recognized_emotion)
    actionlist = [x for x in actions[recognized_emotion]] #get list of actions/files for detected emotion
    
    video_capture.release()  # FMT, close webcam
    songName = "songs\\"+recognized_emotion+"\\"+actionlist[0]+".mp3"
    songIndex = 0
    play_song(songName, songIndex)
    
    voiceEngine.Speak("Would you like to listend to the next one in this set? (y/n)")
    answer = input("Would you like to listend to the next one in this set? (y/n) ")
    while (answer == "y"):
        songIndex += 1
        songName = "songs\\"+recognized_emotion+"\\"+actionlist[songIndex]+".mp3"
        play_song(songName, songIndex)
        voiceEngine.Speak("Would you like to listen to the next one in this set? (y/n)")
        answer = input("Would you like to listen to the next one in this set? (y/n) ")
        
        