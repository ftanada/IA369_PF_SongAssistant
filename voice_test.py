# -*- coding: utf-8 -*-
"""
Created on Wed Nov 15 18:39:44 2017

@author: ftanada
"""

#import pyttsx

#voiceEngine = pyttsx.init()
#voiceEngine.say("Welcome to you Song Mood Assistant")
#voiceEngine.runAndWait()

import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Speak("Hello, it works!")