import pyttsx3

def my_speak(message):
    engine= pyttsx3.init()
    engine.say('{}'.format(message))
    engine.runAndWait()

message='''

Hello everyone Welcome to my Youtube Video 
My name is Soumil shah 
i did my Bachelor in electronic engineering 
persuing my  Master in Electrical Engineering 
Master in Computer Engineering

lets get started with this tutorials 

'''

my_speak(message)