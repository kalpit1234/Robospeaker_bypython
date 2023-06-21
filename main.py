from win32com.client import Dispatch
while True:
    a=input("Enter what to speak\n")
    if(a=="q"):
        break
    convert=Dispatch("SAPI.spVoice")
    convert.Speak(a)




