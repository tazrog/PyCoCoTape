#Written by tazrog
#Designed to save programs on the PC via cassette interface and load them in the CoCo.
#Program still in develpment.
#Use on your own risk. No guarantees.
#For use on a Windows PC. Not tested on Linux or Mac.
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import wave
import pyaudio
#from array import array
import os
import sys
import keyboard
import time
import numpy as np
from tqdm import tqdm
from sys import platform

global frame_rate
global auto_sound
global auto_sound_level
global amplify_setting
global filetxt
frame_rate = 9600
auto_sound =1
auto_sound_level = -9.0
amplify_setting = 70

BASE_DIR = os.getcwd()
if not os.path.exists(BASE_DIR+"\\Programs"):
    os.makedirs(BASE_DIR+"\\Programs") 
if not os.path.exists(BASE_DIR+"\\data"):
    os.makedirs(BASE_DIR+"\\data")
if not os.path.exists(BASE_DIR+"\\data\\amp.txt"):    
    with open(BASE_DIR+'\\data\\amp.txt', 'w') as newamp:        
        ampcontent = str(amplify_setting)
        newamp.write(ampcontent)
        newamp.close

with open(BASE_DIR+'\\data\\amp.txt', 'r') as amptxt:
    amplify_setting = amptxt.read()     
    amptxt.close      

def sound():
    global soundlevel
    global volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)    
    volume = interface.QueryInterface(IAudioEndpointVolume)
    if not os.path.exists(BASE_DIR+"\\data\\volume.txt"):    
        with open(BASE_DIR+'\\data\\volume.txt', 'w') as newvol:
            currentVolumeDb = volume.GetMasterVolumeLevel()
            volcontent = str(currentVolumeDb)
            newvol.write(volcontent)
            newvol.close
          
    with open(BASE_DIR+'\\data\\volume.txt', 'r') as getvol:
        soundlevel= volume.GetMasterVolumeLevel()
        vol1content=float(getvol.read())
        volume.SetMasterVolumeLevel(vol1content, None)     
        getvol.close   
          
def soundsave():
    global soundlevel
    global volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)    
    volume = interface.QueryInterface(IAudioEndpointVolume)
    
    with open(BASE_DIR+'\\data\\volume.txt', 'w') as newvol:
            currentVolumeDb = volume.GetMasterVolumeLevel()
            volcontent = str(currentVolumeDb)
            newvol.write(volcontent)
            newvol.close  
    os.system('cls')
    print ("Volume has been saved")
    time.sleep(2)
    main()

def record():
    global filetxt  
    print ("RECORDING SECTION")
    print ("")    
    fcnt = 0
    FORMAT=pyaudio.paInt16
    CHANNELS=1
    RATE=frame_rate
    CHUNK=1024
    RECORD_SECONDS=1
    wavitems = os.listdir(BASE_DIR+"\\Programs")
    wavList = [wav for wav in wavitems if wav.endswith(".wav")]    
    for fcnt, wavName in enumerate(wavList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, wavName))                
    rfn = "tape_" + str(filetxt)+".wav"       
    FILE_NAME=(BASE_DIR+"\\Programs\\" + rfn)    
    print ("RECORDING Autofile...")        
    print ("")
    print (("The Autofile location is at %s") % BASE_DIR+  "\\Programs")
    print ("***********************************************************************************************")
    print (rfn)
    print ("***RECORDING NOW*** PRESS [q] to stop recording when CoCo is done saving.")
    audio=pyaudio.PyAudio() 
    stream=audio.open(format=FORMAT,channels=CHANNELS, 
        rate=RATE,
        input=True,
        frames_per_buffer=CHUNK)            
    #start recording         
    frames=[]         
    while True:           
        for i in range(0,int(RATE/CHUNK*RECORD_SECONDS)):
            if keyboard.is_pressed('q'):  # if key 'q' is pressed 
                print ("Stopping the recording")
                data=stream.read(CHUNK)
                frames.append(data)  
                stream.stop_stream()
                stream.close()
                audio.terminate()
                #end of recording
                stream.stop_stream()
                stream.close()
                audio.terminate()
                #writing to file
                wavfile=wave.open(FILE_NAME,'wb')
                wavfile.setnchannels(CHANNELS)
                wavfile.setsampwidth(audio.get_sample_size(FORMAT))
                wavfile.setframerate(RATE)
                wavfile.writeframes(b''.join(frames))#append frames recorded to file
                wavfile.close()
                os.system('cls')
                print ("Tape recording done for file named:") 
                print (FILE_NAME)                    
                time.sleep (3)
                os.system('cls')
                keyboard.send('backspace')
                amplify_wav(FILE_NAME, FILE_NAME, int(amplify_setting))                
                main()  
            data=stream.read(CHUNK)
            frames.append(data)   

def amplify_wav(input_file, output_file, gain):
    # Open the input WAV file
    with wave.open(input_file, 'rb') as wf:
        params = wf.getparams()
        n_channels, sampwidth, framerate, n_frames = params[:4]
        audio_data = wf.readframes(n_frames)
    # Convert audio data to numpy array
    audio_array = np.frombuffer(audio_data, dtype=np.int16)
    # Amplify the audio data
    amplified_audio_array = audio_array * gain
    # Ensure amplified data does not overflow
    amplified_audio_array = np.clip(amplified_audio_array, -32768, 32767)
    # Convert back to bytes
    amplified_audio_data = amplified_audio_array.astype(np.int16).tobytes()
    # Write the amplified data to a new WAV file
    with wave.open(output_file, 'wb') as wf:
        wf.setnchannels(n_channels)
        wf.setsampwidth(sampwidth)
        wf.setframerate(framerate)
        wf.writeframes(amplified_audio_data)                

def playtape():    
    sound()
    print ("Welcome to CoCo Python tape player and recorder on PC.")
    print ("")
    print ("PLAYING Autofile TAPE SECTION")
    #check for empty directory    
    chunk = 1024    
    print ("")       
    choice= "tape_"+str(filetxt)+".wav"        
    wavplay = BASE_DIR + "\\Programs\\"+ choice
    try:
        f = wave.open(wavplay,"rb")
    except FileNotFoundError:                
            print ("File Not Found. PLease make another selection")
            time.sleep(4)
            main()
    p = pyaudio.PyAudio()        
    stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
        channels = f.getnchannels(),  
        rate = f.getframerate(),  
        output = True)
    frames = f.getnframes()
    rate = f.getframerate()
    duration = (round(frames/rate))
    mins, secs = divmod(duration, 60)    
    print ("File to be played:")
    print (wavplay) 
    print ("Frame Rate: %d"% f.getframerate())  
    print ("Frame Channel: %d"% f.getnchannels())  
    print ('Duration = {:02d}:{:02d}'.format(mins, secs))
    print ("PC audio level is set to %d percent"% (soundlevel + 1))       
    print ("Prepare CoCo and press enter when ready on the PC.")
    print ("Or enter [99] Main Menu")        
    print ("")                    
    print (("Playing file >>> %s")% (choice))    
    print ("")
    data = f.readframes(chunk)        
    countdown= (round(frames/chunk)) + 2
    #play stream  
    for i in tqdm(range(countdown),desc = "Load Progress:"):             
        stream.write(data)  
        data = f.readframes(chunk)
    #stop stream        
    stream.stop_stream()  
    stream.close()
    #close PyAudio    
    p.terminate()
    f.close()  
    print ("Tape Complete")           
    main()      

def list():
    if platform == 'win32':
        os.system('cls') 
    else:
        os.system('clear') 
    print ("Welcome to PyCoCo2Cass tape player and recorder on PC.")
    print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
    print ("")
    fcnt = 0    
    #check for empty directory
    if len(os.listdir(BASE_DIR+"\\Programs")) < 1:
        print ("") 
        print("There are no Programs files to list. Please record some.")
        time.sleep(2)
        main()
    items = os.listdir(BASE_DIR+"\\Programs")
    fileList = [name for name in items if name.endswith(".wav")]
    for fcnt, fileName in enumerate(fileList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, fileName))    
    print ("")
    print ("[99] Main Menu")
    print ("")
    while True:
        choice = (input("Select WAV file [1-%s] to get file details and/or to delete file. >>> "% fcnt))        
        try:
            choice = int(choice)-1
            if (choice+1 == 99):
                main()
            if (choice < 99):
                if (choice > fcnt):
                    print ("INVALID SELECTION")
                    time.sleep(1)
                    list()
        except ValueError:
            os.system('cls')
            print ("INVALID SELECTION")
            time.sleep(1)
            list()
        if (choice < fcnt):
            fn = (fileList[choice])
            wavplay = BASE_DIR + "\\Programs\\"+ fn
            f = wave.open(wavplay,"rb")
            p = pyaudio.PyAudio()        
            stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
                channels = f.getnchannels(),  
                rate = f.getframerate(),  
                output = True)
            frames = f.getnframes()
            rate = f.getframerate()
            duration = (round(frames/rate))
            mins, secs = divmod(duration, 60)
            if platform == 'win32':
                os.system('cls') 
            else:
                os.system('clear') 
            print ("***File Information***")
            print ("File Location %s"% wavplay) 
            print ("")
            print ("File name: %s"% fn)
            print ("Frame Rate: %d"% f.getframerate())  
            print ("Frame Channel: %d"% f.getnchannels())  
            print ('Duration = {:02d}:{:02d}'.format(mins, secs)) 
            print ("")
            a = input ("Do you want to delete file %s? [Y/n] "% fn)
            rf = str(BASE_DIR + "\\Programs\\" + fn)
            #Delete file
            if (a == "y" or a == "Y"):
                x = input("Are you sure you want to delete %s [Y]? or press any other key to return to list menu. >"% fn)                
                f.close()
                if (x == "Y" or x == "y"):
                    os.remove(rf)
                    print ("%s file removed"% fn)
                    time.sleep(1)            
                    main() 
                else:                   
                    print ("INVALID SELECTION")
                    time.sleep(1)
                    list()                    
            else:
                f.close()
                print ("Returning to list menu.")
                time.sleep(2)
                list() 

def settings():
    #in beta
    global frame_rate
    global auto_sound
    global auto_sound_level
    global amplify_setting
    
    print ("Settings")
    print ("OPTIONS")   
    print ("[1] Frame Rate")
    print ("[2] Save Current Volume Level")
    print ("[3] Amplify Setting")    
    print ("")
    print ("[99] Main Menu")
    setinput=int(input(">>> "))    
    if (setinput == 1):         
        print (f'Current Recording Frame Rate is {frame_rate}')
        print ("Frame Rate Options:")
        print ("[1] 1500")
        print ("[2] 9600")
        print ("[3] 44100")
        print ("[4] 48000")        
        print ("")
        b= input(">>> ")
        if (b== "1"):
            frame_rate = 1500
        if (b == "2"):
            frame_rate = 9600
        if (b=="3"):
            frame_rate = 44100
        if (b == "4"):
            frame_rate = 48000        
        frame_rate = int(frame_rate)
        print ("")
        print ("New fram rate is %d: "% frame_rate)
        c = input ("Press any key to go to main menu")
        main()
    if (setinput == 2):
        soundsave()        
    if (setinput == 3):            
        amplify_setting =input("Enter Amplify Settings [Default - 70]) ")
        with open(BASE_DIR+'\\data\\amp.txt', 'w') as newamp:
            ampcontent = str(amplify_setting)
            newamp.write(ampcontent)
            newamp.close
        main()
    if (setinput == 4):            
            print ("***WARNING*** This could delete WAV files.\n")
            with open('file.txt', 'r') as filetxt:         
                file_number = filetxt.read()
                filetxt.close
            file_number =input("Enter file number. Old file number is %s: "% file_number)
            with open('file.txt', 'w') as filetxt:            
                filetxt.write(file_number)    
                filetxt.close
            main()        
    else:
        main()

def main():    
    #interface 
    global filetxt
    global amplify_setting
    os.system('cls')     
    print ("Welcome to CoCo Python tape player and recorder on PC.")    
    print (("The program's Programs files location is at %s") % BASE_DIR+  "\\Programs")    
    print (("Sound amplify is at %s")% str(amplify_setting))
    print ("Please ensure the tape cables are in the right jacks.")
    print ("")
    print ("OPTIONS.")
    print ("[r###] Record AutoFile")  
    print ("[p###] Play AutoFile")
    print ("") 
    print ("[a###] Set & Save Amplify Level")  
    print ("[v] Save Current Volume Level")      
    print ("[f] File List/Delete Programs") 
    print ("[s] Settings")
    print ("[q] Quit")
    while True:
        x=input("Select an option, then press enter >>> ")        
        try:
            x =str(x)
        except ValueError:            
            print ("INVALID OPTION [%s]"% x)
            time.sleep(1)
            main()
        if (x == "s"):
            settings()
        if ("p" in x):
            filetxt= x[1:]
            playtape()        
        if (x == "f"):
            list()
        if (x == "v"):
            soundsave()
        if ("r" in x):
            filetxt=x[1:]
            record()
        if ("a" in x):
            filetxt=x[1:]
            with open(BASE_DIR+'\\data\\amp.txt', 'w') as newamp:             
                newamp.write(filetxt)
                amplify_setting=int(filetxt)
                newamp.close
            os.system('cls')
            print (('Amplify settings is at %s')% str(amplify_setting))
            time.sleep(2)
            main()
        if (x == "q"):
            quit()
        else:            
            print ("INVALID OPTION")
            print ("Not a valid option [%s]."% x)
            time.sleep (1)
            main()
#Program Start
main()