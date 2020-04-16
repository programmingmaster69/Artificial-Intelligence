import pyttsx3
import time
import speech_recognition as sr
import wikipedia
import os
import webbrowser
from PyDictionary import PyDictionary
from pygame import mixer
import win32api
import pathlib
from functools import lru_cache

o = 1

engine = pyttsx3.init('sapi5')
voice = engine.getProperty('voices')


def speak_and_print(words0):
    print(words0)
    str(words0).format('\n')
    engine.say(words0)
    engine.runAndWait()


def wish_user():
    hour = time.localtime()[3]
    if 0 <= hour < 12:
        speak_and_print("""Good Morning!
I am Jarvis""")
    elif 12 <= hour < 17:
        speak_and_print("""Good Afternoon!
I am Jarvis!""")
    else:
        speak_and_print("""Good Evening!
I am Jarvis!""")


def say_jarvis():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.energy_threshold = 400
        try:
            audio = r.listen(source)
        except Exception:
            speak_and_print("Please check your microphone...")
        try:
            print("Recognizing....")
            u_command = r.recognize_google(audio, language="en-in")
            print(f"You said: {u_command}")
            return u_command.lower()
        except Exception:
            pass
            return 'None'


def search_wiki(u_words):
    if 'wikipedia' in u_words:
        u_words = u_words.replace('wikipedia', '')
    elif 'who is' in u_words:
        u_words = u_words.replace('who is', '')
    elif 'what is' in u_words:
        u_words = u_words.replace('what is', '')
    elif 'tell me' in u_words and 'about' in u_words and 'yourself' not in u_words:
        u_words = u_words.replace('tell me about', '')
    speak_and_print('Please wait for a second....')
    try:
        results = wikipedia.summary(u_words, sentences=2)
        speak_and_print(results)
    except ConnectionError or WindowsError or TimeoutError:
        speak_and_print("Please Check your Internet connection and Try again")
    except Exception:
        speak_and_print("Please be more specific.....")


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.7
        r.energy_threshold = 400
        try:
            audio = r.listen(source)
        except Exception:
            speak_and_print("Please check your microphone and try again...")
    try:
        print("Recognizing...")
        u_command = r.recognize_google(audio, language="en")
        print(f"You said: {str(u_command).capitalize()}")
        return u_command.lower()
    except Exception:
        return ''


def play_music(file0):
    global o
    mixer.init()
    mixer.music.load(file0)
    if o == 1:
        speak_and_print("""You can type 'stop' to stop the song and play another song if you want to 
or you can type 'terminate' to close the music player.
You can also type 'pause' to pause the song and 'resume' to resume the song. """)
        o = 2
    else:
        pass
    mixer.music.play()
    has_started = True
    while True:
        user_input = input().lower()
        if user_input == 'terminate':
            mixer.quit()
            break
        elif user_input == 'stop':
            mixer.music.stop()
            speak_and_print("Do you want me to play another song for you?")
            print("Listening...")
            rechoice = take_command()
            while True:
                if 'yes' in rechoice or 'yeah' in rechoice or 'yup' in rechoice or 'yep' in rechoice or 'ok' in rechoice or 'alright' in rechoice or 'of course' in rechoice:
                    music_files_actual_path0, musicfiles0 = computer_scanner_and_music_displayer()
                    i = 1
                    for filess in musicfiles0:
                        print(f"{i}.{filess}")
                        i += 1
                    speak_and_print("Which song do you want me to play?(Please tell me the number of the song)...")
                    userchoice = take_command()
                    if userchoice.isdigit():
                        userchoice = int(userchoice)
                        filetoplay = music_files_actual_path0[userchoice - 1]
                        play_music(filetoplay)
                        break
                    else:
                        speak_and_print("Sorry. I was expecting a number from you...")
                elif 'no' in rechoice or "don't" in rechoice:
                    break
            break

        elif user_input == 'pause':
            has_started = False
            mixer.music.pause()

        elif user_input == 'resume' and has_started is False:
            has_started = True
            mixer.music.unpause()

        elif user_input == 'resume' and has_started is True:
            print("The music is already being played...")


@lru_cache(maxsize=7)
def computer_scanner_and_music_displayer():
    exclusions = [f'{pathlib.Path.home().drive}\\ProgramData',
                  f'{pathlib.Path.home().drive}\\Windows',
                  f'{pathlib.Path.home().drive}\\Program Files',
                  f'{pathlib.Path.home().drive}\\Program Files(x86)']
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\0')[:-1]
    music_files_actualpath = []
    music_files = []
    for drive in drives:
        for root, dirs, files in os.walk(drive):
            dirs[:] = [d for d in dirs if d not in exclusions]
            for file in files:
                if file.endswith('.mp3'):
                    music_files_actualpath.append(root + '\\' + (str(file)))
                    music_files.append(str(file))
    i = 1
    for file1 in music_files:
        print(f"{i}.{file1}")
        i += 1
    return music_files_actualpath, music_files


def meaning_finder(word0):
    dictionary = PyDictionary()
    try:
        meaning = dictionary.meaning(word0)
        for key, value in meaning.items():
            speak_and_print(f"As a {key},it's meaning is:\n")
            for line in list(value):
                doc = f'''{str(line).capitalize()}'''
                speak_and_print(f"{doc}")
    except Exception:
        pass


if __name__ == '__main__':
    wish_user()
    while True:
        speak_and_print("How may I help you?")
        print("Listening...")
        user_command = take_command()

        if (('wikipedia' in user_command) and ('.com' or '.org' not in user_command)) or ('who is' in user_command) or ('what is' in user_command and 'your' not in user_command and 'meaning of' not in user_command) or (('tell me about' in user_command) and ('yourself' not in user_command)) or ('describe' in user_command and 'yourself' not in user_command):
            search_wiki(user_command)

        elif ('.com' in user_command) or ('.org' in user_command):
            if '.com' in user_command:
                url = user_command.split('.')
                url = url[url.index('com') - 1].split()
                webbrowser.open(f"www.{url[len(url) - 1]}.com")
                speak_and_print('Done')

            elif '.org' in user_command:
                url = user_command.split('.')
                url = url[url.index('org') - 1].split()
                webbrowser.open(f"www.{url[len(url) - 1]}.org")
                speak_and_print('Done')

        elif 'new tab' in user_command:
            webbrowser.open_new_tab('www.google.com')
            speak_and_print('Done')

        elif ('hello' in user_command) or ('hi' in user_command) or ('yo ' in user_command):
            speak_and_print('Hello!')

        elif (('play' in user_command) and ('music' in user_command)) or ('play a song' in user_command):
            speak_and_print("Please wait for a few minutes...\nI'm scanning your computer for songs...")
            music_files_actual_path, musicfiles = computer_scanner_and_music_displayer()
            while True:
                speak_and_print("Which song do you want me to play?(Please tell me the number of the song)...")
                print("Listening...")
                user_choice = take_command()

                if user_choice.isdigit():
                    user_choice = int(user_choice)
                    file_to_play = music_files_actual_path[user_choice - 1]
                    play_music(file_to_play)
                    break
                else:
                    speak_and_print("Sorry. I was expecting a number from you...")

        elif ('time' in user_command) and ('meaning of' not in user_command or 'mean' not in user_command):
            time0 = time.gmtime()
            speak_and_print(f"The time is {time0[3]} hours and {time0[4]} minutes...")

        elif ('about yourself' in user_command) or ('introduce' in user_command) or ('introduction' in user_command) and ('meaning of' not in user_command or 'mean' not in user_command):
            speak_and_print("I'm Jarvis.\nMy work is to help you and assist you.\nI'm created by Arjun ")

        elif 'a joke' in user_command:
            speak_and_print("Sorry. I don't know any joke as I don't have a good sense of humour...")

        elif 'your name' in user_command:
            speak_and_print('My name is Jarvis')

        elif ('idiot' in user_command) or ('stupid' in user_command) or ('useless' in user_command) and (('meaning of' not in user_command) or ('mean' not in user_command)):
            speak_and_print("Sorry. I don't like abusive words...")

        elif 'meaning of' or 'mean' in user_command:
            if 'meaning of' in user_command:
                user_command = user_command.split('meaning of')
                if len(user_command) != (1 or 0):
                    user_command = user_command[1].split()
                    while len(user_command) > 1:
                        del user_command[len(user_command) - 1]
                    meaning_finder(user_command[0])
                else:
                    speak_and_print("Sorry. I didn't get you ...")

            elif 'mean' in user_command:
                user_command = user_command.split('mean')
                if len(user_command) != 0:
                    user_command = user_command[0].split()
                    user_command = user_command[len(user_command) - 1]
                    meaning_finder(user_command)
                else:
                    speak_and_print("Sorry. I didn't get you ...")
            else:
                speak_and_print("Sorry....?")
