# import pyttsx3
# import PyPDF2
# pdfreader = PyPDF2.PdfFileReader(open('story.pdf','rb'))
# speaker = pyttsx3.init()
# for page_num in range(pdfreader.numPages):   
#     text = pdfreader.getPage(page_num).extractText()  ## extracting text from the PDF
#     cleaned_text = text.strip().replace('\n',' ')  ## Removes unnecessary spaces and break lines
#     print(cleaned_text)                ## Print the text from PDF
#     #speaker.say(cleaned_text)        ## Let The Speaker Speak The Text
#     speaker.save_to_file(cleaned_text,'story.mp3')  ## Saving Text In a audio file 'story.mp3'
#     speaker.runAndWait()
# speaker.stop()

import wikipedia
import pyttsx3 
engine = pyttsx3.init('espeak')
voices = engine.getProperty('voices')
# engine.setProperty('voice','english_rp+f2')
engine.setProperty('voice', 'en-scottish')
rate = engine.getProperty('rate')
engine.setProperty('rate', rate-75)
# gender = engine.getProperty('gender')
# engine.SetProperty('gender',female)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
query = input("What Would you like to know?")
results = wikipedia.summary(query, sentences=3)

print(results)
speak(results)

# speak(results)
# import pyttsx3
# engine = pyttsx3.init()
# engine.say('Sally sells seashells by the seashore.')
# engine.say('The quick brown fox jumped over the lazy dog.')
# engine.runAndWait()

# import pyttsx3
# engine = pyttsx3.init()
# voices = engine.getProperty('voices')
# for voice in voices:
#    engine.setProperty('voice', voice.id)
#    print(voice.id)
#    engine.say('The quick brown fox jumped over the lazy dog.')
# engine.runAndWait()