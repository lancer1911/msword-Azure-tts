"""
Convert MS Word documents into mp3 by Azure API.
written by lancer1911
May 1, 2023
"""

import sys
import os
import docx
import configparser
import azure.cognitiveservices.speech as speechsdk
from tkinter import filedialog
from tkinter import Tk


def docx_to_text(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)


def text_to_speech(text, output_filename, subscription_key, region, voice_shortname, speech_recognition_language):
    speech_config = speechsdk.SpeechConfig(subscription=subscription_key, region=region, speech_recognition_language=speech_recognition_language)
    speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio16Khz128KBitRateMonoMp3)
    speech_config.speech_synthesis_voice_name = voice_shortname

    audio_config = speechsdk.audio.AudioOutputConfig(filename=output_filename)
    synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)

    result = synthesizer.speak_text_async(text).get()
    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        print(f"Speech synthesized and saved to {output_filename}")
    elif result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = result.cancellation_details
        print(f"Speech synthesis canceled: {cancellation_details.reason}")
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print(f"Error details: {cancellation_details.error_details}")


# Read settings.cfg
config = configparser.ConfigParser()
config.read('settings.cfg')

subscription_key = config.get('Azure', 'subscription_key')
region = config.get('Azure', 'region')
voice_shortname = config.get('Azure', 'voice_shortname')
speech_recognition_language = config.get('Azure', 'speech_recognition_language')

# Get input file path from command line or ask user to select a file
if len(sys.argv) > 1:
    file_path = sys.argv[1]
else:
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])

# Convert Word document to text
text = docx_to_text(file_path)

# Convert text to speech and save as MP3
input_filename_without_ext, _ = os.path.splitext(os.path.basename(file_path))
output_filename = f"{input_filename_without_ext}-Azure-tts.mp3"
text_to_speech(text, output_filename, subscription_key, region, voice_shortname, speech_recognition_language)
