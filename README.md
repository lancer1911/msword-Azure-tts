# msword-Azure-tts

[中文版本](./README-zh.md)

This Python script converts a Microsoft Word document (`.docx`) into an MP3 audio file using Azure Cognitive Services Text-to-Speech API.

## Prerequisites

1. Python 3.6 or higher
2. An Azure subscription key for the Text-to-Speech service. Follow the instructions at [BobTranslate](https://bobtranslate.com/service/translate/microsoft.html#_2-%E6%B3%A8%E5%86%8C-azure) to obtain an API key.
3. The region for the Azure Text-to-Speech service.
4. A voice shortname for the Text-to-Speech service. A list of available voices can be found at [Voices.md](https://github.com/playht/text-to-speech-api/blob/master/Voices.md).

## Installation

1. Clone the repository or download the source code.
2. Install the required dependencies:

   ```
   pip install -r requirements.txt
   ```

3. Open the `settings.cfg` file and add your Azure subscription key, region, voice shortname, and speech recognition language:

   ```
   [Azure]
   subscription_key = your_subscription_key
   region = your_region
   voice_shortname = voice_shortname # e.g. en-US-EricNeural
   speech_recognition_language = your_recognition_language # e.g. en-US
   ```

   Replace `YOUR_SUBSCRIPTION_KEY`, `YOUR_REGION`, `YOUR_VOICE_SHORTNAME`, and `YOUR_SPEECH_RECOGNITION_LANGUAGE` with the appropriate values.

## Usage

You can run the script by providing a `.docx` file path as a command-line argument or by selecting the file using a file dialog.

### Using Command Line

```bash
python msword-Azure-tts.py sample.docx
```

Replace `sample.docx` with the path to your Word document.

### Using File Dialog

Run the script without any command-line arguments:

```bash
python msword-Azure-tts.py
```

A file dialog will appear, allowing you to select the Word document to convert.

The script will save the generated MP3 audio file in the same directory as the input file with the same name and an `-Azure-tts.mp3` suffix. For example, if the input file is named `sample.docx`, the output file will be named `sample-Azure-tts.mp3`.
```
