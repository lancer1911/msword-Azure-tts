# msword-Azure-tts

[English version](./README.md)

这个Python脚本可以把Microsoft Word文档（`.docx`）转换为MP3音频文件，使用的是Azure Cognitive Services的文字转语音API。现在，该脚本可以转换长篇文档，通过切分文本和合并生成的音频文件，解决了Azure TTS API的10分钟限制问题。

## 预备条件

1. Python 3.6或更高版本
2. Azure文字转语音服务的订阅密钥。你可以按照 [BobTranslate](https://bobtranslate.com/service/translate/microsoft.html#_2-%E6%B3%A8%E5%86%8C-azure) 的说明获取API密钥。
3. Azure文字转语音服务的地区。
4. 文字转语音服务的声音短名。可用的声音列表可以在 [Voices.md](https://github.com/playht/text-to-speech-api/blob/master/Voices.md) 找到。
5. `ffmpeg` 软件包。用于将长文档分割为小块并组合生成的音频文件。

## 安装

1. 克隆仓库或下载源代码。
   你可以通过以下命令克隆仓库：

   ```
   git clone https://github.com/lancer1911/msword-Azure-tts.git
   ```

   如果你没有安装Git，你可以直接下载源代码。在GitHub上进入仓库的主页面，点击“Code”按钮，然后点击“Download ZIP”。下载ZIP文件后，解压以获取源代码。

2. 安装所需的依赖：

   ```
   pip install -r requirements.txt
   ```

3. 安装`ffmpeg`：

   **macOS:**
   ```
   brew install ffmpeg
   ```

   **Linux (Ubuntu/Debian):**
   ```
   sudo apt-get update
   sudo apt-get install ffmpeg
   ```

   **Windows:**
   从[官方网站](https://ffmpeg.org/download.html#build-windows)下载静态构建。解压下载的文件，并将解压文件中的`bin`目录添加到你的系统路径（PATH）。

4. 打开`settings.cfg`文件，添加你的Azure订阅密钥、地区、声音短名和语音识别语言：

   ```
   [Azure]
   subscription_key = your_subscription_key
   region = your_region
   voice_shortname = voice_shortname # 例如：en-US-EricNeural
   speech_recognition_language = your_recognition_language # 例如：en-US
   ```

   请用合适的值替换 `YOUR_SUBSCRIPTION_KEY`、`YOUR_REGION`、`YOUR_VOICE_SHORTNAME` 和 `YOUR_SPEECH_RECOGNITION_LANGUAGE`。

## 使用方法

你可以通过提供一个`.docx`文件路径作为命令行参数来运行脚本，或者通过文件对话框选择文件。

### 使用命令行

```bash
python msword-Azure-tts.py sample.docx
```

请用你的Word文档路径替换`sample.docx`。

### 使用文件对话框

运行没有任何命令行参数的脚本：

```bash
python msword-Azure-tts.py
```

一个文件对话框会出现，让你选择要转换的Word文档。

脚本会将生成的MP3音频文件保存在与输入文件相同的目录中，文件名相同，后缀为`-Azure-tts.mp3`。例如，如果输入文件名为`sample.docx`，则输出文件将被命名为`sample-Azure-tts.mp3`。
