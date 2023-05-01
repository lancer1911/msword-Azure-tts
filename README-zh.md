# msword-Azure-tts

[English version](./README.md)

该 Python 脚本使用 Azure Cognitive Services 文字转语音 API 将 Microsoft Word 文档（`.docx`）转换为 MP3 音频文件。

## 先决条件

1. Python 3.6 或更高版本
2. 文字转语音服务的 Azure 订阅密钥。请按照 [BobTranslate](https://bobtranslate.com/service/translate/microsoft.html#_2-%E6%B3%A8%E5%86%8C-azure) 上的说明获取 API 密钥。
3. Azure 文字转语音服务的区域。
4. 文字转语音服务的语音简称。可在 [Voices.md](https://github.com/playht/text-to-speech-api/blob/master/Voices.md) 找到可用语音列表。

## 安装

1. 克隆存储库或下载源代码。
2. 安装所需的依赖项：

   ```
   pip install -r requirements.txt
   ```

3. 打开 `settings.cfg` 文件，添加 Azure 订阅密钥、区域、语音简称和语音识别语言：

   ```
   [Azure]
   subscription_key = your_subscription_key
   region = your_region
   voice_shortname = voice_shortname # 例如：en-US-EricNeural
   speech_recognition_language = your_recognition_language # 例如：en-US
   ```

   用适当的值替换 `YOUR_SUBSCRIPTION_KEY`、`YOUR_REGION`、`YOUR_VOICE_SHORTNAME` 和 `YOUR_SPEECH_RECOGNITION_LANGUAGE`。

## 使用方法

您可以通过在命令行中提供 `.docx` 文件路径作为参数，或者使用文件对话框选择文件来运行脚本。

### 使用命令行

```bash
python msword-Azure-tts.py sample.docx
```

将 `sample.docx` 替换为您的 Word 文档的路径。

### 使用文件对话框

不带任何命令行参数运行脚本：

```bash
python msword-Azure-tts.py
```

将出现一个文件对话框，允许您选择要转换的 Word 文档。

脚本会将生成的 MP3 音频文件保存在与输入文件相同的目录中，文件名相同，后缀为 `-Azure-tts.mp3`。例如，如果输入文件名为 `sample.docx`，则输出文件名为 `sample-Azure-tts.mp3`。