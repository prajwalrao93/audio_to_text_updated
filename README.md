"# audio_to_text_updated"

Audio format conversion (to .wav format) using FFmpeg
Audio to text conversion using Python


Packages used (for exact versions used refer requirements.txt):
1. tkinter and ttkthemes for the GUI application
2. threading to make the gui interactive
3. openpyxl to work with excel files
4. speechRecognition and wit for the audio to text conversion
5. os and shutil to work with files and folders.


API used for conversion
1. wit.ai - open source api (https://wit.ai/faq)


Build the executable file of the application using "cx_freeze" package. Run the command "python setup.py build" using the command line. Execuble file can be found inside "build\exe.win-amd64-3.8" folder.


Steps to follow while using the app
1. Place the media files in the Media folder.
2. Run the Converter.bat file in the Media folder and wait for it to get completed.
3. Once the bat file has run, now run the speech_to_text.exe file.
4. Copy the folder path of the media folder and paste in the first open space (or click the button next to open box and select the folder)
5. Copy the raw data file path along with file name in the second open space (or click the button next to open box and select the file)
6. Click on Submit button. Wait for the application to run and complete
7. Use Clear button to clear the fields
8. Use Exit button to exit the application
9. Use Help button for help on how to use the application


Challenges faced currently (to be updated if there is any):
1. Reducing the noise in the audio file
2. After trimming the silent part in the audio file, the updated audio is not getting uploaded to speech_recognition package and the program crashes.
3. Audio conversion task needs to be done seperately, cannot be included in the main script because command line interface keeps popping up for every file converted.
4. Lots of "Request Timeout" (Updated the message to "Audio not Recognized". Exact reason unknown, possible reasons due to wit server down, slow network connection, bad audio quality)