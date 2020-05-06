for %%i in (*.mpeg4) do "./FFmpeg/bin/ffmpeg.exe" -i "%%i" "%%~ni.wav"
mkdir FinalMedia
for %%i in (*.wav) do "./FFmpeg/bin/ffmpeg.exe" -i "%%i" -f segment -segment_time 15 -c copy "./FinalMedia/%%~ni_%%03d.wav"
for %%i in (*.mpeg4) do Del "%%i"