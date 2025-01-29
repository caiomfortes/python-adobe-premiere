from moviepy.editor import *
import cv2
import datetime
import os
import glob
import ffmpeg

path = r""


videos_path = glob.glob(os.path.join(path, "*.mp4"))   
for j in range(0,len(videos_path),1):
    # Carregar o vídeo MP4
    video = ffmpeg.input(videos_path[j])

    # Converte o vídeo para MOV
    conver = video.output(videos_path[j][-3:] + '.mov')

    # Executa a conversão
    conver.run()