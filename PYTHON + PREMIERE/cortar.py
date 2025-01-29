from moviepy.editor import *
import cv2
import datetime
import os
import glob


paths = [r""]

save = [r""]


for i in range(0,len(paths),1):
    
    videos_path = glob.glob(os.path.join(paths[i], "*.mp4"))   
    
    for j in range(0,len(videos_path),1):
        video = videos_path[j]
        nome_video = video[-5:]
        if(i == 0 and (j == 1 or j == 2)):
            nome_video = video[-6:]
        
        videoclip = VideoFileClip(video)
        data = cv2.VideoCapture(video)
        frames = data.get(cv2.CAP_PROP_FRAME_COUNT) 
        fps = data.get(cv2.CAP_PROP_FPS)
        seconds = int(frames / fps)

        start_time = 10.02
        end_time = seconds - 10.02

        print (nome_video)
        
        corte = videoclip.subclip(start_time,end_time)
        corte.write_videofile(
            save[i] + '/' + nome_video, codec = 'libx264'
        )