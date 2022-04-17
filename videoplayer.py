import cv2
import numpy as np
import PySimpleGUI as sg
import PIL 
from PIL import Image, ImageTk
import tkinter as tk
import sys
import threading
import time

class VideoPlayer:

    def __init__(self):
        
        self.play = True
        self.mask = False
        self.delay = 0.023
        self.frame = 1
        self.frames = None
        self.lower_color = 0
        self.upper_color = 255

        self.vid = None
        self.photo = None
        self.next = '1'

        # menu_def = [['&File', ['&Open', '&Save', '---', 'Properties', 'E&xit']],
        #            ['&Edit', ['Paste', ['Special', 'Normal', ], 'Undo'],],
        #            ['&Help', '&About...']]

        layout = [
            # [sg.Menu(menu_def)],
             [sg.Text('Select video')], [sg.Input(key='-FILEPATH-'), sg.Button('Browse')],
             [sg.Canvas(size=(500, 300), key='canvas', background_color='white', border_width=1)],
             [sg.Slider(size=(30, 20), range=(0, 100), resolution=1, key='slider', orientation='h', 
             enable_events=True), sg.T('0', key='counter', size=(10, 1))],
             [sg.Button('Next frame'), sg.Button('Pause', key='Play'), sg.Button('Mask', key='Mask'),
             sg.Button('Exit')],
             [sg.Slider(size=(30, 20), range=(0, 255), default_value = 0, resolution=1, key='-LOWER-', orientation='h', enable_events=True), 
            sg.Slider(size=(30, 20), range=(0, 255), default_value = 255, resolution=1, key='-UPPER-', orientation='h', enable_events=True)],
             [sg.Radio('2D: Intensity/time', 'SELECT_TYPE_OF_GRAPH',  key='-INT_T-', default=True),
             sg.VSeperator(),
             sg.Radio('2D: Intensity/width', 'SELECT_TYPE_OF_GRAPH', key='-INT_W-'),
             sg.InputText(size=(5, 10), key='-TIME-'), 
             sg.Text('ms'),
             sg.VSeperator(),
             sg.Radio('3D: Intensity/time/width', 'SELECT_TYPE_OF_GRAPH',  key='-3D_INT-')],
             [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_VIDEO-', font='Helvetica 16')]]

        self.window = sg.Window('Videoplayer', layout).Finalize()

        canvas = self.window.Element('canvas')
        self.canvas = canvas.TKCanvas

        self.load_video()

        while True:
            event, values = self.window.Read()

            if event is None or event == 'Exit':
                break
            if event == 'Browse':
                video_path = None
                try:
                    video_path = sg.filedialog.askopenfile().name
                except AttributeError:
                    print('no video selected, doing nothing')

                if video_path:

                    self.vid = MyVideoCapture(video_path)

#Need add scale
                    self.vid_width = int(self.vid.width * 0.6)
                    self.vid_height = int(self.vid.height * 0.6)

                    self.frames = int(self.vid.frames)

                    self.window.Element('slider').Update(range=(0, int(self.frames)), value=0)
                    self.window.Element('counter').Update('0/%i' % self.frames)
                    self.canvas.config(width=self.vid_width, height=self.vid_height)

                    self.frame = 0
                    self.delay = 1 / self.vid.fps

                    self.window.Element('-FILEPATH-').Update(video_path)

            if event == 'Play':
                if self.play:
                    self.play = False
                    self.window.Element('Play').Update('Play')
                else:
                    self.play = True
                    self.window.Element('Play').Update('Pause')

            if event == 'Next frame':
                self.set_frame(self.frame + 1)

            if event == 'slider':
                self.set_frame(int(values['slider']))

            if event == 'Mask':
                if not self.mask:
                    self.mask = True
                    self.window.Element('Mask').Update('Unmask')
                    if not self.play:
                        self.set_frame(self.frame)
                else:
                    self.mask = False
                    self.window.Element('Mask').Update('Mask')
                    if not self.play:
                        self.set_frame(self.frame)

            self.lower_color = int(values['-LOWER-'])
            self.upper_color = int(values['-UPPER-'])

        self.window.close()
        sys.exit

    def load_video(self):
        thread = threading.Thread(target=self.update, args=())
        thread.daemon = 1
        thread.start()
        
    def update(self):
        start_time = time.time()

        if self.vid:
            if self.play:

                ret, frame = self.vid.get_frame(self.mask, self.lower_color, self.upper_color)

                if ret:
                    self.photo = PIL.ImageTk.PhotoImage(
                        image=PIL.Image.fromarray(frame).resize((self.vid_width, self.vid_height), Image.NEAREST)
                    )
                    self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

                    self.frame +=1
                    self.update_counter(self.frame)
        self.canvas.after(abs(int((self.delay - (time.time() - start_time)) * 1000)), self.update)

    def set_frame(self, frame_no):

        if self.vid:

            ret, frame = self.vid.goto_frame(frame_no, self.mask, self.lower_color, self.upper_color)
            self.frame = frame_no
            self.update_counter(self.frame)

            if ret:
                self.photo = PIL.ImageTk.PhotoImage(
                    image=PIL.Image.fromarray(frame).resize((self.vid_width, self.vid_height), Image.NEAREST)
                )
                self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

    def update_counter(self, frame):

        self.window.Element('slider').Update(value=frame)
        self.window.Element('counter').Update('{}/{}'.format(frame, self.frames))


class MyVideoCapture:

    def __init__(self, video_source):
        
        self.vid = cv2.VideoCapture(video_source)
        if not self.vid.isOpened():
            raise ValueError('Unable to open video source', video_source)

        self.width = self.vid.get(cv2.CAP_PROP_FRAME_WIDTH)
        self.height = self.vid.get(cv2.CAP_PROP_FRAME_HEIGHT)
        self.frames = self.vid.get(cv2.CAP_PROP_FRAME_COUNT)
        self.fps = self.vid.get(cv2.CAP_PROP_FPS)

    def get_frame(self, mask, lower, upper):

        if self.vid.isOpened():
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                if mask:
                    mask = cv2.inRange(frame, lower, upper)
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None

    def goto_frame(self, frame_no, mask, lower, upper):

        if self.vid.isOpened():
            self.vid.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                if mask:
                    mask = cv2.inRange(frame, lower, upper)
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None


    def __del__(self):
        if self.vid.isOpened():
            self.vid.release()


if __name__=='__main__':
    VideoPlayer()