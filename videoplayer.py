import cv2
import numpy as np
import PySimpleGUI as sg
import PIL 
from PIL import Image, ImageTk
import tkinter as tk
import sys
import threading
import time
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.ticker import LinearLocator, FormatStrFormatter

class VideoPlayer:

    def __init__(self):
        
        self.play = True
        self.mask = False
        self.delay = 0.023
        self.frame = 1
        self.frames = None
        self.get_channel = False

        self.vid = None
        self.photo = None
        self.next = '1'

        self.extended_properties = {
            'lower_color': 0,
            'upper_color': 255,
            'mask': False,
            'get_channel': False,
            'channel': {
                'x1': 1,
                'x2': 1,
                'y1': 1,
                'y2': 1
            }
        }

        layout = [
             [sg.Text('Select video')], [sg.Input(key='-FILEPATH-'), sg.Button('Browse')],
             [sg.Canvas(size=(700, 200), key='-CANVAS-', background_color='white', border_width=1)],
             [sg.Slider(size=(30, 20), range=(0, 100), resolution=1, key='-FRAMES-', orientation='h', 
             enable_events=True), sg.T('0', key='-FRAMES_COUNTER-', size=(10, 1))],
             [sg.Button('Next frame'), sg.Button('Pause', key='Play'), sg.Button('Mask', key='-MASK-'),
             sg.Button('Exit')],
             [sg.Slider(size=(30, 20), range=(0, 255), default_value = 0, resolution=1, key='-LOWER-', orientation='h', enable_events=True), 
            sg.Slider(size=(30, 20), range=(0, 255), default_value = 255, resolution=1, key='-UPPER-', orientation='h', enable_events=True)],
             [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL-', font='Helvetica 16')],
             [sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-LEFT-', orientation='h', enable_events=True),
             sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-RIGHT-', orientation='h', enable_events=True)],
             [sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-TOP-', orientation='h', enable_events=True),
             sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-BOTTOM-', orientation='h', enable_events=True)],
             [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_VIDEO-', font='Helvetica 16')],
             [sg.Radio('2D: Intensity/time', 'SELECT_TYPE_OF_GRAPH',  key='-INT_T-', default=True),
             sg.VSeperator(),
             sg.Radio('2D: Intensity/width', 'SELECT_TYPE_OF_GRAPH', key='-INT_W-'),
             sg.InputText(size=(5, 10), key='-TIME-'), 
             sg.Text('ms'),
             sg.VSeperator(),
             sg.Radio('3D: Intensity/time/width', 'SELECT_TYPE_OF_GRAPH',  key='-3D_INT-')],]

        self.window = sg.Window('Videoplayer', layout).Finalize()

        canvas = self.window.Element('-CANVAS-')
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

                    self.vid = MyVideoCapture(video_path, self.extended_properties)

#Need add scale
                    self.vid_width = int(self.vid.width / self.vid.height * 200)
                    self.vid_height = 200

                    self.frames = int(self.vid.frames)

                    self.window.Element('-FRAMES-').Update(range=(0, int(self.frames)), value=0)
                    self.window.Element('-FRAMES_COUNTER-').Update('0/%i' % self.frames)
                    self.canvas.config(width=self.vid_width, height=self.vid_height)

                    self.window.Element('-LEFT-').Update(range=(0, int(self.vid.width)), value=0)
                    self.window.Element('-RIGHT-').Update(range=(0, int(self.vid.width)), value=int(self.vid.width))
                    self.window.Element('-TOP-').Update(range=(0, int(self.vid.height)), value=0)
                    self.window.Element('-BOTTOM-').Update(range=(0, int(self.vid.height)), value=int(self.vid.height))


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

            if event == '-FRAMES-':
                self.set_frame(int(values['-FRAMES-']))

            if values['-LOWER-'] or values['-UPPER-']:
                if not self.play:
                        self.set_frame(self.frame)

            if event == '-MASK-':
                if not self.extended_properties['mask']:
                    self.extended_properties['mask'] = True
                    self.window.Element('-MASK-').Update('Unmask')
                    if not self.play:
                        self.set_frame(self.frame)
                else:
                    self.extended_properties['mask'] = False
                    self.window.Element('-MASK-').Update('Mask')
                    if not self.play:
                        self.set_frame(self.frame)

            if event == '-GET_CHANNEL-':
                if not self.extended_properties['get_channel']:
                    self.extended_properties['get_channel'] = True
                    if not self.play:
                        self.set_frame(self.frame)
                else:
                    self.extended_properties['get_channel'] = False
                    if not self.play:
                        self.set_frame(self.frame)
            
            if self.extended_properties['get_channel']:
                if values['-LEFT-']:
                    self.extended_properties['channel']['x1'] = int(values['-LEFT-'])
                if values['-RIGHT-']:
                    self.extended_properties['channel']['x2'] = int(values['-RIGHT-'])
                if values['-TOP-']:
                    self.extended_properties['channel']['y1'] = int(values['-TOP-'])
                if values['-BOTTOM-']:
                    self.extended_properties['channel']['y2'] = int(values['-BOTTOM-'])

            if event == '-PROCESSING_VIDEO-' and values['-INT_T-'] == True:
                GraphTime(self.extended_properties, video_path)
            if event == '-PROCESSING_VIDEO-' and values['-INT_W-'] == True:
                GraphWidth(self.extended_properties, video_path, values['-TIME-'])
            if event == '-PROCESSING_VIDEO-' and values['-3D_INT-'] == True:
                Graph3D(self.extended_properties, video_path)

            self.extended_properties['lower_color'] = int(values['-LOWER-'])
            self.extended_properties['upper_color'] = int(values['-UPPER-'])

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

                ret, frame = self.vid.get_frame()

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

            ret, frame = self.vid.goto_frame(frame_no)
            self.frame = frame_no
            self.update_counter(self.frame)

            if ret:
                self.photo = PIL.ImageTk.PhotoImage(
                    image=PIL.Image.fromarray(frame).resize((self.vid_width, self.vid_height), Image.NEAREST)
                )
                self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

    def update_counter(self, frame):

        self.window.Element('-FRAMES-').Update(value=frame)
        self.window.Element('-FRAMES_COUNTER-').Update('{}/{}'.format(frame, self.frames))


class MyVideoCapture:

    def __init__(self, video_source, extended_properties):
        
        self.vid = cv2.VideoCapture(video_source)
        if not self.vid.isOpened():
            raise ValueError('Unable to open video source', video_source)

        self.width = self.vid.get(cv2.CAP_PROP_FRAME_WIDTH)
        self.height = self.vid.get(cv2.CAP_PROP_FRAME_HEIGHT)
        self.frames = self.vid.get(cv2.CAP_PROP_FRAME_COUNT)
        self.fps = self.vid.get(cv2.CAP_PROP_FPS)

        self.extended_properties = extended_properties

    def get_frame(self):

        x1 = self.extended_properties['channel']['x1']
        x2 = self.extended_properties['channel']['x2']
        y1 = self.extended_properties['channel']['y1']
        y2 = self.extended_properties['channel']['y2']

        if self.vid.isOpened():
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                if self.extended_properties['mask']:
                    mask = cv2.inRange(frame, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                if self.extended_properties['get_channel']:
                    cv2.rectangle(frame, (x1, y1), (x2, y2), 255, 1)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None

    def goto_frame(self, frame_no):

        if self.vid.isOpened():
            self.vid.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                if self.extended_properties['mask']:
                    mask = cv2.inRange(frame, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None


    def __del__(self):
        if self.vid.isOpened():
            self.vid.release()

class GraphTime:
    def __init__(self, channel, video_source):
        self.count_frame = 0
        self.x1 = channel['channel']['x1']
        self.x2 = channel['channel']['x2']
        self.y1 = channel['channel']['y1']
        self.y2 = channel['channel']['y2']
        self.lower_color = channel['lower_color']
        self.upper_color = channel['upper_color']
        self.vid = cv2.VideoCapture(video_source)
        self.data_graph = {
            'time': None,
            'width': None,
            'intensity_time': [],
            'intensity_width': None
        }

        while True:

            ret, frame = self.vid.read()

            if ret == False:
                self.data_graph['time'] = np.arange(1, self.count_frame + 1, 1)
                cv2.destroyWindow('Video')
                break

            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            frame = frame[self.y1 : self.y2, self.x1 : self.x2]
            mask = cv2.inRange(frame, self.lower_color, self.upper_color)
            frame = cv2.bitwise_and(frame, frame, mask=mask)

            self.data_graph['intensity_time'].append(np.mean(frame))
            self.count_frame += 1

            cv2.imshow('Video', frame)

            if cv2.waitKey(1) & 0xFF == 27:
                cv2.destroyWindow('Video')
                break

        plt.plot(self.data_graph['time'], self.data_graph['intensity_time'])
        plt.xlabel('Time')
        plt.ylabel('Intensity')
        plt.grid(True)
        plt.legend(['Mean pixels intensity'])
        plt.show()

class GraphWidth:
    def __init__(self, channel, video_source, time):
        self.count_frame = 0
        self.x1 = channel['channel']['x1']
        self.x2 = channel['channel']['x2']
        self.y1 = channel['channel']['y1']
        self.y2 = channel['channel']['y2']
        self.lower_color = channel['lower_color']
        self.upper_color = channel['upper_color']
        self.vid = cv2.VideoCapture(video_source)
        self.data_graph = {
            'time': None,
            'width': None,
            'intensity_time': [],
            'intensity_width': []
        }

        while True:

            ret, frame = self.vid.read()

            if ret == False:
                self.data_graph['time'] = np.arange(1, self.count_frame + 1, 1)
                cv2.destroyWindow('Video')
                break

            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            frame = frame[self.y1 : self.y2, self.x1 : self.x2]
            mask = cv2.inRange(frame, self.lower_color, self.upper_color)
            frame = cv2.bitwise_and(frame, frame, mask=mask)

            self.count_frame += 1

            cv2.imshow('Video', frame)

            if  self.count_frame == int(time):
                self.data_graph['intensity_width'] = np.mean(frame, axis=0)
                self.data_graph['width'] = np.arange(1, np.size(frame, 1) + 1, 1)
                cv2.destroyWindow('Video')
                break
            if cv2.waitKey(1) & 0xFF == 27:
                cv2.destroyWindow('Video')
                break

        plt.plot(self.data_graph['width'], self.data_graph['intensity_width'])
        plt.xlabel('Time')
        plt.ylabel('Intensity')
        plt.grid(True)
        plt.legend(['Mean pixels intensity'])
        plt.show()

class Graph3D:
    def __init__(self, channel, video_source):
        self.count_frame = 0
        self.x1 = channel['channel']['x1']
        self.x2 = channel['channel']['x2']
        self.y1 = channel['channel']['y1']
        self.y2 = channel['channel']['y2']
        self.lower_color = channel['lower_color']
        self.upper_color = channel['upper_color']
        self.vid = cv2.VideoCapture(video_source)
        self.data_graph = {
            'time': None,
            'width': None,
            'intensity_time': [],
            'intensity_width': []
        }

        while True:

            ret, frame = self.vid.read()

            if ret == False:
                self.data_graph['time'] = np.arange(1, self.count_frame + 1, 1)
                cv2.destroyWindow('Video')
                break

            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            frame = frame[self.y1 : self.y2, self.x1 : self.x2]
            mask = cv2.inRange(frame, self.lower_color, self.upper_color)
            frame = cv2.bitwise_and(frame, frame, mask=mask)

            self.data_graph['intensity_time'].append(np.mean(frame))
            self.data_graph['intensity_width'].append(np.mean(frame, axis=0))
            self.data_graph['width'] = np.arange(1, np.size(frame, 1) + 1, 1)
            self.count_frame += 1

            cv2.imshow('Video', frame)

            if cv2.waitKey(1) & 0xFF == 27:
                cv2.destroyWindow('Video')
                break

        x = self.data_graph['width']
        y = self.data_graph['time']
        z = self.data_graph['intensity_width']
        z = np.array(z)

        x = x[::5]
        y = y[::5]
        z = z[::5, ::5]

        fig = plt.figure()
        ax = plt.axes(projection='3d')

        x, y = np.meshgrid(x, y)

        surf = ax.plot_surface(x, y, z, cmap=cm.Reds, rstride=1, cstride=1, linewidth=0, antialiased=False)

        ax.set_zlim(0, 255)
        ax.zaxis.set_major_locator(LinearLocator(10))
        ax.zaxis.set_major_formatter(FormatStrFormatter('%.00f'))

        fig.colorbar(surf, shrink=0.5, aspect=5)

        plt.show()



if __name__=='__main__':
    VideoPlayer()