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
from scipy.integrate import simps
from numpy import trapz
from PIL import Image
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

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

#Main menu structure 
        layout = [
             [sg.HorizontalSeparator(color = 'white')],
             [sg.Text('Select video', key='-INSTRUCTION-')],
             [sg.HorizontalSeparator(color = 'white')],
             [sg.Input(key='-FILEPATH-'), sg.Button('Browse')],
             [sg.Canvas(size=(700, 200), key='-CANVAS-', background_color='white', border_width=1)],
             [sg.Slider(size=(30, 20), range=(0, 100), resolution=1, key='-FRAMES-', orientation='h', 
             enable_events=True), sg.T('0', key='-FRAMES_COUNTER-')],
             [sg.Button('Next frame', size=(8, 1)), sg.Button('Pause', size=(8, 1), key='Play'), sg.Button('Mask', size=(8, 1), key='-MASK-'),
             sg.Button('Exit', size=(8, 1))],
             [sg.Text('Lower limit:'), sg.Slider(size=(20, 15), range=(0, 255), default_value = 0, resolution=1, key='-LOWER-', orientation='h', enable_events=True), 
             sg.Text('Upper limit:'), sg.Slider(size=(20, 15), range=(0, 255), default_value = 255, resolution=1, key='-UPPER-', orientation='h', enable_events=True)],
             [sg.HorizontalSeparator(color = 'white')],
             [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL-', font='Helvetica 16')],
             [sg.Text('Left:'), sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-LEFT-', orientation='h', enable_events=True),
             sg.Text('Right:'),sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-RIGHT-', orientation='h', enable_events=True)],
             [sg.Text('Top:'), sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-TOP-', orientation='h', enable_events=True),
             sg.Text('Bottom:'), sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-BOTTOM-', orientation='h', enable_events=True)],
             [sg.HorizontalSeparator(color = 'white')],
             [sg.Radio('2D: Intensity/time', 'SELECT_TYPE_OF_GRAPH',  key='-INT_T-', default=True),
             sg.VSeperator(),
             sg.Radio('2D: Intensity/width', 'SELECT_TYPE_OF_GRAPH', key='-INT_W-'),
             sg.InputText(size=(5, 10), key='-TIME-', default_text = "10"), 
             sg.Text('ms'),
             sg.VSeperator(),
             sg.Radio('3D: Intensity/time/width', 'SELECT_TYPE_OF_GRAPH',  key='-3D_INT-')],
             [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_VIDEO-', font='Helvetica 16')]]

#Create main videoplayer's window
        self.window = sg.Window('Videoplayer', layout, element_justification='c').Finalize()

        canvas = self.window.Element('-CANVAS-')
        self.canvas = canvas.TKCanvas

        self.load_video()

#Main cycle for video processing
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
                    self.window.Element('-INSTRUCTION-').Update('Get channel')

                    self.instruction_text = 'Get channel'

                    self.vid = MyVideoCapture(video_path, self.extended_properties)

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
                    self.window.Element('-INSTRUCTION-').Update('Select type of graph and press button "Convert video to graph"')
                    if not self.play:
                        self.set_frame(self.frame)
                else:
                    self.extended_properties['get_channel'] = False
                    if not self.play:
                        self.set_frame(self.frame)
            
            if self.extended_properties['get_channel']:
                if values['-LEFT-'] or event == '-LEFT-':
                    self.extended_properties['channel']['x1'] = int(values['-LEFT-'])
                    if not self.play:
                        self.set_frame(self.frame)
                if values['-RIGHT-'] or event == '-RIGHT-':
                    self.extended_properties['channel']['x2'] = int(values['-RIGHT-'])
                    if not self.play:
                        self.set_frame(self.frame)
                if values['-TOP-']:
                    self.extended_properties['channel']['y1'] = int(values['-TOP-'])
                    if not self.play:
                        self.set_frame(self.frame)
                if values['-BOTTOM-']:
                    self.extended_properties['channel']['y2'] = int(values['-BOTTOM-'])
                    if not self.play:
                        self.set_frame(self.frame)

            if event == '-PROCESSING_VIDEO-' and values['-INT_W-'] == True:
                Graph(self.extended_properties, video_path, values['-TIME-'])

            self.extended_properties['lower_color'] = int(values['-LOWER-'])
            self.extended_properties['upper_color'] = int(values['-UPPER-'])

        self.window.close()
        sys.exit

#Function multithreading
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

#Get and change frame function
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
                    cv2.rectangle(frame, (x1, y1), (x2, y2), 255, 3)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None

    def goto_frame(self, frame_no):

        x1 = self.extended_properties['channel']['x1']
        x2 = self.extended_properties['channel']['x2']
        y1 = self.extended_properties['channel']['y1']
        y2 = self.extended_properties['channel']['y2']

        if self.vid.isOpened():
            self.vid.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                if self.extended_properties['mask']:
                    mask = cv2.inRange(frame, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                if self.extended_properties['get_channel']:
                    cv2.rectangle(frame, (x1, y1), (x2, y2), 255, 3)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None


    def __del__(self):
        if self.vid.isOpened():
            self.vid.release()

class Graph:
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
        
        area = round(trapz(self.data_graph['intensity_width'], dx=1), 1)

        plt.plot(self.data_graph['width'], self.data_graph['intensity_width'])
        plt.xlabel('Distance')
        plt.ylabel('Intensity signal')
        plt.grid(True)
        plt.legend(['Mean pixels intensity'])
        plt.savefig('graph.png')
        plt.show()
        OutputFile(area)

class OutputFile():
    def __init__(self, output_data):
        def create_data_style():
            ns = NamedStyle(name='highlight')
            ns.font = Font(bold=True, size=18)
            border = Side(style='thin', color='000000')
            ns.border = Border(left=border, top=border, right=border, bottom=border)
            ns.alignment = Alignment(horizontal="center", vertical="center")
            wb.add_named_style(ns)

        def insert_graph(wb):
            wb.create_sheet(title = 'Intensity signal', index = 0)

            create_data_style()

            wb['Intensity signal'].column_dimensions['B'].width = 30
            wb['Intensity signal']['B2'].style = 'highlight'
            wb['Intensity signal']['B3'].style = 'highlight'

            img = openpyxl.drawing.image.Image('graph.png')
            img.anchor = 'D2'

            wb['Intensity signal'].add_image(img)
            wb['Intensity signal']['B2'] = 'Area'
            wb['Intensity signal']['B3'] = '{0:,}'.format(output_data).replace(',', ' ')

        wb = Workbook()

        insert_graph(wb)

        wb.save('test.xlsx')
        



if __name__=='__main__':
    VideoPlayer()