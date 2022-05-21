import cv2
import numpy as np
import PySimpleGUI as sg
import PIL 
from PIL import Image
import tkinter as tk
import threading
import time
import matplotlib.pyplot as plt
from numpy import trapz
from PIL import Image
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

#Main class of application
class App:

    def __init__(self):
    
#Tab for input and processing video
        video_layout = [
            [sg.Input(key='-FILEPATH_VIDEO-'), 
            sg.Button('Browse', key='-BROWSE_VIDEO-')],
            [sg.Canvas(size=(700, 200), key='-CANVAS_VIDEO-', background_color='white', border_width=1)],
            [sg.Slider(size=(30, 20), range=(0, 100), resolution=1, key='-FRAMES-', orientation='h', 
            enable_events=True), 
            sg.T('0', key='-FRAMES_COUNTER-')],
            [sg.Button('Next frame', size=(8, 1)), 
            sg.Button('Pause', size=(8, 1), key='-PLAY-'), 
            sg.Button('Mask', size=(8, 1), key='-MASK-'),
            sg.Button('Blure', size=(8, 1), key='-BLURE_VIDEO-'),
            sg.Text('kX:'), 
            sg.Spin(values=[i for i in range(1,50,2)], initial_value=1, key='-BLURE_VID_VALUE-', enable_events=True),
            sg.Text('k:'), 
            sg.Spin(values=[i for i in range(0,20)], initial_value=0, key='-BLURE_VID_VALUE_K-', enable_events=True)],
            [sg.Text('Lower limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 0, resolution=1, key='-LOWER-', orientation='h', enable_events=True), 
            sg.Text('Upper limit:'),
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 255, resolution=1, key='-UPPER-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL-')],
            [sg.Text('Left:'),
            sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-LEFT-', orientation='h', enable_events=True),
            sg.Text('Right:'),
            sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-RIGHT-', orientation='h', enable_events=True)],
            [sg.Text('Top:'),
            sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-TOP-', orientation='h', enable_events=True),
            sg.Text('Bottom:'), 
            sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-BOTTOM-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Radio('2D: Intensity/width', '-GRAPH_VIDEO-', default=True),
            sg.InputText(size=(5, 10), key='-TIME-', default_text = "30"), 
            sg.Text('frames')],
            [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_VIDEO-')]]

#Tab for input and processing image
        image_layout = [
            [sg.Input(key='-FILEPATH_IMAGE-'), 
            sg.Button('Browse', key='-BROWSE_IMAGE-')],
            [sg.Canvas(size=(700, 200), key='-CANVAS_IMAGE-', background_color='white', border_width=1)],
            [sg.Button('Mask', size=(8, 1), key='-MASK_IMAGE-'),
            sg.Button('Blure', size=(8, 1), key='-BLURE_IMAGE-'),
            sg.Text('kX:'), 
            sg.Spin(values=[i for i in range(1,50,2)], initial_value=1, key='-BLURE_IMG_VALUE-', enable_events=True), 
            sg.Spin(values=[i for i in range(0,20)], initial_value=0, key='-BLURE_IMG_VALUE_K-', enable_events=True)],
            [sg.Text('Lower limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 0, resolution=1, key='-LOWER_IMAGE-', orientation='h', enable_events=True), 
            sg.Text('Upper limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 255, resolution=1, key='-UPPER_IMAGE-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL_IMAGE-')],
            [sg.Text('Left:'), 
            sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-LEFT_IMAGE-', orientation='h', enable_events=True),
            sg.Text('Right:'),
            sg.Slider(size=(20, 15), range=(0, 100), resolution=1, key='-RIGHT_IMAGE-', orientation='h', enable_events=True)],
            [sg.Text('Top:'), 
            sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-TOP_IMAGE-', orientation='h', enable_events=True),
            sg.Text('Bottom:'),
            sg.Slider(size=(20, 15), range=(0, 100), default_value = 0, resolution=1, key='-BOTTOM_IMAGE-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Radio('2D: Intensity/width', '-GRAPH_IMAGE-', default=True)],
            [sg.Button('Convert image to graph', enable_events=True, key='-PROCESSING_IMAGE-')]]

#Tab for output data
        output_layout = [[sg.Canvas(size=(700, 200), key='-OUTPUT_CANVAS-', background_color='white', border_width=1)],
                        [sg.Column([[sg.T('0', key='-AREA-')]], scrollable=True, vertical_scroll_only=True, justification='center', size=(100, 10), key='-AREA_LIST-')],
                        [sg.Button("Save as '.xlsx' file", enable_events=True, key='-SAVE_OUTPUT_FILE-')]]

#Layout for all tabs
        layout = [
            [sg.TabGroup([
                [sg.Tab('Video', video_layout, element_justification='center'),
                 sg.Tab('Image', image_layout, element_justification='center'), 
                 sg.Tab('Output', output_layout, element_justification='center')]
                ], enable_events=True, key='-APP-')],
            [sg.Button('Exit')]]

#Create main videoplayer's window
        screen_width, screen_height = sg.Window.get_screen_size()
        self.window = sg.Window('Signal intensity analysis', layout, size = (int(screen_width * 0.8), int(screen_height * 0.9)), resizable=True, element_justification='center').Finalize()

        canvas_video = self.window.Element('-CANVAS_VIDEO-')
        self.canvas_video = canvas_video.TKCanvas

        self.output_canvas = self.window.Element('-OUTPUT_CANVAS-').TKCanvas

        VideoPlayer(self.window, self.canvas_video, self.output_canvas)

#Base class for processing input file
class FileHandler:

    def __init__(self, window, canvas, output_canvas):
        self.window = window
        self.canvas = canvas
        self.output_canvas = output_canvas

        self.mask = False
        self.get_channel = False
        self.blure = False

        self.extended_properties = {
            'lower_color': 00,
            'upper_color': 255,
            'blure_value': 1,
            'blure_value_k': 1,
            'mask': False,
            'get_channel': False,
            'blure': False,
            'channel': {
                'x1': 1,
                'x2': 1,
                'y1': 1,
                'y2': 1
            }
        }

#Class for processing video file
class VideoPlayer(FileHandler):

    def __init__(self, window, canvas, output_canvas):

        FileHandler.__init__(self, window, canvas, output_canvas)

        self.play = True
        self.delay = 0.023
        self.frame = 1
        self.frames = None
        self.get_channel = False
        self.vid = None
        self.photo = None
        self.next = '1'
        
        self.load_video()

        canvas_image = self.window.Element('-CANVAS_IMAGE-')
        canvas_image = canvas_image.TKCanvas

        while True:
            event, values = self.window.Read()
            if event is None or event == 'Exit':
                break
            if values['-APP-'] == 'Image':
                ImageEditor(window, canvas_image, output_canvas)
            if event == '-BROWSE_VIDEO-':
                video_path = None
                try:
                    video_path = sg.filedialog.askopenfile(filetypes=[("Video", ".MP4 .AVI")]).name
                except AttributeError:
                    print('no video selected, doing nothing')

                if video_path:

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

                    self.window.Element('-FILEPATH_VIDEO-').Update(video_path)

            if event == '-PLAY-':
                if self.play:
                    self.play = False
                    self.window.Element('-PLAY-').Update('Play')
                else:
                    self.play = True
                    self.window.Element('-PLAY-').Update('Pause')

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

            if event == '-BLURE_VIDEO-':
                if not self.extended_properties['blure']:
                    self.extended_properties['blure'] = True
                    self.window.Element('-BLURE_VIDEO-').Update('Unblure')
                    if not self.play:
                        self.set_frame(self.frame)
                else:
                    self.extended_properties['blure'] = False
                    self.window.Element('-BLURE_VIDEO-').Update('Blure')
                    if not self.play:
                        self.set_frame(self.frame)

            if event == '-BLURE_VID_VALUE-':
                if not self.play:
                        self.set_frame(self.frame)

            if event == '-BLURE_VID_VALUE_K-':
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

            if event == '-PROCESSING_VIDEO-':
                if self.extended_properties['get_channel']:
                    GraphVideo(self.extended_properties, video_path, self.output_canvas, self.window, values['-TIME-'], canvas, output_canvas)
                else:
                    sg.popup_ok('No channel selected')

            self.extended_properties['lower_color'] = int(values['-LOWER-'])
            self.extended_properties['upper_color'] = int(values['-UPPER-'])
            self.extended_properties['blure_value'] = values['-BLURE_VID_VALUE-']
            self.extended_properties['blure_value_k'] = values['-BLURE_VID_VALUE_K-']

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

#Class for create and managment video
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

        blure_value = self.extended_properties['blure_value']
        blure_value_k = self.extended_properties['blure_value_k']

        if self.vid.isOpened():
            ret, frame = self.vid.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) 
                if self.extended_properties['mask']:
                    mask = cv2.inRange(frame, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
                    frame = cv2.bitwise_and(frame, frame, mask=mask)
                if self.extended_properties['get_channel']:
                    cv2.rectangle(frame, (x1, y1), (x2, y2), 255, 3)
                if self.extended_properties['blure']:
                    frame = cv2.GaussianBlur(frame, (blure_value, blure_value), blure_value_k)
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

        blure_value = self.extended_properties['blure_value']
        blure_value_k = self.extended_properties['blure_value_k']

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
                if self.extended_properties['blure']:
                    frame = cv2.GaussianBlur(frame, (blure_value, blure_value), blure_value_k)
                return ret, frame
            else:
                return ret, None
        else:
            return 0, None


    def __del__(self):
        if self.vid.isOpened():
            self.vid.release()

#Class for processing image
class ImageEditor(FileHandler):
    def __init__(self, window, canvas, output_canvas):
        FileHandler.__init__(self, window, canvas, output_canvas)

        canvas_video = self.window.Element('-CANVAS_VIDEO-')
        canvas_video = canvas_video.TKCanvas

        while True:
            event, values = self.window.Read()

            if event is None or event == 'Exit':
                break
            if values['-APP-'] == 'Video':
                VideoPlayer(window, canvas_video, output_canvas)
            if event == '-BROWSE_IMAGE-':
                video_path = None
                try:
                    video_path = sg.filedialog.askopenfile(filetypes=[("Image", ".png .jpg .tif .jpeg")]).name
                except AttributeError:
                    print('no video selected, doing nothing')

                if video_path:

                    self.img = MyImageCapture(video_path, self.extended_properties)

                    self.img_width = int(self.img.width / self.img.height * 200) 
                    self.img_height = 200

                    self.canvas.config(width=self.img_width, height=self.img_height)

                    self.window.Element('-LEFT_IMAGE-').Update(range=(0, int(self.img.width)), value=0)
                    self.window.Element('-RIGHT_IMAGE-').Update(range=(0, int(self.img.width)), value=int(self.img.width))
                    self.window.Element('-TOP_IMAGE-').Update(range=(0, int(self.img.height)), value=0)
                    self.window.Element('-BOTTOM_IMAGE-').Update(range=(0, int(self.img.height)), value=int(self.img.height))

                    self.window.Element('-FILEPATH_IMAGE-').Update(video_path)

                    self.photo = PIL.ImageTk.PhotoImage(
                        image=PIL.Image.fromarray(self.img.img).resize((self.img_width, self.img_height), Image.NEAREST)
                    )
                    self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

            if event == '-MASK_IMAGE-':
                if not self.extended_properties['mask']:
                    self.extended_properties['mask'] = True
                    self.window.Element('-MASK_IMAGE-').Update('Unmask')
                    self.update()
                else:
                    self.extended_properties['mask'] = False
                    self.window.Element('-MASK_IMAGE-').Update('Mask')
                    self.update()
            
            if event == '-LOWER_IMAGE-' or event == '-UPPER_IMAGE-':
                self.extended_properties['lower_color'] = int(values['-LOWER_IMAGE-'])
                self.extended_properties['upper_color'] = int(values['-UPPER_IMAGE-'])
                self.update()

            
            if event == '-BLURE_IMAGE-':
                if not self.extended_properties['blure']:
                    self.extended_properties['blure'] = True
                    self.window.Element('-BLURE_IMAGE-').Update('Unblure')
                    self.update()
                else:
                    self.extended_properties['blure'] = False
                    self.window.Element('-BLURE_IMAGE-').Update('Blure')
                    self.update()

            if event == '-BLURE_IMG_VALUE-':
                self.extended_properties['blure_value'] = values['-BLURE_IMG_VALUE-']
                self.update()

            if event == '-BLURE_IMG_VALUE_K-':
                self.extended_properties['blure_value_k'] = values['-BLURE_IMG_VALUE_K-']
                self.update()

            if event == '-GET_CHANNEL_IMAGE-':
                if not self.extended_properties['get_channel']:
                    self.extended_properties['get_channel'] = True
                    self.update()
                else:
                    self.extended_properties['get_channel'] = False
                    self.update()
    
            if self.extended_properties['get_channel']:
                if event == '-LEFT_IMAGE-' or values['-LEFT_IMAGE-']:
                    self.extended_properties['channel']['x1'] = int(values['-LEFT_IMAGE-'])
                    self.update()
                if event == '-RIGHT_IMAGE-' or values['-RIGHT_IMAGE-']:
                    self.extended_properties['channel']['x2'] = int(values['-RIGHT_IMAGE-'])
                    self.update()
                if event == '-TOP_IMAGE-' or values['-TOP_IMAGE-']:
                    self.extended_properties['channel']['y1'] = int(values['-TOP_IMAGE-'])
                    self.update()
                if event == '-BOTTOM_IMAGE-' or values['-BOTTOM_IMAGE-']:
                    self.extended_properties['channel']['y2'] = int(values['-BOTTOM_IMAGE-'])
                    self.update()

            if event == '-PROCESSING_IMAGE-':
                if self.extended_properties['get_channel']:
                    GraphImage(self.extended_properties, video_path, self.output_canvas, self.window, canvas, output_canvas)
                else:
                    sg.popup_ok('No channel selected')

    def update(self):
        if self.img:
            img = self.img.get_frame()
            self.photo = PIL.ImageTk.PhotoImage(
                image=PIL.Image.fromarray(img).resize((self.img_width, self.img_height), Image.NEAREST)
            )
            self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)
            
class MyImageCapture:

    def __init__(self, video_source, extended_properties):
        self.img = cv2.imdecode(np.fromfile(video_source, dtype=np.uint8), cv2.IMREAD_COLOR)

        self.width = self.img.shape[1]
        self.height = self.img.shape[0]

        self.extended_properties = extended_properties

    def get_frame(self):

        x1 = self.extended_properties['channel']['x1']
        x2 = self.extended_properties['channel']['x2']
        y1 = self.extended_properties['channel']['y1']
        y2 = self.extended_properties['channel']['y2']

        blure_value = self.extended_properties['blure_value']
        blure_value_k = self.extended_properties['blure_value_k']

        img = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY) 
        if self.extended_properties['mask']:
            mask = cv2.inRange(img, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
            img = cv2.bitwise_and(img, img, mask=mask)

        if self.extended_properties['blure']:
            img = cv2.GaussianBlur(img, (blure_value, blure_value), blure_value_k)

        if self.extended_properties['get_channel']:
            cv2.rectangle(img, (x1, y1), (x2, y2), 255, 3)
        return img

#Class for create graph
class Graph:
    def __init__(self, channel, source, output, window):
        self.window = window
        self.output = output
        self.source = source
        self.x1 = channel['channel']['x1']
        self.x2 = channel['channel']['x2']
        self.y1 = channel['channel']['y1']
        self.y2 = channel['channel']['y2']
        self.lower_color = channel['lower_color']
        self.upper_color = channel['upper_color']
        self.data_graph = {
            'width': None,
            'intensity_width': []
        }
        self.area = []

    def draw_graph(self, data_for_graph, data_x, data_y, label_x, label_y):
        fig, ax = plt.subplots()
        canvasGraph = FigureCanvasTkAgg(fig, data_for_graph)
        plot_widget = canvasGraph.get_tk_widget()
        plot_widget.grid(row=0, column=0)
        ax.cla()
        ax.set_xlabel(label_x)
        ax.set_ylabel(label_y)
        ax.grid()
        plt.plot(data_x, data_y)     
        fig.canvas.draw()  
        plt.savefig('graph.png')

#Class for create graph from video
class GraphVideo:
    def __init__(self, channel, source, output, window, time, canvas, output_canvas):
        Graph.__init__(self, channel, source, output, window)
        self.count_frame = 0
        self.vid = cv2.VideoCapture(source)
        
        def get_figures(x1, array_length):
            if x1 > array_length:
                return
            for i in range (x1, array_length):
                if self.data_graph['intensity_width'][i] > 0:
                    left_dot = i
                    for j in range(i, len(self.data_graph['intensity_width'])):
                        if self.data_graph['intensity_width'][j] == 0:
                            right_dot = j  - 1
                            array_figure.append((left_dot, right_dot))
                            get_figures(right_dot + 1, array_length)
                            return

        def calc_area(main_array, array_figure):
            area_array = []
            figures =[]
            for i in array_figure:
                left_dot, right_dot = i
                figures.append(main_array[left_dot : right_dot])
            for i in figures:
                area_array.append(round(trapz(i, dx=1), 1))
            return area_array

        def filter_area(area):
            if area > 100:
                return True
            else:
                return False

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

                array_figure = []
                get_figures(1, len(self.data_graph['intensity_width']))

                self.data_graph['width'] = np.arange(1, np.size(frame, 1) + 1, 1)
                self.data_table = {
                    'x': self.data_graph['width'],
                    'y': self.data_graph['intensity_width']
                }
                cv2.destroyWindow('Video')
                break
            if cv2.waitKey(1) & 0xFF == 27:
                cv2.destroyWindow('Video')
                break
        
        self.area = calc_area(self.data_graph['intensity_width'], array_figure)
        self.area = filter(filter_area, self.area)
        self.area = [*self.area]

        area_element = ''
        count = 1
        for i in range(0, len(self.area)):
            if self.area[i] > 100:
                area_element += f'Area{count}: {self.area[i]}\n'
                count += 1

        self.window.Element('-AREA-').Update(area_element)

        window.Element('-AREA-').Update(area_element)

        Graph.draw_graph(self, output, self.data_graph['width'],
                        self.data_graph['intensity_width'],
                        'Distance', 
                        'Signal intensity')

        OutputFile(self.area, self.window, self.data_table, canvas, output_canvas)

#Class for create graph from image
class GraphImage:
    def __init__(self, channel, source, output, window, canvas, output_canvas):
        Graph.__init__(self, channel, source, output, window)
        self.x1 = channel['channel']['x1']
        self.x2 = channel['channel']['x2']
        self.y1 = channel['channel']['y1']
        self.y2 = channel['channel']['y2']
        self.lower_color = channel['lower_color']
        self.upper_color = channel['upper_color']
        self.img = cv2.imdecode(np.fromfile(source, dtype=np.uint8), cv2.IMREAD_COLOR)
        self.data_graph = {
            'width': None,
            'intensity_width': []
        }

        def get_figures(x1, array_length):
            if x1 > array_length:
                return
            for i in range (x1, array_length):
                if self.data_graph['intensity_width'][i] > 0:
                    left_dot = i
                    for j in range(i, len(self.data_graph['intensity_width'])):
                        if self.data_graph['intensity_width'][j] == 0:
                            right_dot = j  - 1
                            array_figure.append((left_dot, right_dot))
                            get_figures(right_dot + 1, array_length)
                            return

        def calc_area(main_array, array_figure):
            area_array = []
            figures =[]
            for i in array_figure:
                left_dot, right_dot = i
                figures.append(main_array[left_dot : right_dot])
            for i in figures:
                area_array.append(round(trapz(i, dx=1), 1))
            return area_array

        def filter_area(area):
            if area > 100:
                return True
            else:
                return False


        img = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        img = img[self.y1 : self.y2, self.x1 : self.x2]
        mask = cv2.inRange(img, self.lower_color, self.upper_color)
        img = cv2.bitwise_and(img, img, mask=mask)

        self.data_graph['intensity_width'] = np.mean(img, axis=0)
        self.data_graph['width'] = np.arange(1, np.size(img, 1) + 1, 1)
        self.data_table = {
            'x': self.data_graph['width'],
            'y': self.data_graph['intensity_width']
        }       

        array_figure = []
        get_figures(1, len(self.data_graph['intensity_width']))
        
        self.area = calc_area(self.data_graph['intensity_width'], array_figure)
        self.area = filter(filter_area, self.area)
        self.area = [*self.area]

        area_element = ''
        count = 1
        for i in range(0, len(self.area)):
            area_element += f'Area{count}: {self.area[i]}\n'
            count += 1

        window.Element('-AREA-').Update(area_element)

        Graph.draw_graph(self, output, self.data_graph['width'],
                        self.data_graph['intensity_width'],
                        'Distance', 
                        'Signal intensity')

        OutputFile(self.area, window, self.data_table, canvas, output_canvas)

#Class for create output 'xlsx' file
class OutputFile():
    def __init__(self, output_data, window, table, canvas, output_canvas):
        def create_data_style(name, bold, font_size):
            ns = NamedStyle(name=name)
            ns.font = Font(bold=bold, size=font_size)
            border = Side(style='thin', color='000000')
            ns.border = Border(left=border, top=border, right=border, bottom=border)
            ns.alignment = Alignment(horizontal="center", vertical="center")
            wb.add_named_style(ns)

        def insert_graph(wb):
            wb.create_sheet(title = 'Intensity signal', index = 0)

            create_data_style('highlight', True, 18)
            create_data_style('table', False, 12)

            wb['Intensity signal'].column_dimensions['B'].width = 30

            img = openpyxl.drawing.image.Image('graph.png')
            img.anchor = 'D2'

            wb['Intensity signal'].add_image(img)

            for i in range(0, len(output_data)):
                wb['Intensity signal'][f'B{2 + 2 * i}'].style = 'highlight'
                wb['Intensity signal'][f'B{3 + 2 * i}'].style = 'highlight'
                wb['Intensity signal'][f'B{2 + 2 * i}'] = f'Area{i + 1}'
                wb['Intensity signal'][f'B{3 + 2 * i}'] = '{0:,}'.format(output_data[i]).replace(',', ' ')

            wb['Intensity signal'][f'A{5 + (len(output_data) - 1) * 2}'] = 'Distance'
            wb['Intensity signal'][f'B{5 + (len(output_data) - 1) * 2}'] = 'Signal Intensity'
            wb['Intensity signal'][f'A{5 + (len(output_data) - 1) * 2}'].style = 'table'
            wb['Intensity signal'][f'B{5 + (len(output_data) - 1) * 2}'].style = 'table'

            shift = 6 + ((len(output_data) - 1) * 2)
            for i in range(0, len(table['x'])):
                wb['Intensity signal'][f'A{i + shift}'] = table['x'][i]
                wb['Intensity signal'][f'B{i + shift}'] = table['y'][i]
                wb['Intensity signal'][f'A{i + shift}'].style = 'table'
                wb['Intensity signal'][f'B{i + shift}'].style = 'table'

        wb = Workbook()

        while True:
            event, values = window.Read()

            if event is None or event == 'Exit':
                    break

            if event == '-SAVE_OUTPUT_FILE-':
                insert_graph(wb)
                path = sg.filedialog.asksaveasfile().name
                wb.save(path+'.xlsx')
                os.remove('graph.png')
                os.remove(path)

            if values['-APP-'] == 'Image':
                canvas.delete('all')
                window.Element('-OUTPUT_CANVAS-').TKCanvas.delete('all')
                ImageEditor(window, canvas, output_canvas)
            if values['-APP-'] == 'Video':
                canvas.delete('all')
                VideoPlayer(window, canvas, output_canvas)
        
if __name__=='__main__':
    App()