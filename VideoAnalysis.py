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
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Text('Select video', key='-INSTRUCTION-')],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Input(key='-FILEPATH_VIDEO-'), 
            sg.Button('Browse', key='-BROWSE_VIDEO-')],
            [sg.Canvas(size=(700, 200), key='-CANVAS_VIDEO-', background_color='white', border_width=1)],
            [sg.Slider(size=(30, 20), range=(0, 100), resolution=1, key='-FRAMES-', orientation='h', 
            enable_events=True), 
            sg.T('0', key='-FRAMES_COUNTER-')],
            [sg.Button('Next frame', size=(8, 1)), 
            sg.Button('Pause', size=(8, 1), key='-PLAY-'), 
            sg.Button('Mask', size=(8, 1), key='-MASK-')],
            [sg.Text('Lower limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 0, resolution=1, key='-LOWER-', orientation='h', enable_events=True), 
            sg.Text('Upper limit:'),
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 255, resolution=1, key='-UPPER-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL-', font='Helvetica 16')],
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
            [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_VIDEO-', font='Helvetica 16')]]

#Tab for input and processing image
        image_layout = [
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Text('Select video', key='-INSTRUCTION-')],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Input(key='-FILEPATH_IMAGE-'), 
            sg.Button('Browse', key='-BROWSE_IMAGE-')],
            [sg.Canvas(size=(700, 200), key='-CANVAS_IMAGE-', background_color='white', border_width=1)],
            [sg.Button('Mask', size=(8, 1), key='-MASK_IMAGE-', font='Helvetica 16')],
            [sg.Text('Lower limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 0, resolution=1, key='-LOWER_IMAGE-', orientation='h', enable_events=True), 
            sg.Text('Upper limit:'), 
            sg.Slider(size=(20, 15), range=(0, 255), default_value = 255, resolution=1, key='-UPPER_IMAGE-', orientation='h', enable_events=True)],
            [sg.HorizontalSeparator(color = 'white')],
            [sg.Button('Get channel', enable_events=True, key='-GET_CHANNEL_IMAGE-', font='Helvetica 16')],
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
            [sg.Button('Convert video to graph', enable_events=True, key='-PROCESSING_IMAGE-', font='Helvetica 16')]]

#Tab for output data
        output_layout = [[sg.Canvas(size=(700, 200), key='-OUTPUT_CANVAS-', background_color='white', border_width=1)],
                        [sg.T(f'Area: 0', key='-AREA-', font='Helvetica 16')],
                        [sg.Button("Save as '.xlsx' file", enable_events=True, key='-SAVE_OUTPUT_FILE-', font='Helvetica 16')]]

#Layout for all tabs
        layout = [
            [sg.TabGroup([
                [sg.Tab('Video', video_layout, element_justification='center'),
                 sg.Tab('Image', image_layout, element_justification='center'), 
                 sg.Tab('Output', output_layout, element_justification='center')]
                ], enable_events=True, key='-APP-')],
            [sg.Button('Exit', font='Helvetica 16')]]

#Create main videoplayer's window
        screen_width, screen_height = sg.Window.get_screen_size()
        self.window = sg.Window('Signal intensity analysis', layout, size = (int(screen_width * 0.8), int(screen_height * 0.9)), resizable=True, element_justification='center').Finalize()

        canvas_video = self.window.Element('-CANVAS_VIDEO-')
        self.canvas_video = canvas_video.TKCanvas

        self.output_canvas = self.window.Element('-OUTPUT_CANVAS-')

        VideoPlayer(self.window, self.canvas_video, self.output_canvas)

#Base class for processing input file
class FileHandler:

    def __init__(self, window, canvas, output_canvas):
        self.window = window
        self.canvas = canvas
        self.output_canvas = output_canvas

        self.mask = False
        self.get_channel = False

        self.extended_properties = {
            'lower_color': 00,
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

            if event == '-PROCESSING_VIDEO-':
                GraphVideo(self.extended_properties, video_path, self.output_canvas, self.window, values['-TIME-'])

            self.extended_properties['lower_color'] = int(values['-LOWER-'])
            self.extended_properties['upper_color'] = int(values['-UPPER-'])

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
                    video_path = sg.filedialog.askopenfile().name
                except AttributeError:
                    print('no video selected, doing nothing')

                if video_path:
                    self.window.Element('-INSTRUCTION-').Update('Get channel')

                    self.instruction_text = 'Get channel'

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
                    GraphImage(self.extended_properties, video_path, self.output_canvas, self.window)

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

        img = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY) 
        if self.extended_properties['mask']:
            mask = cv2.inRange(img, self.extended_properties['lower_color'], self.extended_properties['upper_color'])
            img = cv2.bitwise_and(img, img, mask=mask)

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

#Class for create graph from video
class GraphVideo:
    def __init__(self, channel, source, output, window, time):
        Graph.__init__(self, channel, source, output, window)
        self.count_frame = 0
        self.vid = cv2.VideoCapture(source)
        
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
                self.data_table = {
                    'x': self.data_graph['width'],
                    'y': self.data_graph['intensity_width']
                }
                cv2.destroyWindow('Video')
                break
            if cv2.waitKey(1) & 0xFF == 27:
                cv2.destroyWindow('Video')
                break
        
        self.area = round(trapz(self.data_graph['intensity_width'], dx=1), 1)

        self.window.Element('-AREA-').Update(f'Area: {self.area}')

        def draw_figure(canvas, figure):
            figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
            figure_canvas_agg.draw()
            figure_canvas_agg.get_tk_widget().pack(side="top", fill="both", expand=1)
            return figure_canvas_agg
        
        fig = plt.figure()
        draw_figure(self.output.TKCanvas, fig)

        plt.plot(self.data_graph['width'], self.data_graph['intensity_width'])
        plt.xlabel('Distance')
        plt.ylabel('Intensity signal')
        plt.grid(True)
        plt.legend(['Mean pixels intensity'])
        plt.savefig('graph.png')

        OutputFile(self.area, self.window, self.data_table)

#Class for create graph from image
class GraphImage:
    def __init__(self, channel, source, output, window):
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
        
        self.area = round(trapz(self.data_graph['intensity_width'], dx=1), 1)

        window.Element('-AREA-').Update(f'Area: {self.area}')

        def draw_figure(canvas, figure):
            figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
            figure_canvas_agg.draw()
            figure_canvas_agg.get_tk_widget().pack(side="top", fill="both", expand=1)
            return figure_canvas_agg
        
        fig = plt.figure()
        draw_figure(output.TKCanvas, fig)

        plt.plot(self.data_graph['width'], self.data_graph['intensity_width'])
        plt.xlabel('Distance')
        plt.ylabel('Intensity signal')
        plt.grid(True)
        plt.legend(['Mean pixels intensity'])
        plt.savefig('graph.png')

        OutputFile(self.area, window, self.data_table)

#Class for create output 'xlsx' file
class OutputFile():
    def __init__(self, output_data, window, table):
        pass
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

            wb['Intensity signal']['A5'] = 'Distance'
            wb['Intensity signal']['B5'] = 'Signal Intensity'

            for i in range(0, len(table['x'])):
                wb['Intensity signal'][f'A{i+6}'] = table['x'][i]
                wb['Intensity signal'][f'B{i+6}'] = table['y'][i]

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
                break
        



if __name__=='__main__':
    App()