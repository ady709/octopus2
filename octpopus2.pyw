import tkinter as tk
from tkinter import messagebox, ttk
from time import sleep, time
import win32com.client
from threading import Thread, Lock, Event
from queue import Queue, Empty
from time import sleep, strftime
import pythoncom
import os,sys
from datetime import datetime, timedelta
import pandas as pd

from sap_script import job, grouping, script_name, test_mode_supported


# region processor
class Processor(Thread):
    def __init__(self, sysnr:int, sesnr:int, id:int, own_running:Event, count_job, test_mode, controller):
        super().__init__(target=self.thread_controller, daemon=True)
        self.sysnr = int(sysnr)
        self.sesnr = int(sesnr)
        self.id = id
        self.own_running = own_running
        self.count_job = count_job # controller callback 
        self.test_mode = test_mode # controller callback 
        self.controller = controller
        
        self.q = controller.q
        self.report_q = controller.report_q
        self.lock = controller.lock
        self.input_file = controller.input_file

        self.outxl = pd.DataFrame()
        self.outxl_total = pd.DataFrame()
        self.reportxl = pd.DataFrame()
        self.reportxl_total = pd.DataFrame()

        with self.lock:
            self.state = 'not_started' # started, paused, dead, stopped, will_pause
        try:
            self.settings = self.input_file['Settings'][['Setting','Value']].fillna('').set_index('Setting').to_dict()['Value']
        except Exception:
            self.settings = {'Test Mode':'n'}
        self.start()

    def text_update(self, text_to_add, color='black'):
        self.report_q.put((self.id, text_to_add, color))

    def pause(self, message=' Paused! '):
        with self.lock:
            self.state = 'paused'
        self.own_running.clear()
        self.text_update('\n'+message, color='red')
        self.own_running.wait()
        with self.lock:
            self.state = 'running'

    def thread_controller(self): #################################################
        pythoncom.CoInitialize()
        try:
            self.session = win32com.client.GetObject('SAPGUI').GetScriptingEngine.Children(self.sysnr).Children(self.sesnr)
        except Exception:
            self.text_update(f'{self.sysnr}-{self.sesnr} No connection!')
            with self.lock:
                self.state = 'dead'
            return False
        
        self.text_update(f'{strftime("%H:%M")} Started\n')
        while self.q.qsize()>0: ##############
            try:
                self.session.Info.Transaction
            except Exception:
                with self.lock:
                    self.state = 'dead'
                self.text_update(f'{self.sysnr}-{self.sesnr} Connection lost!') 
                break
            # wait if paused
            with self.lock:
                self.state = 'paused'
            self.own_running.wait()
            with self.lock:
                self.state = 'running'
            
            try:##############  DO THE JOB  ###############################
                a=self.q.get(block=False)
            except Empty:
                break

            self.outxl = pd.DataFrame()
            self.outxl.index = a.index
            self.reportxl = pd.DataFrame()
            self.text_update(f'Starting job ')
            script_result = False
            color = 'black'
            result = False
            try:
                result = job(self, df=a)####
                if result == 'postpone':
                    self.q.task_done()
                    self.q.put(a)
                    self.text_update('Postponing\n')
                    script_result =' --postponing-- '
                    continue
                if result == True:
                    script_result = ' --OK-- '
                    color = 'green'
                else:  
                    script_result = ' --ERR-- '
                    color = 'red'
                    result = False
            except Exception:
                script_result = ' --ERR-- '
                color = 'red'
            #job result
            self.q.task_done()
            self.count_job(result)
            message = ''
            try:
                message = self.session.findById("wnd[0]/sbar").Text
            except Exception:
                pass
            self.text_update(script_result, color)
            if message:
                self.text_update(' ' + message)
            #put script result to all lines of job
            for idx in self.outxl.index:
                self.outxl.loc[idx, 'Script result'] = script_result
                if message:
                    self.outxl.loc[idx,'Status Bar Message'] = message
            #merge this job with outxlf of this processor
            self.outxl_total = pd.concat([self.outxl_total, self.outxl]).copy()
            self.reportxl_total = pd.concat([self.reportxl_total, self.reportxl], ignore_index=True).copy()
            self.text_update('\n')
        
        #end loop
        self.text_update(f'{strftime("%H:%M")} Finished!', color='green')
        self.session = None
        pythoncom.CoUninitialize()
        with self.lock:
            self.state = 'stopped'

# endregion



# region Controller
class Controller:
    def __init__(self, view, script, grouping=None, test_mode_supported=False):
        self.view = view
        self.script = script
        self.grouping = grouping
        
        self.state = 'not_started' # 'not_started', 'started', 'paused', 'stopping', 'stopped', 'finished'
        self.connection = None
        self.sap_sessions = dict()
        self.input_file = None

        self.q = Queue()
        self.report_q = Queue()
        self.jobs_done = 0
        self.jobs_done_error = 0
        self.lock = Lock()
        self.state='not_started'
        self.processors = dict()
        self.procesosors_running = dict()
        self.start_time = self.time_in_last_tick = datetime.now()
        self.elapsed_time = timedelta(seconds=0)
        self.initial_q_size = 0
        self.remaining_time = "unknown"
        self.remaining_time = timedelta(seconds=0)
        self.items_per_minute=0
        self.ticks_time_measure = 0
        self.final_q_size = None
        self.test_mode = False
        self.test_mode_supported = test_mode_supported

        #view callbacks
        self.view.controller_start_work = self.start_work
        self.view.controller_get_state = self.get_state
        self.view.controller_tick_scan_sessions = self.tick_scan_sessions
        self.view.controller_get_connection = self.get_connection
        self.view.controller_check_selected_sessions = self.check_selected_sessions
        self.view.controller_interrupt_work = self.interrupt_work
        self.view.controller_pause = self.pause
        self.view.controller_resume = self.resume
        self.view.controller_tick_work = self.tick_work
        self.view.controller_processor_running_togle = self.processor_running_togle
        self.view.controller_ask_exit = self.ask_exit
        self.view.controller_get_test_mode = self.get_test_mode
        self.view.controller_set_test_mode = self.set_test_mode
        self.view.controller_test_mode_supported = self.get_test_mode_supported

        self.readFile()
        self.sap_connect()
        self.view.start_scan_ticks()

    
    #Read/Load the input file
    def readFile(self):
        try:
            p = os.path.dirname(os.path.abspath(__file__))
            f = os.path.split(p)[1]
            if f != '_internal':
                os.chdir(os.path.dirname(os.path.abspath(__file__)))
            self.input_file = pd.read_excel('input.xlsx', sheet_name=None, dtype='str')
            self.input_file['Input'] = self.input_file['Input'].dropna(how='all').fillna('').reset_index()
            self.input_file['Input']['Script result'] = "Not started"
        except Exception:
            self.view.show_modal('error', 'Input file not found! Please create input.xlsx file in the same folder as this script and have it closed.')
            self.view.exit()
            sys.exit()
        #grouping of items as specified in sap_script
        if not self.grouping is None and len(self.grouping) > 0:
            self.input_file['Input'] = self.input_file['Input'].set_index(grouping)
            for i in self.input_file['Input'].index.drop_duplicates():
                job = self.input_file['Input'].loc[[i]].reset_index().copy()
                job.set_index(job['index'], drop=False, inplace=True)
                self.q.put(job)
            self.input_file['Input'] = self.input_file['Input'].reset_index()
            self.input_file['Input'] = self.input_file['Input'].set_index('index', drop=True)
        else:
            for i in self.input_file['Input'].index.tolist():
                self.q.put(self.input_file['Input'].loc[[i]])
        self.initial_q_size = self.q.qsize()

        self.view.display_job_count(self.initial_q_size)



    def sap_connect(self):
        #scan for open sap connections
        try:
            self.connection = win32com.client.GetObject('SAPGUI').GetScriptingEngine
            if not type(self.connection) is win32com.client.CDispatch or self.connection.Children.count == 0:
                self.connection = None
                raise Exception('No Sap connection!')
        except Exception:
            self.view.show_modal(type='error', message='No SAP!')
            self.view.exit()
            sys.exit()
        return
            

    def tick_scan_sessions(self):
        '''
        Periodically triggered action that scans for sap sessions
        '''
        # before main function is started, check for available sap sessions
        if self.state == 'not_started':
            self.scansapses()
            self.view.update_session_buttons(self.sap_sessions)

    #session scanner
    def scansapses(self):
        if self.connection is None:
            self.sap_connect()
        if self.connection is None:
            self.view.exit()
            sys.exit()
        self.sap_sessions = dict()
        for sysnr,sapsys in enumerate(self.connection.Children):
            for sesnr,ses in enumerate(sapsys.Children):
                if ses.busy:
                    continue
                self.sap_sessions[f'{str(sysnr)}.{str(sesnr)}'] = ses

    def check_selected_sessions(self, session_keys):
        systems = set()
        for s in session_keys:
            sys, ses = s.split('.')
            systems.add(sys)
        if len(systems) == 1:
            if len(session_keys) > self.initial_q_size:
                return (False, 'Select less sessions!')
            else:
                return (True, 'Select session(s)')
        if len(systems) > 1:
            return (False, 'Select only one system!')
        else:
            return (False, 'Select session(s)')
    
    def count_job(self, result):
        with self.lock:
            self.jobs_done += 1
            if not result:
                self.jobs_done_error += 1

    def get_test_mode(self):
        return self.test_mode
    
    def set_test_mode(self, test_mode):
        self.test_mode = test_mode

    def get_test_mode_supported(self):
        return self.test_mode_supported

    def start_work(self, session_keys):
        if not self.check_selected_sessions(session_keys)[0]:
            return
        #self.state = 'started'
        self.view.start_work_window()
        #make processors
        for id, ses in enumerate(session_keys):
            sysnr, sesnr = ses.split('.')
            self.procesosors_running[id] = Event()
            # test mode - start only first worker
            if id == 0:
                self.procesosors_running[id].set()
            else:
                if self.test_mode:
                    self.procesosors_running[id].clear()
                else:
                    self.procesosors_running[id].set()
            self.processors[id] = (Processor(sysnr=sysnr, sesnr=sesnr, id=id, own_running=self.procesosors_running[id],\
                                              count_job=self.count_job, test_mode=self.get_test_mode, controller=self))
        self.view.add_worker_text_fields(len(session_keys))
        self.start_time = datetime.now()
        self.view.tick_work()
    
    def tick_work(self):
        if self.state == 'finished':
            return
        self.ticks_time_measure += 1
        time_since_last_tick = datetime.now() - self.time_in_last_tick
        self.time_in_last_tick = datetime.now()
        
        # drain report queue
        while True:
            try:
                message = self.report_q.get(block=False)
                self.view.update_worker_text_fields(message)
            except Empty:
                break
            except Exception as e:
                self.view.show_modal(type='error', message=str(e))
                break
        #check status
        live_processors = 0
        running_processors=0
        waiting_processors = 0
        stopped_processors = 0
        will_pause_processors = 0
        for id, processor in self.processors.items():
            with self.lock:
                    proc_state = processor.state
            if proc_state == 'running':
                running_processors += 1
                live_processors += 1
            elif proc_state == 'paused':
                waiting_processors += 1
                live_processors += 1
            elif proc_state == 'stopped':
                stopped_processors += 1
                live_processors += 1
            elif proc_state == 'not_started':
                live_processors += 1
            elif proc_state == 'will_pause':
                will_pause_processors += 1
                running_processors += 1
                live_processors += 1
            # dead processors not counted
            view.continue_button_update(id, proc_state)
        
        if stopped_processors == live_processors or live_processors == 0:
            self.state = 'stopped'
        elif waiting_processors+stopped_processors == live_processors and live_processors > 0:
            self.state = 'paused'
        elif running_processors > 0:
            self.state = 'running'


        if self.state == 'paused':
            self.view.work_main_button_toggle('resume')
        elif self.state == 'running':
            self.view.work_main_button_toggle('pause')
        


        if self.state == 'running' and self.ticks_time_measure > 25:
            self.ticks_time_measure = 0
            self.check_progress()
        if self.state == 'running':
            self.elapsed_time += time_since_last_tick
        q_size = self.q.qsize() if self.final_q_size is None else self.final_q_size
        self.view.update_state(running_processors=running_processors, waiting_processors=waiting_processors, stopped_processors=stopped_processors, q_size=q_size,\
                                initial_q_size=self.initial_q_size, jobs_done=self.jobs_done, remaining_time=self.remaining_time,\
                                items_per_minute=self.items_per_minute, elapsed_time=self.elapsed_time, main_label=self.state, jobs_done_error=self.jobs_done_error )

        if self.state == 'stopped':
            self.state = 'finished'
            finished_in = int((datetime.now() - self.start_time).total_seconds() // 60)
            text = (f'Done {self.jobs_done} items after {finished_in} minute{"" if finished_in==1 else "s"}.')
            self.view.work_done(text)
            self.save_result()
            self.view.work_main_button_toggle('exit')

        
    
    def save_result(self):
        #window texts    
        window_texts = self.view.get_worker_text()
        #excel (user created) output
        outxl = pd.DataFrame()
        for id, p in self.processors.items():
            if len(p.outxl_total):
                outxl = pd.concat([outxl,p.outxl_total]).copy()
        if len(outxl):
            input_file_columns = self.input_file['Input'].columns.tolist()
            outxl_columns = outxl.columns.tolist()
            reindex_columns = list(dict.fromkeys(input_file_columns+outxl_columns)) # pythonic way of duplicity removal from list
            self.input_file['Input'] = self.input_file['Input'].combine_first(outxl)
            self.input_file['Input'].update(outxl)
            self.input_file['Input'] = self.input_file['Input'].reindex(columns=(reindex_columns))
        #reports from queries
        reports = pd.DataFrame()
        for id, p in self.processors.items():
            if len(p.reportxl_total):
                reports = pd.concat([reports, p.reportxl_total]).copy()

        #save it in a file
        retry = True
        while retry:
            try:
                with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
                    self.input_file['Input'].to_excel(writer, index=False, sheet_name='Output')
                    if len(window_texts):
                        window_texts.to_excel(writer, index=False, sheet_name='Runtime_messages')
                    if len(reports):
                        reports.to_excel(writer, index=False, sheet_name='Reports')
                retry = False
            except PermissionError:
                retry = self.view.show_modal('retrycancel', 'Close the output file and click retry.')
    
    def check_progress(self):
        now = datetime.now()

        if self.jobs_done > 0:
            sec_per_item = self.elapsed_time.total_seconds() / self.jobs_done
            remaining_time_sec = int((self.q.qsize() * sec_per_item))
            self.remaining_time = timedelta(seconds=remaining_time_sec)

            if sec_per_item:
                self.items_per_minute = int(60//sec_per_item)
            else:
                self.items_per_minute = 0

    def pause(self):
        for id, p in self.processors.items():
            with self.lock:
                if self.processors[id].state == 'running':
                    self.processors[id].state = 'will_pause'
                self.procesosors_running[id].clear()

    def resume(self):
        for id, p in self.processors.items():
            with self.lock:
                if self.processors[id].state == 'will_pause':
                    self.processors[id].state = 'running'
                self.procesosors_running[id].set()

    def interrupt_work(self):
        self.final_q_size = self.q.qsize()
        while True:
            try:
                self.q.get_nowait()
                self.q.task_done()
            except Empty:
                break
    
    def get_state(self):
        return self.state
    
    def get_connection(self):
        return self.connection
    
    def processor_running_togle(self, id):
        if self.procesosors_running[id].is_set():
            self.procesosors_running[id].clear()
            with self.lock:
                if self.processors[id].state == 'running':
                    self.processors[id].state = 'will_pause'
        else:
            self.procesosors_running[id].set()
            with self.lock:
                if self.processors[id].state == 'will_pause':
                    self.processors[id].state = 'running'
            
    def ask_exit(self):
        if self.state == 'finished':
            return 'exit'
        if self.state in ('running','paused') and self.q.qsize():
            return 'cancel'
        return 'nothing'

# endregion



# region View chooser
class View:
    def __init__(self, title):
        self.title = title
        self.scheduled_tick = None
        # callback functions
        self.controller_start_work = None
        self.controller_get_state = None
        self.controller_tick_scan_sessions = None
        self.controller_get_connection = None
        self.controller_check_selected_sessions = None
        self.controller_interrupt_work = None
        self.controller_pause = None
        self.controller_resume = None
        self.controller_tick_work = None
        self.controller_processor_running_togle = None
        self.controller_ask_exit = None
        self.controller_get_test_mode = None
        self.controller_set_test_mode = None
        self.controller_test_mode_supported = None


        #first window - session chooser
        self.start_sessions_chooser()
    def start_sessions_chooser(self):
        self.session_buttons = dict()
        #make window
        self.root = tk.Tk()
        self.root.eval('tk::PlaceWindow . center')
        self.root.attributes('-topmost', True)
        self.root.title('Script...')
        #frame
        self.mainframe = tk.Frame(self.root)
        self.mainframe.pack()
        #title
        self.titleLabelText = tk.StringVar(value=self.title)
        self.titleLabel = tk.Label(self.mainframe, textvariable=self.titleLabelText, font=('',14))
        self.titleLabel.grid(row=0, column=0, columnspan=3, sticky='nwes')
        #label
        self.topLabelText = tk.StringVar()
        self.topLabel = tk.Label(self.mainframe, textvariable=self.topLabelText, font=('',12))
        self.topLabel.grid(row=1, column=0, columnspan=3, sticky='nwes')
        #additional text row
        self.topLabelText2 = tk.StringVar()
        self.topLabel2 = tk.Label(self.mainframe, textvariable=self.topLabelText2, font=('',10))
        self.topLabel2.grid(row=2, column=0, columnspan=2, sticky='we')
        #toolframe
        self.toolframe = tk.Frame(self.mainframe)#, borderwidth=6, relief=tk.RIDGE)
        self.toolframe.grid(row=3, column=0, padx=10, pady=10, sticky='w')
        #work frame
        self.frame = tk.Frame(self.mainframe, borderwidth=6, relief=tk.RIDGE)#, bg='lightgrey')
        self.frame.grid(row=4,column=0, padx=10, pady=5, columnspan=1, sticky='we')

        #toolframe:
        #button
        self.mainButton = tk.Button(self.toolframe, text='Start', command=lambda : self.start_button_action(False), padx=2, pady=2, width=10)
        self.mainButton.grid(row=1,column=0, padx=5, pady=5) #, sticky='W'
        #button start in test mode
        self.test_mode_button = tk.Button(self.toolframe, text='Start in Test Mode', command=lambda : self.start_button_action(True), padx=2, pady=2,\
                                    state=tk.DISABLED)
        self.test_mode_button.grid(row=1,column=1, padx=5, pady=5) #, sticky='W'
        #select all button
        self.selAllButton = tk.Button(self.toolframe, text='Select all', command=self.selall, padx=2, pady=2, width=10)
        self.selAllButton.grid(row=1, column=2, sticky='W', padx=5, pady=5)
        #action on close          
        self.root.protocol('WM_DELETE_WINDOW', self.window_close)
    
    def start_scan_ticks(self):
        #schedule first tick
        self.scheduled_tick = self.root.after(200, self.tick_scan_sessions_view)
    
    def start_button_action(self, test_mode=False):
        selectedses = [key for key, checkbox in self.session_buttons.items() if checkbox[1].get()]
        self.controller_set_test_mode(test_mode)
        self.controller_start_work(selectedses)

    def selall(self):
        for item in self.session_buttons.values():
            if self.selAllButton.cget('text') == 'Select all':
                item[0].select()
            else:
                item[0].deselect()
        if self.selAllButton.cget('text') == 'Select all':
            self.selAllButton.config(text='Select none')
        else:
            self.selAllButton.config(text='Select all')

    def display_job_count(self, count):
        self.topLabelText2.set(f"{count} job{'s' if count > 1 else ''}")

    def window_close(self):
        self.root.destroy()

    def update_session_buttons(self, sessions):
        for key, session in sessions.items():
            if not key in self.session_buttons.keys():
                var = tk.BooleanVar()
                # session_buttons['0.1'] = (tk.Checkbutton, variable_selected)
                self.session_buttons[key] = (tk.Checkbutton(self.frame, text=f'{session.Info.SystemName} ({session.Info.Client}) {session.Info.Transaction}', variable=var), var)

        to_remove = []
        for checkbox_key in self.session_buttons.keys():
            if not checkbox_key in sessions.keys():
                to_remove.append(checkbox_key)

        for checkbox_key in to_remove:
            self.session_buttons[checkbox_key][0].destroy()
            del(self.session_buttons[checkbox_key])
        #place the session buttons
        for r,checkbox_key in enumerate(self.session_buttons.keys()):
            self.session_buttons[checkbox_key][0].grid(row=r, column=0, sticky='w')

    def tick_scan_sessions_view(self):
        if self.controller_get_state() != 'not_started':
            return
        self.controller_tick_scan_sessions()
        if self.controller_get_connection() is not None:
            state, message = self.controller_check_selected_sessions([key for key, checkbox in self.session_buttons.items() if checkbox[1].get()])
            self.topLabelText.set(message)
            if state:
                self.mainButton.config(state=tk.NORMAL)
                self.test_mode_button.config(state=tk.NORMAL if self.controller_test_mode_supported() else tk.DISABLED)
            else:
                self.mainButton.config(state=tk.DISABLED)
                self.test_mode_button.config(state=tk.DISABLED)
        self.scheduled_tick = self.root.after(200, self.tick_scan_sessions_view)

    def show_modal(self, type='info', message='?'):
        if type=='info':
            return messagebox.showinfo('Octopus', message)
        elif type=='error':
            return messagebox.showerror('Octopus', message)
        elif type=='yesno':
            return messagebox.askyesno('Octopus', message)
        elif type=='retrycancel':
            return messagebox.askretrycancel('Octopus', message)

    # region View work window
    def start_work_window(self):
        self.root.after_cancel(self.scheduled_tick)
        for widget in self.root.winfo_children():
            widget.destroy()
        self.worker_text_fields = dict()
        self.worker_continue_buttons = dict()

        self.root.attributes('-topmost', False)
        self.root.title(self.title)
        #frame
        self.mainframe = tk.Frame(self.root)
        self.mainframe.pack(expand=True, fill='both')
        self.mainframe.grid_columnconfigure(0, weight=1)
        #label
        self.topLabelText = tk.StringVar(value='...')
        self.topLabel = tk.Label(self.mainframe, textvariable=self.topLabelText, font=('',12))
        self.topLabel.grid(row=0, column=0, columnspan=3, sticky='nwes')
        #toolframe
        self.toolframe = tk.Frame(self.mainframe)#, borderwidth=6, relief=tk.RIDGE)
        self.toolframe.grid(row=1, column=0, padx=10, pady=10, sticky='we')
        #progress bar
        proc_frame = tk.Frame(self.mainframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=2, column=0, sticky='we', padx=10, pady=5,)
        proc_frame.columnconfigure(0,weight=1)
        self.progress_bar = ttk.Progressbar(proc_frame, value=0)
        self.progress_bar.grid(row=0, sticky='we', padx=2, pady=2)

        #work frame
        self.frame = tk.Frame(self.mainframe, borderwidth=6, relief=tk.RIDGE)#, bg='lightgrey')
        self.frame.grid(row=3,column=0, padx=10, pady=5, columnspan=1, sticky='we')
        self.frame.grid_columnconfigure(0, weight=1)
        
        #toolframe:
        #button
        proc_frame = tk.Frame(self.toolframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=0, column=0, sticky='swn')
        proc_frame.rowconfigure(0,weight=1)
        self.mainButton = tk.Button(proc_frame, text='Pause', command=None, padx=2, pady=2, width=10, borderwidth=2, font=('',14))
        self.mainButton.grid(row=0,column=0, padx=5, pady=5, sticky='sn')
        #frame for slider and always on top
        proc_frame = tk.Frame(self.toolframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=0, column=1, sticky='nsw')
        #slider for transparency
        self.transparency_slider = tk.Scale(proc_frame, from_=0.3, to=1.0, resolution=0.1, orient=tk.HORIZONTAL,\
                                             label='Transparency',showvalue=False, command=lambda val: self.root.attributes('-alpha', float(val)),\
                                                length=130)
        self.transparency_slider.set(1.0)
        self.transparency_slider.grid(row=0, column=0, sticky='we', padx=2, pady=2)
        #always on top checkbox
        self.always_on_top_var = tk.BooleanVar(value=False)
        self.always_on_top_widget = tk.Checkbutton(proc_frame, text='Always on top', variable=self.always_on_top_var, command=lambda: self.root.attributes('-topmost', self.always_on_top_var.get()))
        self.always_on_top_widget.grid(row=1, column=0, sticky='w', padx=2, pady=2)
        # test mode checkbox
        self.test_mode_var = tk.BooleanVar(value=self.controller_get_test_mode())
        self.test_mode_widget = tk.Checkbutton(proc_frame, text='Test Mode', variable=self.test_mode_var, command=lambda: self.controller_set_test_mode(self.test_mode_var.get()))
        self.test_mode_widget.grid(row=2, column=0, sticky='w', padx=2, pady=2)

        #processors info
        proc_frame = tk.Frame(self.toolframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=0, column=3, sticky='nsw')
        label = tk.Label(proc_frame, text='Running', font=('',12))
        label.grid(row=0, column=0, sticky='w')
        label = tk.Label(proc_frame, text='Waiting', font=('',12))
        label.grid(row=1, column=0, sticky='w')
        label = tk.Label(proc_frame, text='Stopped', font=('',12))
        label.grid(row=2, column=0, sticky='w')
        self.running_procs = tk.StringVar(value='0')
        self.waiting_procs = tk.StringVar(value='0')
        self.stopped_procs = tk.StringVar(value='0')
        label = tk.Label(proc_frame, textvariable=self.running_procs, font=('',12))
        label.grid(row=0, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.waiting_procs, font=('',12))
        label.grid(row=1, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.stopped_procs, font=('',12))
        label.grid(row=2, column=1, sticky='e')
        #done/remaining jobs
        proc_frame = tk.Frame(self.toolframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=0, column=4, sticky='nsw')
        label = tk.Label(proc_frame, text='Jobs done', font=('',12))
        label.grid(row=0, column=0, sticky='w')
        label = tk.Label(proc_frame, text='Jobs remaining', font=('',12))
        label.grid(row=1, column=0, sticky='w')
        label = tk.Label(proc_frame, text='Jobs total', font=('',12))
        label.grid(row=2, column=0, sticky='w')

        label = tk.Label(proc_frame, text='Jobs failed', font=('',12))
        label.grid(row=3, column=0, sticky='w')

        self.jobs_done = tk.StringVar(value='0')
        self.jobs_remaining = tk.StringVar(value='0')
        self.jobs_all = tk.StringVar(value='0')
        self.jobs_done_error = tk.StringVar(value='0')

        label = tk.Label(proc_frame, textvariable=self.jobs_done, font=('',12))
        label.grid(row=0, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.jobs_remaining, font=('',12))
        label.grid(row=1, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.jobs_all, font=('',12))
        label.grid(row=2, column=1, sticky='e')
        self.label_jobs_err = tk.Label(proc_frame, textvariable=self.jobs_done_error, font=('',12))
        self.label_jobs_err.grid(row=3, column=1, sticky='e')

        #done/remaining time
        proc_frame = tk.Frame(self.toolframe, borderwidth=6, relief=tk.RIDGE)
        proc_frame.grid(row=0, column=5, sticky='nsw')
        label = tk.Label(proc_frame, text='Elapsed time', font=('',12))
        label.grid(row=0, column=0, sticky='w')
        label = tk.Label(proc_frame, text='ETA', font=('',12))
        label.grid(row=1, column=0, sticky='w')
        label = tk.Label(proc_frame, text='Items/min', font=('',12))
        label.grid(row=2, column=0, sticky='w')
        self.time_elapsed = tk.StringVar(value='0')
        self.time_remaining = tk.StringVar(value='0')
        self.items_minute = tk.StringVar(value='0')
        label = tk.Label(proc_frame, textvariable=self.time_elapsed, font=('',12))
        label.grid(row=0, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.time_remaining, font=('',12))
        label.grid(row=1, column=1, sticky='e')
        label = tk.Label(proc_frame, textvariable=self.items_minute, font=('',12))
        label.grid(row=2, column=1, sticky='e')



        #action on close          
        self.root.protocol('WM_DELETE_WINDOW', self.work_window_close)
        #schedule first tick
        self.scheduled_tick = self.root.after(200, self.tick_work)

    def add_worker_text_fields(self, amount):
        for i in range(amount):
            text_frame = tk.Frame(self.frame, borderwidth=4, relief=tk.RIDGE)
            text_frame.grid_columnconfigure(0, weight=1)
            text_output = tk.Text(text_frame, width=160, height=6, state=tk.DISABLED)
            text_output.tag_config("red", foreground="red")
            text_output.tag_config("green", foreground="green")
            text_output.tag_config("black", foreground="black")
            continue_button = tk.Button(text_frame, width=1, state=tk.NORMAL, text='>', background='green', command= lambda id=i: self.controller_processor_running_togle(id))
            self.worker_continue_buttons[i] = continue_button
            text_output.grid(row=1, column=0, pady=0, padx=0, sticky='nwes')
            continue_button.grid(row=1, column=1, sticky='ens', pady=0)
            text_frame.grid(row=i, column=0, sticky='we', pady=2, padx=2)
            self.worker_text_fields[i] = text_output
        self.root.eval('tk::PlaceWindow . center')

    def continue_button_update(self, id, state):
        color = 'grey'
        if state == 'running':
            color = 'green'
        elif state == 'paused':
            color = 'red'
        elif state == 'stopped':
            color = 'grey'
            self.worker_continue_buttons[id].configure(state=tk.DISABLED)
        elif state == 'will_pause':
            color = 'yellow'
        elif state == 'dead':
            color = 'black'
            self.worker_continue_buttons[id].configure(state=tk.DISABLED)
        self.worker_continue_buttons[id].configure(bg = color)

    def work_window_close(self):
        state = self.controller_ask_exit()
        if state == 'cancel':
            if messagebox.askyesno('Quit now?', message='Cancel all running threads?'):
                self.controller_interrupt_work()
        elif state == 'exit':
            self.root.destroy()
    
    def work_main_button_toggle(self, button_state):
        if button_state == 'pause':
            self.mainButton.config(text='Pause')
            self.mainButton.config(background='lawn green')
            self.mainButton.config(command=self.controller_pause)
        elif button_state == 'resume':
            self.mainButton.config(text='Resume')
            self.mainButton.config(background='tomato2')
            self.mainButton.config(command=self.controller_resume)
        elif button_state == 'exit':
            self.mainButton.config(text='Exit')
            self.mainButton.config(background='gold2')
            self.mainButton.config(command=self.exit)

    def tick_work(self):
        self.controller_tick_work()
        self.scheduled_tick = self.root.after(200, self.tick_work)

    def update_worker_text_fields(self, message):
        try:
            id, text, color = message
        except ValueError:
            return
        if not color in ('red', 'green'):
            color = 'black'
        self.worker_text_fields[id].config(state=tk.NORMAL)
        #self.text_output.insert("1.0", text_to_add+'\n') 
        self.worker_text_fields[id].insert(tk.END, text, color) 
        self.worker_text_fields[id].see(tk.END)
        self.worker_text_fields[id].config(state=tk.DISABLED)

    def get_worker_text(self):
        #window texts    
        window_texts = pd.DataFrame()
        for i, p in self.worker_text_fields.items():
            text = self.worker_text_fields[i].get('1.0',tk.END)
            lines = text.split('\n')
            for l in lines:
                if l in ('\n',''): continue
                idx = len(window_texts)
                window_texts.loc[idx, 'Session#'] = i+1
                window_texts.loc[idx, 'Text'] = l
        return window_texts

    def update_state(self, running_processors, waiting_processors, stopped_processors, q_size, initial_q_size, jobs_done, remaining_time, items_per_minute, elapsed_time, main_label, jobs_done_error):
        self.running_procs.set(str(running_processors)) 
        self.waiting_procs.set(str(waiting_processors))
        self.stopped_procs.set(str(stopped_processors))
        self.jobs_done.set(str(jobs_done))
        self.jobs_remaining.set(str(q_size))
        self.jobs_all.set(str(initial_q_size))
        self.time_elapsed.set(str(elapsed_time).split('.')[0])
        self.time_remaining.set(str(remaining_time).split('.')[0])
        self.items_minute.set(str(items_per_minute))
        if initial_q_size:
            self.progress_bar.config(value= jobs_done/initial_q_size * 100 -0.01 )
        self.topLabelText.set(main_label)
        self.jobs_done_error.set(str(jobs_done_error))
        self.label_jobs_err.config(foreground='red' if jobs_done <= jobs_done_error*2 and jobs_done_error > 0 else 'black')
    
    def work_done(self, status_text):
        self.topLabelText.set(status_text)
        if self.scheduled_tick is not None:
            self.root.after_cancel(self.scheduled_tick)


    def exit(self):
        if self.scheduled_tick is not None:
            self.root.after_cancel(self.scheduled_tick)
        self.root.destroy()
# endregion

# region main
if __name__ == '__main__':
    view = View(script_name)
    controller = Controller(view, script=job, grouping=grouping, test_mode_supported=test_mode_supported)
    view.root.mainloop()