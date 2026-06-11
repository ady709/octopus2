import pandas as pd

####### Control variables ########################################
### Gouping of items from input file #
grouping = None
#grouping = ['Material', 'Plant']
### Name of the script
script_name = 'Example'
### is test mode treated? (If True, make condition before saving, do not save if self.test_mode() == True
test_mode_supported = True
#########################################################################################################################
###### Optional function for task initialization. It runs in the main thread before the workers are launched.
# If you get any data you want to use later in the workers, save them to self.any_variable_name_you_want (except those already used in the controller class ;) ).
# Access the data in workers by self.controller.any_variable_name_you_want. Don't write into the variable from workers, or use with self.lock.
# Have init_task = None if you do not want any init task.
# (The purpose is for example to check some parameters in SU3 before processing the main task queue)
# Return True if all is right and the main queue can be processed, otherwise return False
###############################################################################################################################
# init_task = None
def init_task(self, session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nsu3"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").select()
    self.sap_user_date_format = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").CurListBoxEntry.Value
    return True

### Useful function you may find very helpful, see example of use
def find_in_table(session, table_id:str, column:int|str, text:str, from_beginning:bool = True) -> tuple:
    if type(column) is int:
        column_nr = column
    rows_on_page = session.findById(table_id).VisibleRowCount
    if from_beginning:
        session.findById(table_id).VerticalScrollbar.Position = 0
    max_row = session.findById(table_id).VerticalScrollbar.Maximum
    last_pos = session.findById(table_id).VerticalScrollbar.Position
    while True:
        for i in range(0, rows_on_page):
            absolute_row = i+session.findById(table_id).VerticalScrollbar.Position
            if  absolute_row > max_row:
                return (-1,-1)
            #find column number by name if needed
            if type(column) is str:
                for c in range(session.findById(table_id).Rows.ElementAt(i).Count):
                    if session.findById(table_id).Rows.ElementAt(i).ElementAt(c).Name == column:
                        column_nr = c
                        break     
            #get value
            row_val = session.findById(table_id).Rows.ElementAt(i).ElementAt(column_nr).Text
            if row_val == text:
                return (absolute_row, i)
        session.findById(table_id).VerticalScrollbar.Position += rows_on_page
        new_pos = session.findById(table_id).VerticalScrollbar.Position
        if new_pos == last_pos:
            return (-1,-1)
        last_pos = new_pos
    return (-1,-1)



### The job
def job(self, df):
    if not len(df):
        return False
    session = self.session # change for Session = self.session if you want to use Session. in your script
    #########################################################################################################################
    ###################             YOUR SCRIPT CODE GOES HERE        #######################################################
    #########################################################################################################################
    # if you detect your SAP object is currently locked by other process/user: return 'postpone'
    # if you detect an error: return False
    # if script reaches the end successfuly: return True
    #
    # one job is launched with a dataframe containing all rows corresponding with grouping
    # if grouping == None -> fataframe is only one row of the input.
    # grouping is inteded for processing multiple rows of the input in one go before saving the outcome in sap
    # --- Try changing grouping from None to ['Material', 'Plant']
    # with grouping = None -> transaction is accessed and saved for each row of the input
    # with grouping = ['Material', 'Plant'], transaction is accessed only once per Material&Plant
    #   then all rows of this job are processed and finally it is saved
    #
    # input file can also be accessed directly to reach some global information: self.input_file['sheet_name']
    # sheet Settings of the input file is already transformed to dictionary self.settings.
    # 
    # put your run-time messages to self.text_update(message, [color:black|red|green]) color is optional and black is default
    # put your output results to self.outxl[column]. first make sure the columns are there with NA value, see output_columns = ...
    output_columns = ['Value']
    for c in output_columns:
        if not c in self.outxl.columns:
            self.outxl[c] = pd.NA
    # if you get data from queries, put them in self.reportxl (just put them there, it's merged automatically)

    ### assign key variables
    mat = df.iloc[0]['Material']
    plnt = df.iloc[0]['Plant']

    #putting some info to the text output right in the beginning would be a good idea
    self.text_update(f'{mat} {plnt}: ')

    ### Start the transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
    session.findById("wnd[0]").sendVKey(0) #<<< 0 must be in brackets
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = mat
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[19]").press() #<<< press must be a function call with ()
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = True #<<< True (with capital T)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    ###################### Loop through rows of the job ###############################
    for idx, row in df.iterrows():
        ### Do the task defined on the row
        task(self, idx, row)

    ############# Delete these examples from your real script
    ### just an example of another sheets of the input file
    # Sheet Settings automatically accessible in a dictionary self.settings
    # if you know you have the sheets in your input, no need for catching the exceptions
    try:
        self.text_update (f"Settings example: {self.settings['Some Global Option']}; ")
    except Exception:
        pass
    # Other sheets accessible in self.input_file, which is a dictionary of dataframes
    try:
        self.text_update(f"Other sheet example: {self.input_file['Another_sheet'].loc[0,'A']} {self.input_file['Another_sheet'].loc[0,'B']}; ")
    except Exception:
        pass
    # access variable saved during optional init_task
    try:
        self.text_update(f"Init task data: {self.controller.sap_user_date_format}; ")
    except Exception:
        pass
    
    
    #########################################
    ### Save task, but pause if test mode!
    if not self.test_mode():
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # !!! In this example, it is not a save button but back button !!!
    else:
        self.pause('Paused and not saved, review, save manually, and continue...')
    return True

    # End main function #####################################################################################################

def task(self, idx, row):
    session = self.session
    #do the script
    get = row['Get'] 
    if get == 'matgr':
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01").select()
        matgr = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MATKL").text #members of sapgui are not case sensitive
        #put some info to text window so that you can see the progress. This will be also saved when leaving the program.
        #care about speces between the messages
        self.text_update(f'Value of {get} is {matgr}; ')
        #put the info to the output excel file
        self.outxl.loc[idx,'Value'] = matgr
    elif get == 'dismm':
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select()
        session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = row['Plant']
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        dismm = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text 
        #put some info to text window so that you can see the progress. This will be also saved when leaving the program.
        #care about speces between the messages
        self.text_update(f'Value of {get} is {dismm}; ')
        #put the info to the output excel file
        self.outxl.loc[idx,'Value'] = dismm
    elif get == 'lang':
        language = row['Option']
        session.findById("wnd[0]/tbar[1]/btn[30]").press()
        table_id = 'wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT'
        #this function will scroll dows the table until it finds text and returns absolute row and row in the current page
        abs_row, page_row = find_in_table(session=session, table_id=table_id, column=0, text=language)
        #so that now we know on which page certain value is. If the value is not found, -1, -1 is returned
        if abs_row > -1:
            lang_text = session.findById(f'{table_id}/txtSKTEXT-MAKTX[1,{page_row}]').text
        else:
            lang_text = '?'
        #put some info to text window so that you can see the progress. This will be also saved when leaving the program.
        #care about speces between the messages
        self.text_update(f'{language} description is {lang_text}; ')
        #put the info to the output excel file
        self.outxl.loc[idx,'Value'] = lang_text   
        session.findById("wnd[0]/tbar[0]/btn[3]").press() 
    else:
        self.text_update(f'Don\'t know what to get ?? ')
        #put the info to the output excel file
        self.outxl.loc[idx,'Value'] = '???'
        #example of input that was not converted to text
        self.text_update(f"Date is {row['Date']}, number is {row['Number']}; ")
    
    return


    