# Octopus 2
Is a control engine for SAP GUI scripting. It provides a GUI for controlling the conections and manages input and output files.<br>
**The main feature is its ability to run in multiple transactions at a time.**<br>
<img width="600" height="337" alt="2026-06-08 11_14_57-" src="https://github.com/user-attachments/assets/7c6ce3f1-d2c0-4fce-9f81-6d107a850330" />
<br>
## SAP Script
To customize the script, modify the file `sap_script.py`<br>
Follow the remarks in the example file provided to learn how to put your script in it.<br> Basic python and pandas knowledge would be helpful.
## Input file
**`input.xlsx`**
### sheet Input
List of the tasks with parameters are provided in excel input file  in sheet **`Input`**.<br>
This sheet defines the queue of tasks the worker threads pick their jobs from.<br>
Input for the example script:
Replace MAT_1, MAT_2, MAT_3, and plant 0001 with your valid materials and plant so that you can see the example in action. This sheet is of course mandatory.
| Material   |   Plant | Get   | Option   |
|:-----------|--------:|:------|:---------|
| MAT_1      |    0001 | matgr |          |
| blabla     |    0001 | bla   |          |
| blabla     |    0001 | bla   |          |
| MAT_1      |    0001 | dismm |          |
| MAT_2      |    0001 | matgr |          |
| MAT_2      |    0001 | dismm |          |
| MAT_3      |    0001 | matgr |          |
| MAT_3      |    0001 | dismm |          |
| MAT_1      |    0001 | lang  | NL       |
| MAT_1      |    0001 | lang  | AY       |

### sheet Settings
Sheet `Settings` can contain some global parameters for the script, for example ECM to be used if necessary, etc...<br>
This sheets is automatically converted into a dictionary. 
| Setting            | Value      | Remark                                                                                                 |
|:-------------------|:-----------|:-------------------------------------------------------------------------------------------------------|
| Some Global Option | some value | Use this to specify some global variables for your script, for example: ECM to be used if needed, etc… |

### Other sheets in the input file
All sheets of the input file are accessible in dictionary `self.input_file`<br>
(In the example, this sheet is named `Another_sheet`)
| A                | B             |
|:-----------------|:--------------|
| Some global data | can live here |
<br>
All is ilustrated in the example script. Build the input excel file as described to see the example running.

### Grouping
In the top of the `sap_script.py`, there's variable `grouping`.<br>
If it's set to `None`, every row in the input is one job that is processed and saved.<br>
If it is a list of column names of the input, then one job is a dataframe of unique combinations of those columns. Then, the script can process all those rows before finaly saving the changes.<br>
Ilustrated in the example, try changing which value is commented and see the different behavior.

### The input file must be closed before executing the program.

## Test mode
In the end of the job, you can test `self.test_mode()`. If it returns `False`, just save the changes and the script goes on with the next job in the queue. If it returns `True`, do not save but call `self.pause('message...')`.
The script will pause here until you press the continue button in the Octopus2 window. If you do not want to treat the save or not condition, just set the variable `test_mode_supported` to False and the test mode won't be avaiable.
If allowed, the script can be either started normally or in the test mode. If it's started in test mode, only one worker thread starts working after start. The test mode can also be toggled on/off at any time, regardless how the work was started.<br>
As long as the test mode checkbox is checked, each worker should pause instead of saving the changes when it's done. You can also call the pause function anywhere in the script you want.

## Starting the script
Have one folder for each different script you make. In each folder, there need to be these files:
-`octpopus2.pyw` The executable, as it stands, not to be modified (unless you want to).
-`sap_script.py` The actual sap script you modify to meet your goal. 
-`input.xlsx` The excel input file.
<br>
Run the `octpopus2.pyw` and a session chooser appears:<br>
<img width="325" height="311" alt="image" src="https://github.com/user-attachments/assets/bf397c05-e013-4929-a4cd-cb9967f02665" />
<br>
Pick as many sessions you want but less then the number of jobs. Start either with or without test mode.<br>
For each selected session, one text output will appear in main window. Left to it is a button which can pause/continue each individual worker. When it's green, the worker is running and clicking the button changes it's color to yellow, which means, the script will pause before it graps another job from the queue. When it's red, it's waiting to be clicked and then it continues with work.<br>
There's also main big button which can pause/continue all workers. After all jobs are done and output file is saved, the main button changes function from start/pause to Exit.<br>
To interrupt the running script, click the window close button and choose yes. It will remove the remaining jobs from the queue but the currently running tasks still need to finish.

## Output file
After script finishes, it saves an output file `output.xlsx`, where you find the data you decided to appear there along with the original input. Automatic column `Script result` indicates whether the job suceeded or failed on an exception. Study the provided example to find out more about the output file.

## Multi-threading
SAP GUI can draw only one window at a time. Therefore, if a script is very quick, it may not have much significant impact when it runs in multiple transactions at a time; all the threads wait for the other SAP window to re-draw. On the other hand, if the script often waits for response from SAP server, that's where running the script in multiple threads can speed the process significantly.

## The provided example only gets some data from mm03
Do not be afraid to run it repeatedly to learn about the program.

Feedback appreciated. No responsibility taken. Use at your own risk. Hopefully enjoy :)



