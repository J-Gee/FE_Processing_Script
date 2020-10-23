import datetime
import tkinter as tk
from tkinter import ttk, filedialog
from os import listdir
from os.path import isfile, join
import pandas as pd
import seaborn as sns
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
import xlwings as xw

'''
Author: Jack C. Gee
'''
#########################################################
'''
PARAMETERS
'''
template_filename = "RESULTS TEMPLATE"
filetype = ".csv"
template_dir = "C:/Users/jackh/OneDrive - The University of Liverpool/PhD/First Year/Master Data/GProcessing/Excel " \
               "Results (Processed)/Template files/"
output_dir = "C:/Users/jackh/OneDrive - The University of Liverpool\PhD\First Year/Master Data/GProcessing/Excel " \
             "Results (Processed)/"
default_batch_loc = "C:/Users/jackh/OneDrive - The University of Liverpool/PhD/First Year/Master " \
                    "Data/GProcessing/Hiden Output (Unprocessed)/"

every_nth_file = 5 # Which nth file is the one to process
# 5th is default from 2 cup 3 from sample samples

headspace_volume = 7  # (mL) Correct with more accurate value when possible
pressure = 1  # (atm)
ideal_gas_cons = 0.08205  # (L*atm / K*mol)
temperature = 293  # (K)

molar_vol_gas = (ideal_gas_cons * temperature) / pressure


#########################################################

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        master.title("FE Processing Script")
        self.pack()
        container = tk.Frame(master)
        '''
        Builds a TK frame from master, container is the contents of the window (frame), master is the whole application.
        '''
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Make frames to display pages
        self.frames = {}
        for F in (StartMenu, ExcelView, Output):
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartMenu)

    # Raises desired frame to top
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def update_frame(self, cont, data_type, data=None):
        frame = self.frames[cont]
        if data_type == "listbox_option_update":
            frame.excelview_listbox_options_update(data)
        if data_type == "listbox_selected_update":
            frame.excelview_listbox_selected_update(data)
        if data_type == "return_parameters":
            r_list = frame.excelview_return_parameters()
            return r_list
        if data_type == "data_processing":
            frame.output_update(data)
        if data_type == "update_graph":
            frame.output_update_graphs()
        if data_type == "update_dropboxes":
            frame.output_update_dropboxes(data)


class StartMenu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.startmenu_create_widgets()
        self.pack()

    def startmenu_create_widgets(self):
        self.start_menu_frame = tk.LabelFrame(self)
        self.start_menu_frame.grid(ipadx=100, ipady=100)
        # labels

        view_all_label = tk.Label(self.start_menu_frame, text="View all by:")

        view_all_label.grid(row=3, column=1, rowspan=1, padx=0, sticky="ew")

        # buttons
        process_new_batch_button = ttk.Button(self.start_menu_frame, text="Process new batch", command=data_processing)
        view_by_access_button = ttk.Button(self.start_menu_frame, text="Access", command=view_by_access)
        view_by_excel_button = ttk.Button(self.start_menu_frame, text="Excel", command=view_by_excel)

        process_new_batch_button.grid(row=2, column=3, rowspan=1, padx=0, sticky="ew")
        view_by_access_button.grid(row=3, column=2, rowspan=1, padx=0, sticky="ew")
        view_by_excel_button.grid(row=3, column=3, rowspan=1, padx=0, sticky="ew")


class ExcelView(tk.Frame):
    # self is own self, parent is container and controller is master
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.excelview_create_widgets()
        # self.update_label()

    def excelview_create_widgets(self):

        # Submit needs to start processing

        self.LHframe = tk.LabelFrame(self)
        self.RHframe = tk.LabelFrame(self)
        self.LHframe.grid(row=0, column=0, sticky="w", padx=10, rowspan=2)
        self.RHframe.grid(row=1, column=1, sticky="wn", padx=50)

        # Listbox - Options
        self.listbox_options = tk.Listbox(self.LHframe, width=60, height=40, borderwidth=2)
        self.listbox_options.grid(row=2, rowspan=2, column=0, padx=10)
        scrollbar_options = tk.Scrollbar(self.LHframe, orient="vertical")
        scrollbar_options.config(command=self.listbox_options.yview)
        self.listbox_options.config(yscrollcommand=scrollbar_options.set)
        scrollbar_options.grid(row=2, rowspan=2, column=0, sticky="ens")

        # Listbox - Selected
        self.listbox_selected = tk.Listbox(self.LHframe, width=60, height=40, borderwidth=2)
        self.listbox_selected.grid(row=2, rowspan=2, column=4, padx=10)
        scrollbar_selected = tk.Scrollbar(self.LHframe, orient="vertical")
        scrollbar_selected.config(command=self.listbox_selected.yview)
        self.listbox_selected.config(yscrollcommand=scrollbar_selected.set)
        scrollbar_selected.grid(row=2, rowspan=2, column=4, sticky="ens")

        # Labels
        select_directory_label = tk.Label(self.LHframe, text="Processed Batches")
        select_directory_label.grid(row=0, column=0, rowspan=2, padx=30, sticky="w")
        selected_label = tk.Label(self.LHframe, text="Selected Batches:")
        selected_label.grid(row=0, column=4, sticky="w", ipadx=10)

        # Buttons
        remove_choice_button = tk.Button(self.LHframe, text="<", command=listbox_remove_choice)
        remove_choice_button.grid(row=2, column=2, sticky="wes", ipadx=10)
        select_choice_button = tk.Button(self.LHframe, text=">", command=listbox_select_choice)
        select_choice_button.grid(row=2, column=3, sticky="wes", ipadx=10)
        remove_all_button = tk.Button(self.LHframe, text="<<", command=listbox_remove_all)
        remove_all_button.grid(row=3, column=2, sticky="wen", ipadx=10, pady=5)
        select_all_button = tk.Button(self.LHframe, text=">>", command=listbox_select_all)
        select_all_button.grid(row=3, column=3, sticky="wen", ipadx=10, pady=5)

        submit_nav_output = tk.Button(self.RHframe, text="View Selected",
                                      command=view_selected_batches)
        submit_nav_output.grid(sticky="es", row=6,
                               column=1,
                               ipadx=10)  # get(0, END) to get whole list, get(ACTIVE) for highlighted, delete(Active)

    def excelview_listbox_options_update(self, data):
        self.listbox_options.delete(0, "end")
        for i in data:
            self.listbox_options.insert("end", i)

    def excelview_listbox_selected_update(self, data):
        if data == 0:  # remove
            self.listbox_options.insert("end", self.listbox_selected.get("active"))
            self.listbox_selected.delete("active")
        if data == 1:  # add
            self.listbox_selected.insert("end", self.listbox_options.get(
                "active"))  # adds active to selected and removes from options
            self.listbox_options.delete("active")
        if data == 2:  # remove all
            list = self.listbox_selected.get(0, "end")
            self.listbox_selected.delete(0, "end")
            for i in list:
                self.listbox_options.insert("end", i)
            self.listbox_selected.delete(0, "end")
        if data == 3:  # add all
            list = self.listbox_options.get(0, "end")
            self.listbox_options.delete(0, "end")
            for i in list:
                self.listbox_selected.insert("end", i)

    def excelview_return_parameters(self):
        if self.listbox_selected.get(0) == "":
            return "Error - Please select files to process"
        '''
        .get() returns int value for state, 0 = unselected, if total = 0 then nothing selected, throw error
        '''
        s_list = self.listbox_selected.get(0, "end")  # selected list
        return s_list


class Output(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

    def output_create_widgets(self):
        self.output_LHframe = tk.LabelFrame(self, text="Parameters")
        self.output_RHframe = tk.LabelFrame(self, text="Seaborn Graphics")
        self.output_RHframe_toolbarframe = tk.Frame(self)

        self.output_LHframe.grid(row=0, column=0, sticky="w", padx=10)
        self.output_RHframe.grid(row=0, column=1, sticky="we")
        self.output_RHframe_toolbarframe.grid(row=1, column=1, sticky="e")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.output_LHframe.columnconfigure(0, weight=2)
        self.output_LHframe.rowconfigure(0, weight=2)
        self.output_RHframe.columnconfigure(0, weight=1)
        self.output_RHframe.rowconfigure(0, weight=1)

        # Label config
        preset_label = tk.Label(self.output_LHframe, text="Presets")
        graph_parameter_label = tk.Label(self.output_LHframe, text="Graph Parameters")
        plot_type_label = tk.Label(self.output_LHframe, text="Plot Type:")
        x_axis_label = tk.Label(self.output_LHframe, text="X Axis:")
        y_axis_label = tk.Label(self.output_LHframe, text="Y Axis:")
        hue_label = tk.Label(self.output_LHframe, text="Hue:")

        # Button config
        nav_file_browser = ttk.Button(self.output_LHframe, text="Return",
                                      command=lambda: self.controller.show_frame(StartMenu))
        process_graph_button = ttk.Button(self.output_LHframe, text="Process", command=update_graphs)

        stripplot_button = ttk.Button(self.output_LHframe, text="Stripplot", command=lambda: update_graphs("stripplot"))
        lineplot_button = ttk.Button(self.output_LHframe, text="Lineplot", command=lambda: update_graphs("lineplot"))

        # Dropbox config
        headers = ["form_id"]
        headers += list(self.processed_output_df)
        plot_type_options = [
            "stripplot", "lineplot"
        ]

        x_axis_options = headers

        y_axis_options = headers

        hue_options = headers

        self.plot_type_var = tk.StringVar(self.output_LHframe)
        self.plot_type_var.set("None")  # default value
        plot_type_drop = tk.OptionMenu(self.output_LHframe, self.plot_type_var, *plot_type_options)

        self.x_axis_var = tk.StringVar(self.output_LHframe)
        self.x_axis_var.set("None")  # default value
        x_axis_drop = tk.OptionMenu(self.output_LHframe, self.x_axis_var, *x_axis_options)

        self.y_axis_var = tk.StringVar(self.output_LHframe)
        self.y_axis_var.set("None")  # default value
        y_axis_drop = tk.OptionMenu(self.output_LHframe, self.y_axis_var, *y_axis_options)

        self.hue_var = tk.StringVar(self.output_LHframe)
        self.hue_var.set("None")  # default value
        hue_drop = tk.OptionMenu(self.output_LHframe, self.hue_var, *hue_options)

        # Graph config
        self.fig = Figure(figsize=(13, 7), dpi=100)
        self.graph_canvas = FigureCanvasTkAgg(self.fig, master=self.output_RHframe)

        self.graph_canvas.get_tk_widget().grid(column=0, row=0, rowspan=5, columnspan=5)
        self.canvas_toolbar = NavigationToolbar2Tk(self.graph_canvas, self.output_RHframe_toolbarframe)

        # LH packing
        preset_label.grid(row=0, column=0, pady=10)
        graph_parameter_label.grid(row=0, column=1, pady=10)
        plot_type_label.grid(row=1, column=1, pady=10)
        x_axis_label.grid(row=1, column=2, pady=10)
        y_axis_label.grid(row=1, column=3, pady=10)
        hue_label.grid(row=1, column=4, pady=10)
        plot_type_drop.grid(row=2, column=1, pady=10)
        x_axis_drop.grid(row=2, column=2, pady=10)
        y_axis_drop.grid(row=2, column=3, pady=10)
        hue_drop.grid(row=2, column=4, pady=10)
        nav_file_browser.grid(row=5, column=0, pady=10)
        process_graph_button.grid(row=5, column=5, pady=10)
        stripplot_button.grid(row=1, column=0, pady=10)
        lineplot_button.grid(row=2, column=0, pady=10)

    def output_update_dropboxes(self, parameters=None):
        # updates dropboxes for presets and collects selection for use in graph update.
        self.plot_type_var.set(parameters[0])
        self.x_axis_var.set(parameters[1])
        self.y_axis_var.set(parameters[2])
        self.hue_var.set(parameters[3])

    def output_update_graphs(self):
        self.fig.clear()
        chart = self.fig.subplots()

        # Get parameters from dropboxes
        plot_type = self.plot_type_var.get()
        x_var = self.x_axis_var.get()
        y_var = self.y_axis_var.get()
        hue_var = self.hue_var.get()
        # quick check to see if using index string as axis, pulls correct column from dataframe and rotates text (45d)
        if x_var == "form_id":
            x_var = self.processed_output_df.index
            for item in chart.get_xticklabels():
                item.set_rotation(45)

        if y_var == "form_id":
            y_var = self.processed_output_df.index

        if plot_type == "stripplot":
            chart = sns.stripplot(x=x_var,
                                  y=y_var,
                                  hue=hue_var,
                                  data=self.processed_output_df,
                                  ax=chart)

        if plot_type == "lineplot":
            chart = sns.lineplot(x=x_var,
                                 y=y_var,
                                 hue=hue_var,
                                 data=self.processed_output_df,
                                 ax=chart)

        self.graph_canvas.draw()

    def output_update(self, data):
        '''
        Checks saving directory,
        Generates output filename,
        reads in CSV,
        process data to get avg, 2sd and gas mol,
        adds this data to dataframe,
        moves to next csv; iterates through all in batch,
        output to single CSV
        '''

        # Generates output file using template
        global template_filename, template_dir, output_dir, filetype
        print(data[2])
        self.output_filename = (data[2]) + "_" + str(datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
        # batch_id + yyyy-mm-dd_HH-MM-SS

        global headspace_volume, molar_vol_gas, every_nth_file
        nth_counter = 1
        processed_output_dict = {}
        for i in data[0]:
            i2 = (i.split("_"))[2]  # takes file name, splits for FormulationIdxxxxxx
            if nth_counter != every_nth_file: # skips through scans until desired is reached
                nth_counter += 1
                continue
            nth_counter = 1 # resets counter as desired has been reached
            form_datetime = i.split("_")[3]  # takes the datetime from formulation filename
            form_datetime = form_datetime.split(".")[0:2]  # cuts off .csv extension
            form_datetime = "20" + form_datetime[0] + form_datetime[1]
            # adds 20 to start of date to give date format e.g 20200314
            if processed_output_dict == {}:  # If nothing in dict, first should be an n2_sample
                pass
                #processed_output_dict.setdefault("N2_samp", []).append((True))

            elif i2.split("Id")[1] not in processed_output_dict["form_id"]:
                #processed_output_dict.setdefault("N2_samp", []).append((True))
                '''If nothing in dict, add this first formID as add. sampling
                If this formID not in dict, mark this as the add. sampling (As this occurs first, 
                if the list is ordered then this should be the first entering the dict)'''

            else:
                pass
                #processed_output_dict.setdefault("N2_samp", []).append((False))

            processed_output_dict.setdefault("form_id", []).append((i2.split("Id"))[1])  # takes the xxxxxxx from formID
            processed_output_dict.setdefault("form_datetime", []).append(form_datetime)
            to_skip = list(range(0, 32)) + [33, 34] # reads to line 33 in csv (headers), then skips first 2 scans
            current_file_df = pd.read_csv((data[1] + "/" + i), skiprows=to_skip)
            # Reads the formatted output sheet into a dataframe
            current_file_df = current_file_df.dropna(1)
            # removes empty space / NaNs to the right
            current_file_df.rename_axis()

            for col in current_file_df.columns:
                if "%" in col or "Baratron" in col:
                    processed_output_dict.setdefault(("{} Avg").format(col), []).append(
                        current_file_df[("{}").format(col)].mean())

                    processed_output_dict.setdefault(("{} 2STD").format(col), []).append(
                        current_file_df[("{}").format(col)].std() * 2)
                if col == "% H2" or col == "% O2":
                #if "H2" in col or "O2" in col:
                    avg_gas_per = current_file_df[("{}").format(col)].mean()  # per = %
                    avg_gas_vol_mL = avg_gas_per * headspace_volume
                    if avg_gas_vol_mL == 0: # incase no desired gas in vial
                        avg_gas_umol = 0
                    else:
                        avg_gas_umol = ((avg_gas_vol_mL/1000) / molar_vol_gas) * 10 ** 6
                    # Finds gas mol, flips and divides for umol

                    processed_output_dict.setdefault("{} umol".format(col), []).append(avg_gas_umol)

        processed_output_df = pd.DataFrame.from_dict(data=processed_output_dict)
        # Converts the dictionary to a Pandas Dataframe

        processed_output_df["form_id"] = pd.to_numeric(processed_output_df["form_id"])
        processed_output_df["form_datetime"] = pd.to_datetime(processed_output_df["form_datetime"],
                                                              format="%Y%m%d%H%M%S")

        processed_output_df.set_index("form_id", inplace=True)
        self.processed_output_df = processed_output_df

        Output.output_csv_processing(self)

        self.output_create_widgets()
        app.show_frame(Output)

    def output_csv_processing(self):
        global template_filename, template_dir, output_dir, filetype

        self.processed_output_df.to_csv(output_dir + self.output_filename + ".csv")


def data_processing():
    global default_batch_loc

    dirname = filedialog.askdirectory(initialdir=default_batch_loc, title="Select batch to process")
    app.update_frame(cont=StartMenu, data_type="label_update", data=dirname)
    ''' Cont = desired page, inputType= will be label, listbox, etc; input= filename/dir'''
    dir_contents = [f for f in listdir(dirname) if isfile(join(dirname, f))]

    batch_id = (dirname.split("/"))[-1]

    r_list = dir_contents, dirname, batch_id
    # sends dir and files names to be processed - should go through all csvs in file.
    app.update_frame(cont=Output, data_type="data_processing", data=r_list)


def update_graphs(preset=None):
    '''
    Sends graph request to graph builder
    if request_type == None, no preset being used - otherwise update dropboxes to match preset then process graph.
    parameters [plotType, xVar, yVar, hueVar]
    '''
    parameters = []

    if preset == "stripplot":
        parameters = ["stripplot", "form_id", "% N2 Avg", "N2_samp"]

    if preset == "lineplot":
        parameters = ["lineplot", "% O2 Avg", "% N2 Avg", "N2_samp"]

    # If parameters is not blank - continue
    if parameters:
        app.update_frame(cont=Output, data_type="update_dropboxes", data=parameters)

    app.update_frame(cont=Output, data_type="update_graph")


def listbox_remove_choice():
    app.update_frame(cont=ExcelView, data_type="listbox_selected_update", data=0)


def listbox_select_choice():
    app.update_frame(cont=ExcelView, data_type="listbox_selected_update", data=1)


def listbox_remove_all():
    app.update_frame(cont=ExcelView, data_type="listbox_selected_update", data=2)


def listbox_select_all():
    app.update_frame(cont=ExcelView, data_type="listbox_selected_update", data=3)


def view_by_access():
    pass
    '''
    Not implemented - likely to use LABMAN's HOLLY database - need more info from MIF Team / LABMAN
    '''


def view_by_excel():
    global output_dir
    dir_contents = [f for f in listdir(output_dir) if isfile(join(output_dir, f)) and f.split(".")[1] == "csv"]
    # checks if isfile and isCSV == true
    app.update_frame(cont=ExcelView, data_type="listbox_option_update", data=dir_contents)
    app.show_frame(ExcelView)


def view_selected_batches():
    global output_dir
    return_batches = app.update_frame(cont=ExcelView, data_type="return_parameters")
    li = []
    for filename in return_batches:
        df = pd.read_csv(output_dir + filename, index_col=None, header=0)
        li.append(df)
    all_batches_df = pd.concat(li, axis=0, ignore_index=True)

    wb = xw.Book()
    sheet = wb.sheets["Sheet1"]
    sheet.range("A1").value = all_batches_df


# Maintains tkinter interface
app = Application(master=tk.Tk())
app.mainloop()
