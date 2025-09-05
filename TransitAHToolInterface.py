# -*- coding: utf-8 -*-
"""
Title: TransiT Abstraction Heirarchy Tool
Author: Joshua Duvnjak
Description: A program to load excel sheets and output an Abstraction Heirarchy
(in network format).

"""

## Load packages
import pandas as pd
import igraph as ig
import numpy as np
import matplotlib.pyplot as plt
import openpyxl

#https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
import os
import sys

#Load matplotpib packges for Tkinter https://matplotlib.org/stable/gallery/user_interfaces/embedding_in_tk_sgskip.html
from matplotlib.backend_bases import key_press_handler
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,NavigationToolbar2Tk)


##Tk Interface
#Made with code from https://tkdocs.com/tutorial/
from tkinter import *
from tkinter import ttk

#https://stackoverflow.com/questions/9239514/filedialog-tkinter-and-opening-files
from tkinter import filedialog


class MyApp(Tk):  
    def __init__(self):
        super().__init__()
        
        #Add application title
        self.title('TransiT Abstraction Heirarchy Tool')
        
        #Add style
        #https://tkdocs.com/tutorial/styles.html
        theme = ttk.Style()
        theme.theme_use('clam')
        
        #Define minimum size of the window
        self.minsize(400,400)
        
        #Set dataframes
        self.graphData = pd.DataFrame()
        self.loadedData = pd.DataFrame(columns=['phase','id','parent'])
        
        #set path variables
        self.openPath = ''
        self.savePath = ''
        
        #Set default values for the graph customisation 
        self.graphFigSize = StringVar()
        self.graphFigSize.set(8)
        self.graphTextSize = StringVar()
        self.graphTextSize.set(9)
        self.graphSquareSize = StringVar()
        self.graphSquareSize.set(15)
        self.graphLabelDist = StringVar()
        self.graphLabelDist.set(2)
        self.graphMarginSize= StringVar()
        self.graphMarginSize.set(0.15)
        self.graphLayout = StringVar()
        self.graphLayout.set('Standard')
        self.graphStatVis = StringVar()
        self.graphStatVis.set('None')
        self.graphLabelAngle = StringVar()
        self.graphLabelAngle.set(4.6)
        
        #Set defafult values for statistics checkboxes
        self.graphCloseness = StringVar()
        self.graphBetweeness = StringVar()
        self.graphEigen = StringVar()
        self.graphPagerank = StringVar()
        self.graphStrength = StringVar()
        self.addStatsToExcel = StringVar()
        
        
        #Add Menus https://tkdocs.com/tutorial/menus.html
        self.option_add('*tearOff', FALSE)
        self.appTopBar = Menu(self)

        #Add top level options to menu bar
        self.appTopFile = Menu(self.appTopBar)
        self.appTopExport = Menu(self.appTopBar)
        self.appTopStatistics = Menu(self.appTopBar)
        
        self['menu'] = self.appTopBar
        
        self.appTopBar.add_cascade(menu=self.appTopFile,label='File')
        self.appTopBar.add_cascade(menu=self.appTopExport,label='Export')
        self.appTopBar.add_cascade(menu=self.appTopStatistics,label='Add Statistics to Graph')
        self.appTopBar.add_command(label='Graph Formatting',command=self.settingsWindow)
        self.appTopBar.add_command(label='Help',command=self.helpWindow)
        self.appTopBar.add_command(label='About',command=self.aboutWindow)
        
        #set up image
        self.transitImage = PhotoImage(file=self.imagePath('transitLogo.png'))
        
    #code from https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
    #and https://cx-freeze.readthedocs.io/en/latest/faq.html
    def imagePath(self,path2):
        try:
            if getattr(sys, "frozen", False):
                formPath = sys.prefix
            else:
                formPath = os.path.dirname(__file__)
        except:
            pass
        return os.path.join(formPath,path2)

    
    #Function to set up the graph formatting window and interface
    def settingsWindow(self):
        #Create Window
        settingsGraphWindow = Toplevel(self)
        settingsGraphWindow.title('Graph Formatting')
        
        self.backgroundGraphSettings = ttk.Frame(settingsGraphWindow)
        self.backgroundGraphSettings.grid(column=0,row=0)
        
        self.graphSettingsFrame = ttk.LabelFrame(self.backgroundGraphSettings,text='Graph Formatting')
        self.graphSettingsFrame.grid(column=0,row=0,pady=(10,10),padx=(10,10))
        
        ttk.Label(self.graphSettingsFrame,text='Figure Scaling').grid(column=0,row=1)
    
        self.graphFigSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=30.0,textvariable=self.graphFigSize)
        self.graphFigSpinbox.grid(column=1,row=1)
        ttk.Label(self.graphSettingsFrame,text='Figure Scaling affects how the graph is created. The graphs width is 20% larger than the height').grid(column=0,row=2,columnspan=3,pady=(0,10))
        
        ttk.Label(self.graphSettingsFrame,text='Text Size').grid(column=0,row=3)
        self.graphTextSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=30.0,textvariable=self.graphTextSize)
        self.graphTextSpinbox.grid(column=1,row=3)
        ttk.Label(self.graphSettingsFrame,text='Text Size affects the size of the text').grid(column=0,row=4,columnspan=3,pady=(0,10))
        
        ttk.Label(self.graphSettingsFrame,text='Text distance from node').grid(column=0,row=5)
        self.graphTextSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=10.0,textvariable=self.graphLabelDist)
        self.graphTextSpinbox.grid(column=1,row=5)
        ttk.Label(self.graphSettingsFrame,text='This value affects how far the text labels are away from the node square').grid(column=0,row=6,columnspan=3,pady=(0,10))
        
        ttk.Label(self.graphSettingsFrame,text='Text Angle from Node').grid(column=0,row=7)
        self.graphTextSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=10.0,textvariable=self.graphLabelAngle)
        self.graphTextSpinbox.grid(column=1,row=7)
        ttk.Label(self.graphSettingsFrame,text='This value affects how far the text labels are away from the node square').grid(column=0,row=8,columnspan=3,pady=(0,10))
        
        ttk.Label(self.graphSettingsFrame,text='Square Size').grid(column=0,row=10)
        self.graphSquareSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=30.0,textvariable=self.graphSquareSize)
        self.graphSquareSpinbox .grid(column=1,row=10)
        ttk.Label(self.graphSettingsFrame,text='Square Size affects the size of the node squares').grid(column=0,row=11,columnspan=3,pady=(0,10))
        
        ttk.Label(self.graphSettingsFrame,text='Figure Margin Size').grid(column=0,row=13)
        #https://stackoverflow.com/questions/62884385/how-to-set-a-step-for-increment-and-decrement-of-spinbox-value-in-python
        self.graphFigSpinbox = ttk.Spinbox(self.graphSettingsFrame,from_=0.0,to=10.0,increment=0.05,textvariable=self.graphMarginSize)
        self.graphFigSpinbox.grid(column=1,row=13)
        ttk.Label(self.graphSettingsFrame,text='Figure Margin Size affects the whitespace around the edges of the graph').grid(column=0,row=14,columnspan=3,pady=(0,10))
        
        self.graphRadioFrame = ttk.Frame(self.graphSettingsFrame)
        self.graphRadioFrame.grid(column=0,row=16,pady=(10,10),padx=(10,10),columnspan=2)
        
        #Radiobutton for changing graph layout
        #https://tkdocs.com/tutorial/widgets.html
        ttk.Label(self.graphRadioFrame,text='Standard or Circular Layout').grid(column=0,row=11)
        standardLayoutButton = ttk.Radiobutton(self.graphRadioFrame,text='Standard',variable=self.graphLayout,value='Standard')
        standardLayoutButton.grid(column=0,row=12)
        circLayoutButton = ttk.Radiobutton(self.graphRadioFrame,text='Circular',variable=self.graphLayout,value='Circular')
        circLayoutButton.grid(column=1,row=12)
        
        #Radiobutton of statistics visualisations for graph
        ttk.Label(self.graphRadioFrame,text='Change Square Size based on').grid(column=0,row=13)
        noneButton = ttk.Radiobutton(self.graphRadioFrame,text='None',variable=self.graphStatVis,value='None')
        noneButton.grid(column=0,row=14)
        eigenButton = ttk.Radiobutton(self.graphRadioFrame,text='Eigenvector',variable=self.graphStatVis,value='Eigen')
        eigenButton.grid(column=1,row=14)
        closenessButton = ttk.Radiobutton(self.graphRadioFrame,text='Closeness',variable=self.graphStatVis,value='Closeness')
        closenessButton.grid(column=0,row=16)
        betweenessButton = ttk.Radiobutton(self.graphRadioFrame,text='Betweeness',variable=self.graphStatVis,value='Betweeness')
        betweenessButton.grid(column=1,row=16)
        pagerankButton = ttk.Radiobutton(self.graphRadioFrame,text='Google Pagerank',variable=self.graphStatVis,value='Pagerank')
        pagerankButton.grid(column=0,row=15)
        strengthButton = ttk.Radiobutton(self.graphRadioFrame,text='Strength',variable=self.graphStatVis,value='Strength')
        strengthButton.grid(column=1,row=15)

        ttk.Label(self.graphSettingsFrame, text='Note: exported graphs use a different library and may not appear exactly as displayed on screen.').grid(column=0, row=17)
    
    #Function to set up the help window and interface
    def helpWindow(self):
        #Create Window
        helpWindow = Toplevel(self)
        helpWindow.title('Help')
        
        self.backgroundHelp = ttk.Frame(helpWindow)
        self.backgroundHelp.grid(column=0,row=0)
        
        self.helpFrame = ttk.LabelFrame(self.backgroundHelp,text='Help')
        self.helpFrame.grid(column=0,row=0,pady=(10,10),padx=(10,10))
        ttk.Label(self.helpFrame,text='This tool is intended to be used with the Abstraction Heirarchy method').grid(column=0,row=0)
        ttk.Label(self.helpFrame,text='Please use a source like Naikar and Sanderson (2001) for help with the method').grid(column=0,row=1,pady=(0,10))
        ttk.Label(self.helpFrame,text='Terminology').grid(column=0,row=2)
        ttk.Label(self.helpFrame,text='Node: a single areftact or entity on an absraction heirarchy').grid(column=0,row=3)
        ttk.Label(self.helpFrame,text='Edge: a single connection between nodes on an absraction heirarchy').grid(column=0,row=4,pady=(0,10))
        ttk.Label(self.helpFrame,text='Technical Help').grid(column=0,row=5)
        ttk.Label(self.helpFrame,text='Please use the "Refresh Graph" button after changing graph formatting').grid(column=0,row=6,pady=(0,10))

        
    #Function to set up the help window and interface
    def aboutWindow(self):
        #Create Window
        aboutWindow = Toplevel(self)
        aboutWindow.title('About')
        
        self.backgroundAbout = ttk.Frame(aboutWindow)
        self.backgroundAbout.grid(column=0,row=0)
        self.aboutFrame = ttk.LabelFrame(self.backgroundAbout,text='About This Application')
        self.aboutFrame.grid(column=0,row=0,pady=(10,10),padx=(10,10))
        ttk.Label(self.aboutFrame,text='Lead Programmer: Joshua Duvnjak').grid(column=0,row=0)
        ttk.Label(self.aboutFrame,text='Contributor: Guy Walker').grid(column=0,row=1)
        ttk.Label(self.aboutFrame,text='This work for the TransiT Research Hub https://www.transit.ac.uk/ \n was supported by the Engineering and Physical Sciences Research Council (EPSRC) Grant Number: EP/Z533221/1').grid(column=0,row=2)
        ttk.Label(self.aboutFrame,text='Please see the code for references to other documentation that aided in the development of this work').grid(column=0,row=3)
        ttk.Label(self.aboutFrame,image=self.master.transitImage,text='Funded By the TransiT Project',compound='left').grid(column=0,row=10)

    #Function to load the data from the selected excel file to a dataframe
    def importData(self):
        #Input excel file, Load the data files into pandas
        #try:
            self.loadedData = pd.read_excel(self.openPath, header=0)
        #except:
            #print('error')
            #return
    #Function to Generate the edges between nodes   
    def loadData(self):
        #Try to Extract the required values to a numpy array      
        try:
            #Extract the required values to a numpy array
            self.graphData = self.loadedData[['phase','id','parent']]
            
            #replace empty spaces in table and conver to numpy for analysis
            self.graphData = (self.graphData.fillna('')).to_numpy()
        except:
            #Raise error if the format is incorrect
            raise ValueError('Please use the correct excel format')

        #Empty lists to store data file values
        self.lines = []
        self.names = []
        self.phases = []

        #Go through each row
        for x in range(0,self.loadedData['id'].size):
            
            #Add a column to list which can be added to nodes
            self.names+= [self.graphData[x][1]]
            self.phases+= [self.graphData[x][0]]
            
            #Find the parents of a node and then calculate which node it links too
            for parent in self.graphData[x][2].rsplit(','):
                for y in range(0,self.loadedData['id'].size):
                    if parent.strip() == self.graphData[y][1].strip():
                        self.lines += [(x,y)]
    
    #Function to Create the AH Graphs
    def createAHGraph(self,display=True):
        
        #Set figure size, an empty figure needs to be set if there is no data to plot
        if self.loadedData.empty == True:
            self.figure, self.axis = plt.subplots()
            return(self.figure)
        else:
            self.figure, self.axis = plt.subplots(layout='tight',figsize=((float(self.graphFigSize.get())*1.2,int(self.graphFigSize.get()))))

        ##Create the social network output
        #create the graph (number of nodes loaded from data, lines loaded from data)
        self.AH = ig.Graph(self.loadedData['id'].size,self.lines)
        self.AH.vs["id"]=self.names
        self.AH.vs["phase"]=self.phases

        self.AH.vs["eigenvector"]=self.AH.eigenvector_centrality()
        self.AH.vs["betweenness"]=self.AH.betweenness()
        self.AH.vs["closeness"]=self.AH.closeness()
        self.AH.vs["pagerank"]=self.AH.pagerank()
        self.AH.vs["strength"]=self.AH.strength()

        #Add Graph Metrics if checkbox selected
        for x in range(0,len(self.AH.vs["id"])):
            if self.graphEigen.get() == '1':
                self.AH.vs[x]["id"] =self.AH.vs[x]["id"]+"\n Eigenvector = "+str(self.AH.vs[x]["eigenvector"])[0:5]
            if self.graphBetweeness.get() == '1':
                self.AH.vs[x]["id"] =self.AH.vs[x]["id"]+"\n Betweenness = "+str(self.AH.vs[x]["betweenness"])[0:5]
            if self.graphCloseness.get() == '1':
                self.AH.vs[x]["id"] =self.AH.vs[x]["id"]+"\n Closeness = "+str(self.AH.vs[x]["closeness"])[0:5]
            if self.graphPagerank.get() == '1':
                self.AH.vs[x]["id"] =self.AH.vs[x]["id"]+"\n Pagerank = "+str(self.AH.vs[x]["pagerank"])[0:5]
            if self.graphStrength.get() == '1':
                self.AH.vs[x]["id"] =self.AH.vs[x]["id"]+"\n Strength = "+str(self.AH.vs[x]["strength"])[0:5]

        #Create a size based visualisation for the graph based on statitics
        #https://python.igraph.org/en/stable/tutorials/betweenness.html#sphx-glr-tutorials-betweenness-py
        if self.graphStatVis.get() == 'Eigen':
            self.AH.vs["size"] = ig.rescale(self.AH.vs["eigenvector"],(5,30))
        elif self.graphStatVis.get() == 'Closeness':
            self.AH.vs["size"] = ig.rescale(self.AH.vs["closeness"])
        elif self.graphStatVis.get() == 'Betweeness':
            self.AH.vs["size"] = ig.rescale(self.AH.vs["betweeness"])
        elif self.graphStatVis.get() == 'Pagerank':
            self.AH.vs["size"] = ig.rescale(self.AH.vs["pagerank"])
        elif self.graphStatVis.get() == 'Strength':
            self.AH.vs["size"] = ig.rescale(self.AH.vs["strength"])
        

        #Set colours for individual phases
        self.phase_colours={'Purpose':'LightGreen','Values':'LightSeaGreen','Functions':'LightSteelBlue','Processes':'LightSkyBlue','Physical':'MediumSlateBlue'}

        #Create a layout (graph node formatting)
        if self.graphLayout.get() == 'Standard':
            self.layout = self.AH.layout('tree')
        else:
            self.layout = self.AH.layout('rt_circular')
        
        #Plot the social network graph
        self.axis.margins(float(self.graphMarginSize.get()),0)
        
        #A range of plot types dependent on how the visualisation types
        if self.graphStatVis.get() == 'None':
            if display:
                self.layout.rotate(180)
                ig.plot(
                    self.AH,
                    target=self.axis,
                    layout=self.layout,
                    vertex_label=self.AH.vs["id"],
                    vertex_label_dist=self.graphLabelDist.get(),
                    vertex_label_angle=self.graphLabelAngle.get(),
                    vertex_shape='rectangle',
                    vertex_label_size=self.graphTextSize.get(),
                    vertex_size=self.graphSquareSize.get(),
                    vertex_color=[self.phase_colours[phase] for phase in self.AH.vs["phase"]],
                    edge_color='Silver',
                )
                return(self.figure)
            else:
                ig.plot(
                    self.AH,
                    target=self.savePath,
                    layout=self.layout,
                    vertex_label=self.AH.vs["id"],
                    vertex_label_dist=self.graphLabelDist.get(),
                    vertex_label_angle=self.graphLabelAngle.get(),
                    vertex_shape='rectangle',
                    vertex_label_size=self.graphTextSize.get(),
                    vertex_size=self.graphSquareSize.get(),
                    vertex_color=[self.phase_colours[phase] for phase in self.AH.vs["phase"]],
                    edge_color='Silver',
                    margin= float(self.graphMarginSize.get())*200
                )
                return
        else:
            if display:
                self.layout.rotate(180)
                ig.plot(
                    self.AH,
                    target=self.axis,
                    layout=self.layout,
                    vertex_label=self.AH.vs["id"],
                    vertex_label_dist=self.graphLabelDist.get(),
                    vertex_label_angle=self.graphLabelAngle.get(),
                    vertex_shape='rectangle',
                    vertex_label_size=self.graphTextSize.get(),
                    vertex_size=self.AH.vs["size"],
                    vertex_color=[self.phase_colours[phase] for phase in self.AH.vs["phase"]],
                    edge_color='Silver',
                )
                return(self.figure)
            else:
                ig.plot(
                    self.AH,
                    target=self.savePath,
                    layout=self.layout,
                    vertex_label=self.AH.vs["id"],
                    vertex_label_dist=self.graphLabelDist.get(),
                    vertex_label_angle=self.graphLabelAngle.get(),
                    vertex_shape='rectangle',
                    vertex_label_size=self.graphTextSize.get(),
                    vertex_size=self.AH.vs["size"],
                    vertex_color=[self.phase_colours[phase] for phase in self.AH.vs["phase"]],
                    edge_color='Silver',
                    margin= float(self.graphMarginSize.get())*200
                )
                return
        
        
class MyWindow(ttk.Frame):
    def __init__(self,master):
        super().__init__(master)
        
        #Define the main interface layout and make it resizeable 
        self.grid(column=0,row=0,sticky=(N,W,E,S))
        self.rowconfigure(0, weight=0)
        self.columnconfigure(0, weight=1, minsize=200)
        self.columnconfigure(1, weight=0)

        #create frame to hold the interactive widgets
        self.buttonFrame = ttk.Frame(self) 
        self.buttonFrame.grid(column=0,row=0,sticky=(N,W,E,S))

        #Add frame for widgets
        #https://github.com/rdbende/Azure-ttk-theme
        self.AddNodeFrame = ttk.LabelFrame(self.buttonFrame, text='Add Node')
        self.AddNodeFrame.grid(column=0,row=1, pady=(10,10))
        
        #Create Add node Widgets
        ttk.Label(self.AddNodeFrame,text='Function:').grid(column=0,row=1)
        self.nodeFunction = StringVar()
        self.nodeFunctionCombo = ttk.Combobox(self.AddNodeFrame,textvariable=self.nodeFunction)
        self.nodeFunctionCombo['values'] = ['Purpose','Values','Functions','Processes','Physical']
        self.nodeFunctionCombo.grid(column=1,row=1)
        ttk.Label(self.AddNodeFrame,text='Name:').grid(column=0,row=3)
        self.nodeName = StringVar()
        ttk.Entry(self.AddNodeFrame,textvariable=self.nodeName).grid(column=1,row=3)
        ttk.Button(self.AddNodeFrame,text='Add', command=self.addNode).grid(column=1,row=4)
        
        #Add frame for widgets
        self.AddLinkFrame = ttk.LabelFrame(self.buttonFrame, text='Add Link')
        self.AddLinkFrame.grid(column=0,row=2, pady=(0,10))
        
        #Create Add Link Widgets
        ttk.Label(self.AddLinkFrame,text='Name:').grid(column=0,row=7)
        self.linkName = StringVar()
        self.linkFromNameCombo = ttk.Combobox(self.AddLinkFrame,textvariable=self.linkName)
        self.linkFromNameCombo.grid(column=1,row=7)
        ttk.Label(self.AddLinkFrame,text='Parent node:').grid(column=0,row=9)
        self.linkParents = StringVar()
        self.linkToNameCombo = ttk.Combobox(self.AddLinkFrame,textvariable=self.linkParents)
        self.linkToNameCombo.grid(column=1,row=9)
        ttk.Button(self.AddLinkFrame,text='Add', command=self.addLink).grid(column=1,row=10,pady=(0,10))
        
        #Add frame for delete widgets
        self.AddLinkFrame = ttk.LabelFrame(self.buttonFrame, text='Delete Node or Clear Links')
        self.AddLinkFrame.grid(column=0,row=3, pady=(0,10))
        
        #Create Add Delete Widgets
        ttk.Label(self.AddLinkFrame,text='Name:').grid(column=0,row=7)
        self.deleteOne = StringVar()
        self.deleteOneCombo = ttk.Combobox(self.AddLinkFrame,textvariable=self.deleteOne)
        self.deleteOneCombo.grid(column=1,row=7)
        ttk.Button(self.AddLinkFrame,text='Delete Node', command=self.deleteNode).grid(column=0,row=10,pady=(0,10),padx=(10,0))
        ttk.Button(self.AddLinkFrame,text='Clear All Links', command=self.deleteLinks).grid(column=1,row=10,pady=(0,10))

        #Add Misc widgets frame
        self.miscButtonFrame = ttk.LabelFrame(self.buttonFrame, text='Functions')
        self.miscButtonFrame.grid(column=0,row=4, pady=(0,10),padx=(10,10))
        
        #Add misc buttons
        ttk.Button(self.miscButtonFrame,text='Refresh Graph', command=self.resetGraph).grid(column=0,row=0)

        
        #Add TransiT image
        #https://tkdocs.com/tutorial/widgets.html
        ttk.Label(self.buttonFrame,image=self.master.transitImage,text='Funded By the TransiT Project',compound='left').grid(column=0,row=5, pady=(0,10),padx=(10,10))
        
        
        self.graphFrame = ttk.LabelFrame(self, text='Abstraction Heirarchy')
        self.graphFrame.grid(column=1,row=0, rowspan=11,sticky=(N,W,E,S),padx=5,pady=5)
        self.graphFrame.rowconfigure(0, weight=0)
        self.graphFrame.columnconfigure(0, weight=1, minsize=200)
        
        #https://matplotlib.org/stable/gallery/user_interfaces/embedding_in_tk_sgskip.html
        self.tempFigure= self.master.createAHGraph()
        self.graphFigure = FigureCanvasTkAgg(self.tempFigure, master=self.graphFrame)
        self.graphFigure.draw()

        #Draw Tk https://stackoverflow.com/questions/12913854/displaying-matplotlib-navigation-toolbar-in-tkinter-via-grid
        self.graphFigure.get_tk_widget().grid(column=0,row=0, padx=5,pady=5)
        
        
        
        ## Add commands to the menu bar, this is done here for inhertiance
        #Commands for Export
        self.master.appTopExport.add_command(label='Export Graph as PNG', command=self.exportPNG)
        self.master.appTopExport.add_command(label='Export Graph as PDF', command=self.exportPDF)
        self.master.appTopExport.add_command(label='Export Graph as GML', command=self.exportGML)

        #Commands for File
        self.master.appTopFile.add_command(label='New', command=self.newData)
        self.master.appTopFile.add_command(label='Open Excel File', command=self.openWindow)
        self.master.appTopFile.add_separator()
        self.master.appTopFile.add_command(label='Save Data to Excel', command=self.saveData)
        self.master.appTopFile.add_separator()
        self.master.appTopFile.add_command(label='Reset Graph Back to Saved Data', command=self.resetData)
        
        #Checkboxes for Statistics
        self.master.appTopStatistics.add_checkbutton(label='Eigenvector Centraility', variable=self.master.graphEigen, onvalue=1,offvalue=0)
        self.master.appTopStatistics.add_checkbutton(label='Closeness', variable=self.master.graphCloseness, onvalue=1,offvalue=0)
        self.master.appTopStatistics.add_checkbutton(label='Betweeness', variable=self.master.graphBetweeness, onvalue=1,offvalue=0)
        self.master.appTopStatistics.add_checkbutton(label='Google Pagerank', variable=self.master.graphPagerank, onvalue=1,offvalue=0)
        self.master.appTopStatistics.add_checkbutton(label='Strength', variable=self.master.graphStrength, onvalue=1,offvalue=0)
        self.master.appTopStatistics.add_separator()
        self.master.appTopStatistics.add_checkbutton(label='Add All Statistics to Excel', variable=self.master.addStatsToExcel, onvalue=1,offvalue=0)

    #code from https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
    def imagePath(path):
        try:
            formPath = sys._MEIPASS
        except Exception:
            formPath = os.path.abspath('.')
        return os.path.join(formPath,path)

    
    #Menu Function for opening a file
    def openWindow(self):
        self.master.openPath = filedialog.askopenfilename(title='Open a file',filetypes=[("Excel Files",'*.xlsx')])
        self.master.importData()
        self.resetGraph()
     
    #Menu Function for reseting data back to the saved file
    def resetData(self):
        self.master.importData()
        self.resetGraph()
    
    #Menu Function for creating a new file
    def newData(self):
        self.master.loadedData = pd.DataFrame(columns=['phase','id','parent'])
        self.resetGraph()
        
    #Menu Function for Saving Data 
    def saveData(self):      
        #Remove excess rows
        self.master.loadedData = self.master.loadedData[['phase','id','parent']]
        
        #Add statistics to dataframe
        if self.master.addStatsToExcel.get() == '1':
            self.master.loadedData.insert(3, 'eigenvector', self.master.AH.eigenvector_centrality())
            self.master.loadedData.insert(4, 'betweenness', self.master.AH.betweenness())
            self.master.loadedData.insert(5, 'closeness', self.master.AH.closeness())
            self.master.loadedData.insert(6, 'pagerank', self.master.AH.pagerank())
            self.master.loadedData.insert(7, 'strength', self.master.AH.strength())
        
        #Output the new dataframe to excel
        self.master.savePath = filedialog.asksaveasfilename(title='Save a file',filetypes=[("Excel Files",'*.xlsx')])
        if '.xlsx' not in self.master.savePath:
            self.master.savePath = self.master.savePath+'.xlsx'
        self.master.loadedData.to_excel(self.master.savePath)
        
        #Set the currently open path/file to be the newly saved path/file
        self.master.openPath = self.master.savePath
    
    #Menu Function for exporting AH Graph
    def exportPNG(self):
        self.master.savePath = filedialog.asksaveasfilename(title='Save a file',filetypes=[('PNG Files','*.png')])
        if '.png' not in self.master.savePath:
            self.master.savePath = self.master.savePath+'.png'
        self.master.createAHGraph(False)
    
    #Menu Function for exporting AH Graph
    def exportPDF(self):
        self.master.savePath = filedialog.asksaveasfilename(title='Save a file',filetypes=[('PDF Files','*.pdf')])
        if '.pdf' not in self.master.savePath:
            self.master.savePath = self.master.savePath+'.pdf'
        self.master.createAHGraph(False)
    
    #Menu Function for exporting AH Graph
    def exportGML(self):
        self.master.savePath = filedialog.asksaveasfilename(title='Save a file',filetypes=[('GML Files','*.gml')])
        if '.gml' not in self.master.savePath:
            self.master.savePath = self.master.savePath+'.gml'
        self.master.AH.write_gml(self.master.savePath)
    
    #Button Function for adding nodes
    def addNode(self):
        #Create a new temporary dataframe with the new value and then join this to the old dataframe
        #https://pandas.pydata.org/pandas-docs/version/1.4/reference/api/pandas.DataFrame.append.html#pandas.DataFrame.append
        tempDf = pd.DataFrame([[str(self.nodeFunction.get()),str(self.nodeName.get()),'']],columns=['phase','id','parent'])
        self.master.loadedData = (pd.concat([self.master.loadedData[['phase','id','parent']], tempDf])).reset_index(drop=True)
        self.resetGraph()
     
    #Button Function for adding links between nodes
    def addLink(self):
        #Create the link in the dataframe using the data from the combobox inputs
        for x in range(0,self.master.loadedData['id'].size):          
            if self.master.loadedData['id'].loc[x] == self.linkName.get():
                if self.master.loadedData['parent'].loc[x] == '':
                    self.master.loadedData['parent'].loc[x] = str(self.linkParents.get())
                else:
                    self.master.loadedData['parent'].loc[x] = str(self.master.loadedData['parent'].loc[x])+','+str(self.linkParents.get())
        
        self.resetGraph()
        
    #Button Function for deleting a node
    def deleteNode(self):
        for x in range(0,self.master.loadedData['id'].size):          
            if self.master.loadedData['id'].loc[x] == self.deleteOne.get():
                self.master.loadedData=self.master.loadedData.drop([x])
        self.resetGraph()
        
    #Button Function for clearing all links for a node
    def deleteLinks(self):
        for x in range(0,self.master.loadedData['id'].size):          
            if self.master.loadedData['id'].loc[x] == self.linkName.get():
                self.master.loadedData['parent'].loc[x] == ''
        self.resetGraph()

    #Function for updating comboboxes in the interfaces based on current dataframe  
    def updateBoxes(self):
        #update comboboxes in the interfaces based on current dataframe
        #https://tkdocs.com/tutorial/widgets.html#entry
        self.linkFromNameCombo['values'] = np.unique(self.master.loadedData['id'].to_numpy()).tolist()
        self.linkToNameCombo['values'] = np.unique(self.master.loadedData['id'].to_numpy()).tolist()
        self.deleteOneCombo['values'] = np.unique(self.master.loadedData['id'].to_numpy()).tolist()

    #A function to reload the data into the class after any changes and then draw a new graph. This also resets the options in the comboboxes
    def resetGraph(self):
        self.master.loadData()
        self.graphFigure.get_tk_widget().destroy()
        self.graphFigure = FigureCanvasTkAgg(self.master.createAHGraph(), master=self.graphFrame)
        self.graphFigure.draw()
        self.graphFigure.get_tk_widget().grid(column=0,row=0)
        self.updateBoxes()
           
        
main= MyApp()
main2 = MyWindow(main)
main2.mainloop()