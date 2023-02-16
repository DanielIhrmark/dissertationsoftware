# First Contact v.2
###NOTES:######
#API not OK according to GDPR (used local server). JLT will not be updated 
#Local server is unstable on low-performance devices


# Setup
from tkinter import filedialog
import tkinter as tk
from shutil import copyfile, rmtree
import os
import zipfile
from lxml import etree
import glob
import pandas as pd
import language_tool_python
import textstat
import re
import tkinter.scrolledtext as scrolledtext
import dash_core_components as DCC
import dash_html_components as DHC
import dash_table
import plotly.express as px
import dash
import webbrowser
from collections import Counter
import plotly.graph_objects as go
from threading import Timer


introductionPage = "https://docs.google.com/document/d/1tgwOU-GirS9_MBvhl7d_Fb10lgZiqrO5DWM5HHmDpVo/edit?usp=sharing"

#GUI main loop root frame
root = tk.Tk()
root.geometry("380x400+300+300")
root.title("First Contact v.1")

#Textbox for updating progress and print redirector
textframe = tk.Frame(root)
textframe.grid(row=8, column=2, sticky='w')
logbox = scrolledtext.ScrolledText(textframe, width=15, height=10, undo=True)
logbox.pack(expand=True, fill='both')
logboxLabel = tk.Label(root, text = "Processed files:")
logboxLabel.grid(row=7, column=2, sticky='w')

#Other interactions for root
#Display instructions
instructionsVar = tk.IntVar(value=1)
tk.Checkbutton(root, text="Display Instructions", variable=instructionsVar).grid(row=11, column=2)

#Var list for patterns
typosVar = tk.IntVar(value=1)
typographyVar = tk.IntVar()
grammarVar = tk.IntVar(value=1)
punctuationVar = tk.IntVar(value=1)
confusedVar = tk.IntVar()
casingVar = tk.IntVar()
nonstandardVar = tk.IntVar()
styleVar = tk.IntVar()
collocateVar = tk.IntVar()
miscVar = tk.IntVar()

def openNewWindow():
    #Pattern selection in new window
    # Creating a new window
    newWindow = tk.Toplevel(root)

    # sets the title of the new window
    newWindow.title("Pattern selection")

    # sets the geometry of new window
    newWindow.geometry("200x300")

    # A Label widget for title in window
    tk.Label(newWindow, text="Select patterns for analysis:").grid(row=2, sticky='w')

    #Creating checkboxes for pattern selection

    tk.Checkbutton(newWindow, text="TYPOS", variable=typosVar).grid(row=3, sticky='w')

    tk.Checkbutton(newWindow, text="TYPOGRAPHY", variable=typographyVar).grid(row=4, sticky='w')

    tk.Checkbutton(newWindow, text="GRAMMAR", variable=grammarVar).grid(row=5, sticky='w')

    tk.Checkbutton(newWindow, text="PUNCTUATION", variable=punctuationVar).grid(row=6, sticky='w')

    tk.Checkbutton(newWindow, text="CONFUSED_WORDS", variable=confusedVar).grid(row=7, sticky='w')

    tk.Checkbutton(newWindow, text="CASING", variable=casingVar).grid(row=8, sticky='w')

    tk.Checkbutton(newWindow, text="NONSTANDARD_PHRASE", variable=nonstandardVar).grid(row=9, sticky='w')

    tk.Checkbutton(newWindow, text="STYLE", variable=styleVar).grid(row=10, sticky='w')

    tk.Checkbutton(newWindow, text="COLLOCATIONS", variable=collocateVar).grid(row=11, sticky='w')

    tk.Checkbutton(newWindow, text="MISC", variable=miscVar).grid(row=12, sticky='w')

openPatternSelection = tk.Button(root,
             text ="Pattern Selection",
             command = openNewWindow).grid(row=11, column=0)
spacerBottom = tk.Label(root, text="")
spacerBottom.grid(row=10, column=0, sticky='w')


# Creates list of all .docx files in directory
def docxMorph():
    global Essays
    Essays = flist

    # Creates .txt versions of all .docx documents in directory
    for i in Essays:
        zip_dir = (i)
        zip_dir_zip_ext = os.path.splitext(zip_dir)[0] + '.zip'
        copyfile(zip_dir, zip_dir_zip_ext)
        zip_ref = zipfile.ZipFile(zip_dir_zip_ext, 'r')
        zip_ref.extractall('./temp')
        data = etree.parse('./temp/word/document.xml')

        result = [node.text.strip() for node in
                  data.xpath("//w:t", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})]

        with open(os.path.splitext(zip_dir)[0] + '.txt', 'wb') as txt:
            joined_result = '\n'.join(result).encode('UTF-8')
            txt.write(joined_result)
        zip_ref.close()
        rmtree('./temp')
        os.remove(zip_dir_zip_ext)

    # Source: https://stackoverflow.com/questions/44760366/how-do-i-write-a-python-script-that-can-read-doc-docx-files-and-convert-them-to

#Setting up dataframes to receive results of analysis
def langAnalysisSetup():
    # Create list of the .txt files created above
    def filebrowser2():
        return [f for f in glob.glob(selectedDir + "/*.txt")]

    global EssaysText
    EssaysText = filebrowser2()

    column_names = ['message', 'replacements', 'offset', 'context', 'sentence', 'category', 'ruleId', 'ruleIssueType', 'offsetInContext', 'errorLength', 'fileName']
    global df
    df = pd.DataFrame(columns=column_names)

    column_names = ['Sentence Count', 'Word Count', 'Complexity Score', 'fileName']
    global df2
    df2 = pd.DataFrame(columns=column_names)

#Function to carry out the actual analysis
def analysisEssay(df, df2):
    for essay in EssaysText:
        logbox.insert(tk.INSERT, re.sub(selectedDir,"", essay + '\n'))
        logbox.yview(tk.END)
        root.update_idletasks()
        fName = essay
        File = open(essay, encoding="UTF-8")  # open file
        lines = File.read()  # read all lines
        text = lines.replace("\n", " ")

        # Grammar and spelling check
        tool = language_tool_python.LanguageTool('en-US') #Needs to be tied into interface
        lines2 = re.sub(r" ?\([^)]+\)", "", lines)  # Removing in-text citations due to langcheck error
        hits = tool.check(lines2)  # Run error recognition
        for i in hits:
            new_row = {'message': i.message, 'replacements': i.replacements, 'offset': i.offset, 'context': i.context,
                       'sentence': i.sentence, 'category': i.category, 'ruleId': i.ruleId, 'ruleIssueType': i.ruleIssueType,
                       'offsetInContext': i.offsetInContext, 'errorLength': i.errorLength, 'fileName': os.path.basename(fName)}
            df = df.append(new_row, ignore_index=True)


        new_row2 = {'Sentence Count':textstat.sentence_count(text), 'Word Count':textstat.lexicon_count(text), 'Complexity Score':textstat.text_standard(text, float_output=True), 'fileName':os.path.basename(fName)}
        df2 = df2.append(new_row2, ignore_index=True)

    df.to_csv('OutputJLT.csv', index=False)
    df2.to_csv('OutputTextstat.csv', index=False)

#Enable user to select directory of student texts
def askdirectory():
  dirname = tk.filedialog.askdirectory()
  if dirname:
    var.set(dirname)

def UserFileInput(status,name):
  optionFrame = tk.Frame(root)
  optionLabel = tk.Label(optionFrame)
  optionLabel["text"] = name
  optionLabel.grid(row=2, column=0, sticky='w')
  text = status
  var = tk.StringVar(root)
  var.set(text)
  w = tk.Entry(optionFrame, textvariable= var)
  w.grid(row=3, column=0, sticky='w')
  optionFrame.grid(row=3, column=0, sticky='w')
  return w, var

#Update files shown
def Print_entry():
    lbox.delete('0', 'end')
    global selectedDir
    selectedDir = var.get()
    global flist
    flist = [f for f in glob.glob(selectedDir + "/*.docx")]
    for item in flist:
        item = re.sub(selectedDir,"", item)
        lbox.insert(tk.END, item)

#Buttons to select directory and load files
dirBut = tk.Button(root, text='Browse input', command = askdirectory)
dirBut.grid(row=4, column=0, sticky='w')
spacer = tk.Label(root, text="")
spacer.grid(row=5, column=0, sticky='w')
getBut = tk.Button(root, text='Update file list', command = Print_entry)
getBut.grid(row=6, column=0, sticky='w')


w, var = UserFileInput("", "Directory:")

#List of files in directory in GUI
lboxLabel = tk.Label(root, text = "Files in directory:")
lboxLabel.grid(row=7, column=0, sticky='w')

lbox = tk.Listbox(root)
lbox.grid(row=8,column=0, sticky='w')

# Main function starting sub-functions in order
def analysisFunction():
    docxMorph()
    langAnalysisSetup()
    analysisEssay(df, df2)
    for f in glob.glob(selectedDir + "/*.txt"):
        os.remove(f)
    logbox.insert(tk.INSERT,"Analysis \n")
    logbox.insert(tk.INSERT,"Completed!")
    logbox.yview(tk.END)

def reportGenerator():
    dfJLT = pd.read_csv('OutputJLT.csv')
    dfTextstat = pd.read_csv('OutputTextstat.csv')
    sampleList = dfTextstat['fileName'].to_list()
    port1 = 8050
    root.destroy()
    dfJLT.pop('offset')
    dfJLT.pop('offsetInContext')
    dfJLT.pop('errorLength')

    def open_browser(port):
        webbrowser.open_new("http://localhost:{}".format(port))

    def open_browser_URL(url):
        webbrowser.open_new(url)

    column_names = ['TYPOS','TYPOGRAPHY','GRAMMAR','PUNCTUATION','CONFUSED_WORDS','CASING','NONSTANDARD_PHRASE','STYLE','COLLOCATIONS','fileName']
    JLTCategories = pd.DataFrame(columns=column_names)

    for sample in sampleList:
        sampleDF = dfJLT.loc[dfJLT['fileName'] == sample]
        sampleCat = sampleDF['category'].to_list()
        sampleCatDict = Counter(sampleCat)
        catDict = dict(sampleCatDict)
        catDict['fileName'] = sample
        new_row = catDict
        JLTCategories = JLTCategories.append(new_row, ignore_index=True)

    #Cleaning according to pattern selection
    def purge(pattern):
        indexNames = dfJLT[dfJLT['category'] == pattern].index
        dfJLT.drop(indexNames, inplace=True)

    if typosVar.get() == 0:
        purge('TYPOS')

    if typographyVar.get() == 0:
        purge('TYPOGRAPHY')

    if grammarVar.get() == 0:
        purge('GRAMMAR')

    if punctuationVar.get() == 0:
        purge('PUNCTUATION')

    if confusedVar.get() == 0:
        purge('CONFUSED_WORDS')

    if casingVar.get() == 0:
        purge('CASING')

    if nonstandardVar.get() == 0:
        purge('NONSTANDARD_PHRASE')
        purge('NONSTANDARD_PHRASES')

    if styleVar.get() == 0:
        purge('STYLE')

    if collocateVar.get() == 0:
        purge('COLLOCATIONS')

    if miscVar.get() == 0:
        purge('MISC')

    JLTCounts = dfJLT['fileName'].value_counts().rename_axis('fileName').reset_index(name='counts')
    dfTextstat['Error Patterns'] = dfTextstat['fileName'].map(JLTCounts.set_index('fileName')['counts'])
    dfTextstat['AvgSentLength'] = round(dfTextstat['Word Count']/dfTextstat['Sentence Count'],2)
    dfTextstat['Errors per Sentence'] = round(dfTextstat['Error Patterns']/dfTextstat['Sentence Count'],2)
    totStats = pd.merge(dfTextstat, JLTCategories, on='fileName')

    #Dash
    # Style sheet
    external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

    #Figures
    owFig = px.scatter(totStats, x="Errors per Sentence", y="Complexity Score",
                     size="Word Count", color="fileName", hover_name="fileName",
                     log_x=True, size_max=20, title="Overview of Complexity Analysis and Error Patterns per Sentence")

    layout = {'title': {'text': 'Pattern Categories per Text'}}
    x = totStats['fileName'].to_list()
    owFigJLT = go.Figure(data=[go.Bar(
        name = 'Typos',
        x = x,
        y = totStats['TYPOS'].to_list()
   ),
                       go.Bar(
        name = 'Typography',
        x = x,
        y = totStats['TYPOGRAPHY'].to_list()
   ),
                       go.Bar(
        name = 'Grammar',
        x = x,
        y = totStats['GRAMMAR'].to_list()
                       ),
                       go.Bar(
        name = 'Punctuation',
        x = x,
        y = totStats['PUNCTUATION'].to_list()
                       ),
                       go.Bar(
        name = 'Confused Words',
        x = x,
        y = totStats['CONFUSED_WORDS'].to_list()
                       ),
                       go.Bar(
        name = 'Casing',
        x = x,
        y = totStats['CASING'].to_list()
                       ),
                       go.Bar(
        name = 'Non-standard Phrase',
        x = x,
        y = totStats['NONSTANDARD_PHRASE'].to_list()
                       ),
                       go.Bar(
        name = 'Style',
        x = x,
        y = totStats['STYLE'].to_list()
                       ),
                       go.Bar(
        name = 'Collocation',
        x = x,
        y = totStats['COLLOCATIONS'].to_list()
                       ),
                       go.Bar(
        name = 'Miscellaneous',
        x = x,
        y = totStats['MISC'].to_list()
                       )
        ], layout = layout)

    pivotDF = pd.melt(dfJLT, id_vars =['fileName', 'category'], value_vars =['ruleId'])
    pivotDF = pivotDF.groupby(pivotDF.columns.tolist()).size().reset_index(). \
        rename(columns={0: 'frequency'})

    owFigJLTCat = px.sunburst(pivotDF, path=['fileName', 'category', 'value'], values='frequency',
                  color='category', hover_data=['frequency'], title="Categories and Sub-categories in Error Patterns per Text")

    app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

    #For indivdual analysis text selection
    def generate_table(dataframe):
        #From CharmingData on youtube.
        return dash_table.DataTable(
            columns=[
                {"name": i, "id": i, "deletable": True, "selectable": True, "hideable": True}
                if i == "message" or i == "replacements" or i == "context" or i == "sentence"
                else {"name": i, "id": i, "deletable": True, "selectable": True}
                for i in dataframe.columns
            ],
            data=dataframe.to_dict('records'),  # the contents of the table
            export_format="xlsx",  # For trajectory and storage
            export_headers="display", #Table is exported as shown
            editable=True,  # allow editing of data inside all cells
            filter_action="native",  # allow filtering of data by user ('native') or not ('none')
            sort_action="native",  # enables data to be sorted per-column by user or not ('none')
            sort_mode="multi",  # sort across 'multi' or 'single' columns
            row_deletable=True,  # choose if user can delete a row (True) or not (False)
            page_action="native",  # all data is passed to the table up-front or not ('none')
            page_current=0,  # page number that user is on
            page_size=100,  # number of rows visible per page
            style_cell={  # ensure adequate header width when text is shorter than cell's text
                'minWidth': 95, 'maxWidth': 95, 'width': 95
            },
            style_cell_conditional=[  # align text columns to left. By default they are aligned to right
                {
                    'if': {'column_id': c},
                    'textAlign': 'left'
                } for c in ['message', 'replacements', 'context', 'sentence', 'category', 'ruleId', 'ruleIssueType', 'fileName']
            ],
            style_data={  # overflow cells' content into multiple lines
                'whiteSpace': 'normal',
                'height': 'auto'
            },
            style_header={  # overflow cells' content into multiple lines
                'whiteSpace': 'normal',
                'height': 'auto'
            }
        )

    #Overview
    app.layout = DHC.Div([
        DCC.Tabs([
            DCC.Tab(label='Overview', children=[
                # Page header
                DHC.H1(children='Overview'),

                # Graphs
                DHC.Div(children=""),
                DCC.Graph(
                id='scatGraph',
                figure=owFig
                ),
                DHC.Div(id='table-container2'),
                DCC.Graph(
                id='JLTCat',
                figure=owFigJLT
                ),
                DCC.Graph(
                id='treeGraph',
                figure=owFigJLTCat,
                style={'width': '120vh', 'height': '120vh'}
                ),
                DHC.Br(),
                DHC.Br(),
                DHC.Div(id='table-container3')
                ]),
            DCC.Tab(label='Individual Overview', children=[
                # Page header
                DHC.H1(children='Individual Data'),

                #Dropdown menu for individual texts
                DHC.H4(children='Uploaded texts:'),
                DCC.Dropdown(id='dropdown', options=[
                    {'label': i, 'value': i} for i in dfJLT.fileName.unique()
                ], multi=False, placeholder='Select text...'),
                DHC.Div(id='barGraphInd'),
                DHC.Br(),
                DHC.Br(),
                DHC.Div(id='table-container')
                        ])
                    ])
                ])

    @app.callback(
        dash.dependencies.Output('table-container', 'children'),
        [dash.dependencies.Input('dropdown', 'value')])
    def display_table(dropdown_value):
        if dropdown_value is None:
            return generate_table(dfJLT)
        else:
            dff = dfJLT[dfJLT.fileName.str.contains(dropdown_value)]
            return generate_table(dff)

    @app.callback(
        dash.dependencies.Output('table-container2', 'children'),
        [dash.dependencies.Input('dropdown', 'value')])
    def display_table(dropdown_value):
        if dropdown_value is None:
            return generate_table(dfTextstat)
        else:
            dff = dfTextstat[dfTextstat.fileName.str.contains(dropdown_value)]
            return generate_table(dff)

    @app.callback(
        dash.dependencies.Output('table-container3', 'children'),
        [dash.dependencies.Input('dropdown', 'value')])
    def display_table(dropdown_value):
        if dropdown_value is None:
            return generate_table(totStats)
        else:
            dff = totStats[totStats.fileName.str.contains(dropdown_value)]
            return generate_table(dff)

    @app.callback(
        dash.dependencies.Output('barGraphInd', 'children'),
        [dash.dependencies.Input('dropdown', 'value')])
    def display_table(dropdown_value):
        if dropdown_value is None:
            return [
                DCC.Graph(id='bar-chart',
                          figure=px.bar(
                              data_frame=totStats,
                              x="fileName",
                              y='Error Patterns',
                              title="JLT Hits in Texts"
                          ).update_layout(showlegend=True, xaxis={'categoryorder': 'total ascending'})
                          .update_traces()
                          )
            ]
        else:
            dff = pivotDF[pivotDF.fileName.str.contains(dropdown_value)]
            return [
                DCC.Graph(id='bar-chart',
                figure=px.sunburst(data_frame=dff, path=['category', 'value'], values='frequency',
                  color='category',
                    title="Categories and Sub-categories in Selected Text", width=800, height=800)
                          )
            ]



    # Running the Dash app
    if __name__ == '__main__':
        Timer(1, open_browser(port1)).start();
        #If checkbox indicates instructions are desired
        if instructionsVar.get() == 1:
            Timer(2, open_browser_URL(introductionPage)).start();
        app.run_server(debug=False, use_reloader=False)



#Buttons for analysis and Dash
AnalysisBut = tk.Button(root, text='Analyze texts', command = analysisFunction)
AnalysisBut.grid(row=4, column=2, sticky='w')
spacer = tk.Label(root, text="")
ReportBut = tk.Button(root, text='Generate report', command = reportGenerator)
ReportBut.grid(row=6, column=2, sticky='w')



root.mainloop()

