##################################
#  XLKitLearn                    #
#  (C) Daniel Guetta, 2020       #
#      daniel@guetta.com         #
#      guetta@gsb.columbia.edu   #
#  Version 10.25                 #
##################################

# =====================
# =   Load Packages   =
# =====================

# Interaction with Excel
import xlwings as xw

import pdb
# Basic packages
import pandas as pd
import numpy as np
import scipy as sp
import itertools
import time
import warnings
import keyword
import os
import signal
from collections import OrderedDict
import traceback
import sys

# Data preparation utilities
import patsy as pt
import sklearn.feature_extraction.text as f_e
import nltk

# Sparse matrices
from scipy import sparse

# Estimators, learners, etc...
warnings.filterwarnings("ignore", category = DeprecationWarning)
from sklearn.linear_model import LinearRegression, LogisticRegression, Lasso
from sklearn.neighbors import KNeighborsRegressor, KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor

# Statsmodels learners for p-values
import statsmodels.api as sm

# Validation utilities and metrics
from sklearn import model_selection as sk_ms
from sklearn.metrics import r2_score, roc_auc_score, roc_curve, make_scorer
import sklearn.inspection as sk_i

# Plotting (ensure we don't use qt)
import matplotlib as mpl
mpl.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import json

# Verification
from datetime import datetime, timedelta
import hashlib
import requests

# =================
# =   Utilities   =
# =================

class AddinError(Exception):
    pass
    
def trim_edges(x):
    '''
    This function returns the input string without its first and last character
    '''
    return x[1:-1]
    
def remove_dupes(x):
    '''
    This function removes duplicates from the list passed to it while preserving
    the order of elements
    '''
    
    seen = set()
    return [i for i in x if not (i in seen or seen.add(i))]

def listify(*args):
    '''
    For each argument, if it is a list, it will be returned as-is. If not, a single-
    item list containing that item will be returned.
    '''
    
    out = tuple(i if type(i) is list else [i] for i in args)
    
    if len(out) == 1:
        return out[0]
    else:
        return out

def round_to_string(x,n):
    '''
    This function will take a number x. If it is an integer, it will convert it
    as a string. If not, it will round it to n decimal places and convert it to
    a string.
    '''
    
    if x == int(x):
        return str(int(x))
    else:
        return str(round(x,n))

def wrap_line(text, max_width):
    '''
    This function will take a long line of text and wrap it to be no
    wider than max_width
    '''
    
    text = [i.split(' ') for i in text.split('\n')]
    out = []
    
    for cur_row in text:
        buffer = ''
        for cur_word in cur_row:
            if len(buffer + ' ' + cur_word) > max_width:
                out.append(buffer)
                buffer = ''
            buffer = buffer + ' ' + cur_word
        
        out.append(buffer)
    
    return '\n'.join(out)
        
def pad_string(string, size, sep):
    '''
    This function will take a string, and pad it with the sep character to make
    it size-long
    '''
    
    return  string + ' ' + (sep*(size-len(string))) + ' '
        
        
class D(OrderedDict):
    '''
    This class extends the standard Python dictionary class in the following ways
      1. It allows elements to be accessed and set as attributes:
            - d.element will be equivalent to d.get('element', default=None)
            - d.element=value will be equivalent to d['element']=value if the
              dictionary has an alement called 'element'. If not, an
              AttributeError is thrown
      2. It allows each entry to also have an element called english_key, which
         contains a more verbose version of the key (in English, for example).
         This can be used in two ways
            - d.key(x) returns the key for the element with english_name x. If
              x does not exist, or multiple elements have english_name x, an
              AttributeError is thrown
            - d.english_keys() returns a list of all English keys for the
              dictionary - it is the equivalent of d.key()
      3. It provides an iterator, d.zip_entries() that allows us to iterate through
         (key, value) tuples
              
    '''

    def __getattr__(self, item):
        if item in self:
            return self[item]
        raise AttributeError(item)

    def __setattr__(self, key, value):
        if key in self:
            self[key] = value
            return
        raise AttributeError(key)

    def zip_entries(self):
        '''
        Iterates through (key, value) tuples in the dictionary
        '''
        
        return zip(self.keys(), self.values())

    def key(self, english_key):
        '''
        Given an english_key, this function will return the corresponding key
        '''
    
        matching_keys = [i for i in self if self[i].english_key == english_key]
        if len(matching_keys) == 1:
            return matching_keys[0]
        raise AttributeError(key)

    def english_keys(self):
        '''
        Will return a list of English keys in this dictionary
        '''
        
        return [self[i].english_key for i in self if 'english_key' in self[i]]

def levenshtein_ratio_and_distance(s, t, ratio_calc = False):
    """ levenshtein_ratio_and_distance:
        Calculates levenshtein distance between two strings.
        If ratio_calc = True, the function computes the
        levenshtein distance ratio of similarity between two strings
        For all i and j, distance[i,j] will contain the Levenshtein
        distance between the first i characters of s and the
        first j characters of t
    """
    
    # From https://www.datacamp.com/community/tutorials/fuzzy-string-python
    
    if len(s) == 0: return len(t)
    if len(t) == 0: return len(s)
    
    # Initialize matrix of zeros
    rows = len(s)+1
    cols = len(t)+1
    distance = np.zeros((rows,cols),dtype = int)

    # Populate matrix of zeros with the indeces of each character of both strings
    for i in range(1, rows):
        for k in range(1,cols):
            distance[i][0] = i
            distance[0][k] = k

    # Iterate over the matrix to compute the cost of deletions,insertions and/or substitutions    
    for col in range(1, cols):
        for row in range(1, rows):
            if s[row-1] == t[col-1]:
                cost = 0 # If the characters are the same in the two strings in a given position [i,j] then the cost is 0
            else:
                # In order to align the results with those of the Python Levenshtein package, if we choose to calculate the ratio
                # the cost of a substitution is 2. If we calculate just distance, then the cost of a substitution is 1.
                if ratio_calc == True:
                    cost = 2
                else:
                    cost = 1
            distance[row][col] = min(distance[row-1][col] + 1,      # Cost of deletions
                                 distance[row][col-1] + 1,          # Cost of insertions
                                 distance[row-1][col-1] + cost)     # Cost of substitutions
    if ratio_calc == True:
        # Computation of the Levenshtein Distance Ratio
        Ratio = ((len(s)+len(t)) - distance[row][col]) / (len(s)+len(t))
        return Ratio
    else:
        # print(distance) # Uncomment if you want to see the matrix showing how the algorithm computes the cost of deletions,
        # insertions and/or substitutions
        # This is the minimum number of edits needed to convert string a to string b
        return distance[row][col]

# ============================
# =   Set global constants   =
# ============================


EXCEL_INTERFACE = D(     interface_sheet = 'Add-in',
                     settings_dict_comma = '`',
                     settings_dict_colon = '|',
                           list_splitter = '&',
                            version_cell = 'B3',
                              email_cell = 'F17',
                            run_id_sheet = 'code_text',
                             run_id_cell = 'B1',
                                pid_cell = 'C1',
                               path_cell = 'D1',
                     graph_line_per_inch = 3,
                            output_width = 116,
                       output_code_width = 82,
                          max_table_rows = 8000,
               variable_importance_limit = 100)

PREDICTIVE_CONFIG = D(     settings_cell = 'D9',
                             status_cell = 'F7',
                             english_key = 'predictive_addin',
                       expected_settings = D( model = D(english_key='Model name'),
                                            formula = D(english_key='Formula'),
                                             param1 = D(default=''),
                                             param2 = D(default=''),
                                             param3 = D(default=''),
                                      training_data = D(english_key='Training data', kind='r'),
                                                  K = D(english_key='K', kind='i', default=None),
                                            ts_data = D(english_key='Is data time series data?', kind='b'),
                                    evaluation_perc = D(english_key='Size of evaluation set', kind='i', default=None),
                                    evaluation_data = D(english_key='Evaluation data', kind='r', default=None),
                                    prediction_data = D(english_key='Prediction data', kind='r', default=None),
                                               seed = D(english_key='Seed', kind='i', default=123),
                                       output_model = D(english_key='Output the model?', kind='b'),
                          output_evaluation_details = D(english_key='Output evaluation details?', kind='b'),
                                        output_code = D(english_key='Output code?', kind='b') ) )
                                        

TEXT_CONFIG = D( settings_cell = 'D14',
                   status_cell = 'F12',
                   english_key = 'text_addin',
             expected_settings = D( source_data = D(english_key='Source data'),
                                         max_df = D(english_key='Upper limiting frequency', default=1.0, kind='f'),
                                         min_df = D(english_key='Lower limiting frequency', default=1, kind='f'),
                                   max_features = D(english_key='Maximum words', kind='i'),
                                     stop_words = D(english_key='Remove stop words?', kind='b'),
                                         tf_idf = D(english_key='TF-IDF?', kind='b'),
                                     lda_topics = D(english_key='Number of LDA topics', default=None, kind='i'),
                                           seed = D(english_key='Seed', kind='i'),
                                      eval_perc = D(english_key='Evaluation percentage', default=0, kind='i'),
                                        bigrams = D(english_key='Include bigrams?', kind='b'),
                                           stem = D(english_key='Stem words?', kind='b'),
                                    output_code = D(english_key='Output code?', kind='b'),
                                  sparse_output = D(english_key='Sparse output?', kind='b'),
                                   max_lda_iter = D(english_key='LDA iterations', default=10, kind='i') ) )

MAX_LR_ITERS = 500

LINEAR_REGRESSION = 'lr'
NEAREST_NEIGHBORS = 'knn'
DECISION_TREE = 'dt'
BOOSTED_DT = 'bdt'
RANDOM_FOREST = 'rf'

MODELS = D( {LINEAR_REGRESSION : D( english_key = 'Linear/logistic regression',
                                        params = D(param1 = D(english_key='Lasso penalty',
                                                              sklearn_name='alpha',
                                                              kind='f',
                                                              list_default=0))),
            NEAREST_NEIGHBORS : D( english_key = 'K-Nearest Neighbors',
                                        params = D(param1 = D(english_key='Neighbors',
                                                              sklearn_name='n_neighbors',
                                                              kind='i',
                                                              list_default=0),
                                                   param2 = D(english_key='Weighting',
                                                              sklearn_name='weights',
                                                              kind='s',
                                                              list_default="uniform"))),
            DECISION_TREE     : D( english_key = 'Decision tree',
                                        params = D(param1 = D(english_key='Tree depth',
                                                              sklearn_name='max_depth',
                                                              kind='i')),
                                  import_logic = D(subpackage='tree', class_name = 'DecisionTree')),
            BOOSTED_DT        : D( english_key = 'Boosted decision tree',
                                        params = D(param1 = D(english_key='Tree depth',
                                                              sklearn_name='max_depth',
                                                              kind='i'),
                                                   param2 = D(english_key='Max trees',
                                                              sklearn_name='n_estimators',
                                                              kind='i'),
                                                   param3 = D(english_key='Learning rate',
                                                              sklearn_name='learning_rate',
                                                              kind='f',
                                                              list_default=0.1)),
                                  import_logic = D(subpackage='ensemble', class_name = 'GradientBoosting') ),
            RANDOM_FOREST     : D( english_key = 'Random forest',
                                        params = D(param1 = D(english_key='Tree depth',
                                                              sklearn_name='max_depth',
                                                              kind='i'),
                                                   param2 = D(english_key='Number of trees',
                                                              sklearn_name='n_estimators',
                                                              kind='i')),
                                  import_logic = D(subpackage='ensemble', class_name = 'RandomForest'))})

# ========================
# =   Excel Connectors   =
# ========================

class ExcelConnector:
    '''
    This class mediates the connection to Excel via xlwings. It exposes a single
    attribute, the xlwings workbook object, called wb.
    
    It should be initiated with a single parameter (workbook). If None or omitted,
    xlwings.Book.caller() will be used. If not, it will connect to workbook
    '''

    def __init__(self, workbook=None):
        if workbook is None:
            self._wb = xw.Book.caller()
        else:
            self._wb = xw.Book(workbook)
    
    @property
    def wb(self):
        return self._wb

class ExcelOutput:
    '''
    This class handles the output of data to the Excel spreadsheet.
    '''

    def __init__(self, sheet, excel_connector):
        # Prepare a variable to contain our output
        self._out = []
        
        # Prepare a variable to contain formatting
        # output
        self._formatting_fields = ["font_medium", "font_large", "bottom_thick",
                                    "top_thin", "italics", "bold", "align_center",
                                    "align_right", "expand", "courier", "align_left", "number_format"]
        self._formatting_data = {i:[] for i in self._formatting_fields}
        
        # Prepare graph formatting data and start with a graph number of 0
        self._graph_formatting = []
        self._graph_number = 0
        self._graphs = {}
        
        # Prepare a variable to hold the current indentation level in the output
        self._cur_indent = 0
        
        # Store the sheet name, Excel interface, and formatting macro for
        # post-processing
        self._wb = excel_connector.wb
        self._sheet = self._wb.sheets(sheet)
        self._format_sheet = self._wb.macro("format_sheet")
        
    @staticmethod
    def _col_letter(col_num):
        '''
        This function takes a numeric column number, and returns the
        letter corresponding to that excel column (eg 1 = A, 2 = B, etc...)
        '''
        
        if col_num <= 26:
            return chr(64 + col_num)
        elif col_num <= 27*26:
            first_letter = chr( 64 + int(np.floor( (col_num - 1)/26)) )
            return first_letter + ExcelOutput._col_letter( ( (col_num-1) % 26) + 1 )
        else:
            first_letter = chr( 64 + int(np.floor( (col_num-1) /(26*26))) )
            return first_letter + ExcelOutput._col_letter( (col_num-1) % (26*26) + 1 )
    
    def _determine_indent(self, indent_level):
        '''
        Various functions in this class require an indent as an argument, which can
        be provided in one of three forms
          - If a blank argument is provided, the current indentation level is used
          - If an integer is provided, that integer is used as the indentation level
          - If a string is provided, that string is converted to an integer and added
            to the indentation level
        This function takes this argument and returns the indentation level
        '''
        
        if indent_level == "":
            return self._cur_indent
        elif isinstance(indent_level, int):
            return indent_level
        else:
            return self._cur_indent + int(indent_level)
    
    def add_header(self, text, level, content = ""):
        '''
        This function will create a new header in the Excel spreadsheet, and increment
        the indent level
        
        It takes the following argument
          - text : the text of the header
          - level : the format of the header; 0 for title, 1 for subheading, and 2
                    for a sub-sub-heading with content
          - content : when level = 2, the content of the subheading
        '''
        
        if level == 0:
            self.add_row( [text], [ ["font_large", "bold"] ], indent_level = 0 )
            self.add_blank_row()
            self._cur_indent = 1
        elif level == 1:
            self.add_row( [text], [ ["font_medium", "bold"] ], indent_level = 1 )
            self.add_blank_row()
            self._cur_indent = 2
        elif level == 2:
            self.add_row( [ f"{text}{':' if text != '' else ''}",              str(content) ],
                          [ ["bold", "align_right"], "align_left" ],
                          indent_level = self._cur_indent )
        else:
            raise
    
    def add_blank_row(self):
        '''
        Add a blank row to the output
        '''
        
        self.add_row( [ '' ] )
    
    def add_row(self, content, format = [], indent_level = "", split_newlines = False):
        '''
        This function will add a single row to our Excel output. It takes the
        following arguments
          - content : EITHER a single string, to be output to a single cell
                          OR a list, containing a row to be output
          - format : The format of the output.
                      EITHER a single string, containing formatting that will be
                             applied to every cell output
                          OR a list of formatting instructions. The first element will
                             be applied to the first column, the second to the second,
                             etc... Each of these instructions can either be lists, for
                             multiple formatting instructions, or strings
                     Any instruction can be None or ''
          - indent_level: the indentation level. If it is an integer, that integer
                          is used as an indentation level. If it is a string, the
                          number in that string is added to the current indentation
                          level (self.cur_indent)
          - split_newlines: If True, the function will check whether any element
                            in content contains a newline. If it does, it will
                            split those and print them in multiple Excel cells.
                            If this happens, each of the resulting rows will have the
                            same format
        
        Note that elements in the format list are always assumed to refer to different
        columns. So for example, if
            content = 'hello'
        and
            format = ['bold', 'align_right', 'bold']
        The function will interpret these three elements are referring to three columns,
        and only 'bold' will be applied to the output.
        '''
        
        # If the content are a single string, stick them in a list
        content = listify(content)
        
        # If the format is a single string, apply it to every column
        if type(format) is not list:
            format = [format]*len(content)
        
        # Find the indentation level
        indent_level = self._determine_indent(indent_level)
        
        if split_newlines:
            # Split the columns
            content = [i.split('\n') for i in content]
            
            # Transpose the results
            n_rows = max([len(i) for i in content])
            content = [[col[row] if len(col) > row else ''
                            for col in content] for row in range(n_rows)]
            self.add_rows(content, format, indent_level)
        else:
            # Add the data, with the appropriate indent level
            self._out.append([""]*indent_level + content)
            
            # Handle formatting; find the cell reference for each cell, and appending it
            # to the formatting dictionary
            for col, col_format in enumerate(format):
                cell_reference = self._col_letter(indent_level + col + 1) + str(len(self._out))
                for cur_format in filter(None, listify(col_format)):
                    self._formatting_data[cur_format].append(cell_reference)
                                                                    
    def add_rows(self, content, format = [], indent_level = ""):
        '''
        This function will add a multiple rows to our Excel output. Each row must
        have the same format.
        
        It takes the following arguments
          - content : a two-dimensional list; each element is a list corresponding
                      to one row, with each element corresponding to one column
          - format: a list containing as many entries as columns in content. Each
                    element in the list can either be
                      - A string, for a single formatting instruction for that column
                      - A list, for multiple formatting instructions for that column
                      - None or '', for no formatting instructions
          - indent_level : the indentation level. If it is an integer, that integer
                           is used as an indentation level. If it is a string, the
                           number in that string is added to the current indentation
                           level (self.cur_indent)

        The benefit of using this function instead of multiple add_row calls is that
        the formatting instructions format each of the columns as one - this can
        result in more efficient macro post-processing in Excel
        '''
        
         # Find the indentation level
        indent_level = self._determine_indent(indent_level)
        
        # Add the content
        self._out.extend([[""]*indent_level + i for i in content])
        
        # Push the format
        first_row = len(self._out) - len(content) + 1
        last_row = len(self._out)
        
        for col, col_format in enumerate(format):        
            col_letter = self._col_letter(indent_level + col + 1)
            col_range = f'{col_letter}{first_row}:{col_letter}{last_row}'
            for cur_format in filter(None, listify(col_format)):
                self._formatting_data[cur_format].append(col_range)
    
    def add_table(self, df, three_dp = True, indent_level = "", alt_message=''):
        '''
        This function will output a Pandas table to Excel, formatting it in a
        standard format
        
        It takes the following arguments
          - df : the dataframe in question
          - three_dp : whether the table should be formatted to three decimal places.
                       If not, the numbers will be output as-is. Options are
                         - True : whole table formatting to 3dp
                         - False : none of the table
                         - -1 : only the last column to 3 dp
          - indent_level : the indentation level. If it is an integer, that integer
                           is used as an indentation level. If it is a string, the
                           number in that string is added to the current indentation
                           level (self.cur_indent)
        '''
        
        if len(df) > EXCEL_INTERFACE.max_table_rows:
            file_path = self._wb.sheets('code_text').range(EXCEL_INTERFACE.path_cell).value
            delim = '/' if '/' in file_path else '\\'
            file_path = file_path + delim
            
            file_name = 'file_' + str(int(np.random.uniform(0, 99999999))) + '.csv'
            
            self.add_row(f'This table has more than {EXCEL_INTERFACE.max_table_rows} rows. It will '
                            f'not be printed. I\'ve saved it to a file called {file_name} instead.')
            df.to_csv(file_path + file_name, index=False)
            self.add_blank_row()
            return False
        
        # Find the indentation level
        indent_level = self._determine_indent(indent_level)
        
        # Add the data to our output
        table_data = ( [ [""]*indent_level + df.columns.tolist() ] +
                        [ [""]*indent_level + i.tolist() for i in df.values ] )
        self._out = self._out + table_data
        
        # Find the top, bottom, left, and right coordinates of our table
        bottom_row  = len(self._out)
        top_row     = bottom_row - len(df)
        left_col,   right_col = self._col_letter(indent_level+1), self._col_letter(indent_level+len(df.columns))
        
        # Expand and center every cell
        for f in ['align_center', 'expand']:
            self._formatting_data[f].append(f'{left_col}{top_row}:{right_col}{bottom_row}')
                
        # Make the headers bold and add lines above and below
        for f in ['bottom_thick', 'top_thin', 'bold']:
            self._formatting_data[f].append(f'{left_col}{top_row}:{right_col}{top_row}')
        
        # Add a line to the bottom of the table
        self._formatting_data["bottom_thick"].append(f'{left_col}{bottom_row}:{right_col}{bottom_row}')
               
        # Italicize the first column (minus the header)
        for f in ["italics", "align_center"]:
            self._formatting_data[f].append(f'{left_col}{top_row+1}:{left_col}{bottom_row}')
        
        # Format the body of the table as 3 d.p. numbers if needed. Rather than
        # doing each column individually, try and format contiguous ranges all
        # at once
        start_number_format = None
        for col_n, col in enumerate(list(df.columns) + [None]):
            try:
                if (col is not None) and (df[col].apply(lambda x : len(str(float(x)).split('.')[1]) >= 4).sum() > 0):
                    if (start_number_format is None): start_number_format = col_n
                else:
                    raise
            except:
                if start_number_format is not None:
                    start_col = self._col_letter(indent_level+1+start_number_format)
                    end_col = self._col_letter(indent_level+1+col_n - 1)
                    self._formatting_data['number_format'].append(f'{start_col}{top_row+1}:{end_col}{bottom_row}')
                    start_number_format = None
        
        return True
        
    def add_graph(self, fig, indent_level = "", manual_shift=None):
        '''
        This function adds a graph to our Excel spreadsheet.
        
        It takes the following arguments
          - graph : a matplotlib figure
          - indent_level : the indentation level. If it is an integer, that integer
                           is used as an indentation level. If it is a string, the
                           number in that string is added to the current indentation
                           level (self.cur_indent)
        '''
        
        # Find the indentation level
        indent_level = self._determine_indent(indent_level)
        
        # Get the graph name
        self._graph_number += 1
        graph_name = f"g_{self._graph_number}"
        
        # Add the graph to the spreadsheet
        self._graphs[graph_name] = fig
        
        # Get the figure dimensions
        size = fig.get_size_inches()
        
        top_row = len(self._out)+1
        if manual_shift is None:
            # Skip 5 rows per inch, plus one extra
            self._out.extend([[]]*int(np.ceil(size[1]*EXCEL_INTERFACE['graph_line_per_inch'])+1))
        
            # Add the graph's formating data
            self._graph_formatting.append(f'{graph_name},{self._col_letter(indent_level+1)}{top_row},{size[0]},{size[1]}')
        else:
            self._graph_formatting.append(f'{graph_name},{self._col_letter(indent_level+1+manual_shift[0])}{top_row+manual_shift[1]},{size[0]},{size[1]}')
            
    def _get_output_array(self):
        '''
        This function combines the formatting and data to produce a final output
        string to Excel
        '''
        
        # Begin by printing out formatting cells in sequence
        out = [[",".join(self._formatting_data[i])] for i in self._formatting_fields]
        
        # Add the number of rows and columns
        out.extend([[max([len(i) for i in self._out])], [len(self._out)]])

        # Add the graph formatting data
        out.append(["|".join(self._graph_formatting)])
        
        # Add the remaining output
        out.extend(self._out)
        
        # Pad each row with spaces, to ensure each row has the same length
        max_cols = max([len(i) for i in out])

        out = [i + ['']*(max_cols - len(i)) for i in out]
        
        return out

    def _output_to_spreadsheet(self):
        '''
        This function will output all data in this object to Excel, and format the
        data as required.
        
        It returns the number of the last row output to the Excel spreadsheet. This
        can be used by the calling class to print out additional information after the
        write if needed
        '''
        
        # Get the output array
        out_array = self._get_output_array()
        
        # If any cells say true or false, add a quote before them
        for row_n, row in enumerate(out_array):
            for cell_n, cell in enumerate(row):
                try:
                    if 'true' in cell.lower() or 'false' in cell.lower():
                        out_array[row_n][cell_n] = "'" + out_array[row_n][cell_n]
                except:
                    pass
        
        # Output in chunks of 500 rows; xlwings can die if too much data is
        # output at once
        cur_row = 1
        row_chunk = 500
        while cur_row <= len(out_array):
            self._sheet.range("A" + str(cur_row)).value = out_array[(cur_row - 1):(cur_row - 1 + row_chunk)]
            cur_row = cur_row + row_chunk
        
        # Display the graphs
        for graph in self._graphs:
            self._sheet.pictures.add(self._graphs[graph], name=graph, update=True)
        
        # Close all graphics we might have opened
        plt.close('all')
        
        # Return the number of rows output to the spreadsheet
        return len(out_array)
    
class AddinOutput(ExcelOutput):
    '''
    This class inherits from ExcelOutput and handles the output of model
    results to Excel
    
    In addition to standard output, it will also output various runtime
    statistics
    '''

    def __init__(self, sheet, excel_connector):
        # Initialize the parent object
        ExcelOutput.__init__(self, sheet, excel_connector)
        
        # Keep track of the start time
        self._start_time = time.time()
        
        # Keep track of any events whose time needs to be benchmarked
        self._events = []
        self._event_times = []
      
    def log_event(self, name):
        '''
        This function will log an event, and keep track of when it happened
        
        It takes one argument - the name of the event
        '''
        
        self._events.append(name)
        self._event_times.append(time.time())
    
    def finalize(self, settings_string):
        '''
        This function will finalize the output of the model to Excel, and append
        time profiling statistics
        '''
        
        self.add_header("Technical Details", 1)
        
        # Print the settings string
        self.add_header("Settings", 2, settings_string)
        
        # Print the version number
        self.add_header("Add-in version #", 2, self._wb
                                                   .sheets(EXCEL_INTERFACE.interface_sheet)
                                                   .range(EXCEL_INTERFACE.version_cell)
                                                   .value[1:]
                                                   .split()[1])
        
        # Add a blank row and print the time profiles
        self.add_blank_row()
        
        times = [i-j for i,j in zip(self._event_times, [self._start_time] + self._event_times[:-1])]
        for event, t in zip(self._events, times):
            self.add_header( event, 2, str(round(t, 2)) )
            
        # Create a row for the write time (we'll modify this below once we've
        # written to Excel)
        self.add_header( "Write time", 2, 1 )
        
        # Create a row for the overhead time with the sum of all times so far. This
        # will be modified by the Excel macro
        self.add_header( "Overhead time", 2, self._event_times[-1] - self._start_time )
        
        # Output to the spreadsheet, and save the bottom row
        bottom_row = self._output_to_spreadsheet()
        
        # Modify the write time
        self._sheet.range(f'D{bottom_row-1}').value = str(round(time.time() - self._event_times[-1], 2))
        
        # Format the sheet
        self._format_sheet()

class AddinErrorOutput(ExcelOutput):
    '''
    This class inherits from ExcelOutput and handles the output of model
    errors to Excel
    
    It exposes the following additional function
      - add_error: this function add an error to the error report. If the
                   optional critical argument is True, an error is raised
                   to return to Excel. Otherwise, it just continues
      - add_error_category: adds an header for the category of error
    '''

    def __init__(self, sheet, excel_connector, source):
        # Initialize the parent object
        ExcelOutput.__init__(self, sheet, excel_connector)
        
        # Print the title
        self.add_header( "Add-in Error Report", 0 )
        
        # Create a flag to log whether we've recorded any error
        self._has_error = False
        
        # Create a flag to log whether we've recorded any error since the
        # last header. This is so that if we add a new header with no errors
        # logged under the previous one, we can print "No errors recorded in
        # this section"
        self._error_since_last_header = True
        
        # Store the source
        self._source = source
    
    def add_error(self, text, critical=False):
        '''
        Add an error to the report. If critical is True, an error is raised
        to return to Excel
        '''
        
        bullet = chr(8594)
        
        text = wrap_line(text, EXCEL_INTERFACE.output_width)
        
        self.add_row([bullet, text],
                        [['align_right'], ['courier']],
                        indent_level='-1',
                        split_newlines=True)
        
        self._has_error = True
        self._error_since_last_header = True

        if critical:
            self.finalize()
    
    def add_error_category(self, text):
        '''
        Add a new error category. If no error was printed since the last category
        was pushed, output "No errors recorded in this section"
        '''
                
        if not self._error_since_last_header:
            self.add_row( ["No error recorded in this section"], ["italics"] )
            
        self.add_blank_row()
        
        self.add_header(text, level = 1)     

        self._error_since_last_header = False
    
    def finalize(self):
        '''
        If self.has_error is true, this function outputs the error report to Excel,
        and raises a AddinError. If not, it does nothing. This allows us to check - at
        various points in our run - whether there is an error, end if there is, and
        continue if not.
        '''
        
        if self._has_error: 
            self.add_error_category('')
            
            # Send the error to the server
            try:
                run_id = self._wb.sheets(EXCEL_INTERFACE.run_id_sheet).range(EXCEL_INTERFACE.run_id_cell).value
            except:
                run_id = 'unknown'
                
            try:
                requests.post(url = 'http://guetta.org/addin/error.php',
                                    headers={'User-Agent': 'XY'},
                                    data = {'run_id':run_id , 'source':self._source , 'error_type':'caught_exception', 'error_text': str(self._get_output_array()), 'platform':os.name},
                                    timeout = 10 ) 
            except:
                pass
            
            # Output to the spreadsheet
            self._output_to_spreadsheet()
        
            # Run the formatting macro
            self._format_sheet()
            
            raise AddinError

# =======================
# =   Addin instances   =
# ======================= 

class AddinInstance:

    def __init__(self, excel_connector, out_err, config, udf_server):
        
        self._out_err = out_err
        self._wb = excel_connector.wb
        self._model_sheet = excel_connector.wb.sheets(EXCEL_INTERFACE.interface_sheet)
        
        self._status_cell = config.status_cell
        self._settings_cell = config.settings_cell
        self._expected_settings = config.expected_settings
        
        self._start_time = time.time()
        self._udf_server = udf_server
    
    def log_run(self):
        self._v_message = ''
        
        try:
            # Get the registered email
            try:
                reg_email = self._model_sheet.range(EXCEL_INTERFACE.email_cell).value 
            except:
                reg_email = 'unknown'
            
            # Get version number
            try:
                v_number = self._model_sheet.range(EXCEL_INTERFACE.version_cell).value[1:].split()[1]
            except:
                v_number = 'unknown'
            
            # Get run ID
            try:
                run_id = self._wb.sheets(EXCEL_INTERFACE.run_id_sheet).range(EXCEL_INTERFACE.run_id_cell).value
            except:
                run_id = 'unknown'
            
            if reg_email[-1] == '.':
                plt.xkcd()
                        
            # Submit the request
            req_res = requests.post(url = 'http://guetta.org/addin/validate.php',
                                    headers={'User-Agent': 'XY'},
                                    data = {'run_id':run_id, 'platform':os.name, 'version':v_number, 'email':reg_email,
                                            'settings_string':self._raw_settings_string, 'xlwings_conf':''},
                                    timeout = 10)

            req_res = req_res.json()

            if 'custom_message' in req_res:
                self._v_message += req_res['custom_message']
            
            if 'latest_version' in req_res:
                latest_version = req_res['latest_version']
                
                latest_version = latest_version.replace(',','.')
                v_number = v_number.replace(',','.')
                
                try:
                    float_latest_version = float(latest_version)
                    float_v_number = float(v_number)
                except:
                    # If error, give the warning
                    float_latest_version = 1
                    float_v_number = 0
            
                if float_v_number < float_latest_version:
                    self._v_message += 'You are not using the latest version of XLKitLearn. Please download the '
                    self._v_message += f'latest version at guetta.org/xlkitlearn_latest. The latest version is {latest_version}, '
                    
                    try:
                        if v_number.split('.')[0] == latest_version.split('.')[0]:
                            self._v_message += 'and it will only require-you to re-download this XLKitLearn.xlsm file - '
                            self._v_message += 'no need to re-do the first lengthy installation step.'
                        else:
                            self._v_message += 'and the version of Python XLKitLearn uses has changed, so you\'ll need to '
                            self._v_message += 're-do the installation process from the start. Promise the new features '
                            self._v_message += 'will be worth it!'
                    except:
                        pass
                    
        except:
            pass
        
    def update_status(self, message):
        message = message + ' (elapsed time: ' + str(round(time.time() - self._start_time, 2)) + ' seconds)'
        self._model_sheet.range(self._status_cell).value = message
        
        print(message)
        
        if self._udf_server and time.time() - self._start_time > 80:
            self._out_err.add_error('The add-in has been running for 80 seconds. Unfortunately, Excel is only able '
                                        'to keep an active Python connection for a limited amount of time. When running '
                                        'long runs of the add-in, please uncheck the "keep an active Python connection" '
                                        'option on the settings sheet.', critical=True)
        
    def _read_settings_string(self):
        # Begin by parsing the dictionary. The settings dictionary will
        # be in the form
        #     {'key1'|'val1'`'key2'|'val2'}
        # Where ` is denoted by settings_dict_comma
        #       | is settings_dict_colon
        settings_string = self._model_sheet.range(self._settings_cell).value
        
        # Save the raw setting string
        self._raw_settings_string = settings_string
        
        self.log_run()
        
        try:
            # Ensure we have enclosing braces and remove them
            assert settings_string[0] == '{'
            assert settings_string[-1] == '}'
            settings_string = trim_edges(settings_string)
        
            # Split the settings string into constituent settings
            settings_string = settings_string.split(EXCEL_INTERFACE.settings_dict_comma)
        
            # Convert the settings string into a dictionary
            settings_string = D({trim_edges(i.split(EXCEL_INTERFACE.settings_dict_colon)[0]).strip() :
                                    trim_edges(i.split(EXCEL_INTERFACE.settings_dict_colon)[1]).strip()
                                                                                for i in settings_string})
            
            # Ensure it contains what we expect
            assert set(settings_string) == set(self._expected_settings)
            
        except:
            # If we couldn't read the settings, fail with a critical error
            self._out_err.add_error(f'Error parsing the settings string. Clearing cell {self._settings_cell} '
                                                            'and re-trying might help.', critical=True)
        # Save the settings string
        self._settings = settings_string
        
        # Clean
        for setting_name, setting_spec in self._expected_settings.zip_entries():
            self._settings[setting_name] = self._parse_setting(self._settings[setting_name], setting_spec)
            
    def _parse_setting(self, x, spec):
        if ('kind' not in spec) or (spec.kind is None):
            validations = []
            translator  = lambda x : x
        elif spec.kind == 'i':
            english_key = 'whole number'
            validations  = [lambda x : int(float(x)) == float(x)]
            translator   = lambda x : int(float(x))
        elif spec.kind == 'f':
            english_key = 'number'
            validations  = []
            translator   = lambda x : float(x)
        elif spec.kind == 'b':
            english_key = 'boolean (True or False value)'
            validations  = [lambda x : x in ['True', 'False']]
            translator   = lambda x : {'True':True, 'False':False}[x]
        elif spec.kind == 's':
            english_key = 'string'
            validations = []
           
            def translator(x):
                if 'sklearn_name' in spec:
                  if spec.sklearn_name == 'weights':
                    if x == 'u' or x == 'uniform':
                      return 'uniform'
                    elif x == 'd' or x == 'distance':
                      return 'distance'
                return x
           
        elif spec.kind == 'r':
            english_key = 'Range in an Excel spreadsheet'
            validations  = []
            
            def translator(x):
                if x[:5] == 'File:':
                    return D(file = x[6:].strip())
                else:
                    sheet, cell = x.split('!')
                    if sheet[0] == "'": sheet = trim_edges(sheet)
                    if ']' in sheet: sheet = sheet.split(']')[1]
                    return D(sheet=sheet, cell=cell)            
            
        x = x.strip()
        if x == '':
            if 'default' in spec:
                return spec.default
            else:
                self._out_err.add_error(f'Error processing parameter {spec.english_key}. '
                                            'This parameter cannot be blank, but you provided a blank value.')
        else:
            try:
                for validator in validations:
                    assert validator(x)
                
                return translator(x)
            except:
                self._out_err.add_error(f'Error processing parameter {spec.english_key}. This parameter '
                                     f'needs to be a {english_key}, but you provided {x}.')
    
    @property
    def output_code(self):
        return self._settings.output_code

    @property
    def seed(self):
        return self._settings.seed
    
    @property
    def settings_string(self):
        return self._raw_settings_string

class TextAddinInstance(AddinInstance):
    
    def __init__(self, excel_connector, out_err, udf_server):
        AddinInstance.__init__(self, excel_connector, out_err, TEXT_CONFIG, udf_server)
    
    def load_settings(self):
        self._read_settings_string()
        
        self._out_err.finalize()
        
        for param in ['min_df', 'max_df']:
            if self._settings[param] is not None:
                if (self._settings[param] < 0) or (self._settings[param] > 1):
                    self._out_err.add_error(f'Error parsing the {TEXT_CONFIG.expected_settings[param].english_key} '
                                                'parameter. The parameter, if you provide it, needs to be greater '
                                                'than 0 and less than 1.')
        
        if (self._settings.min_df is not None) and (self._settings.max_df is not None):
            if self._settings.max_df <= self._settings.min_df:
                self._out_err.add_error('Error parsing the upper and lower limiting frequencies. The upper level '
                                            'needs to be higher than the lower leve.')
        
        if self._settings.eval_perc is not None:
            if (self._settings.eval_perc < 0) or (self._settings.eval_perc >= 100):
                self._out_err.add_error('Error parsing the evaluation percentage. The parameter needs to be greater '
                                            'than 0 and less than 100.')
            else:
                self._settings.eval_perc /= 100
                
        if (self._settings.seed is None) and (self._settings.eval_perc is not None):
            self._out_err.add_error( 'You specified you would later be using the data with a train/test split, but '
                                        'you did not specify a seed. Without a seed, there\'s no way to ensure the '
                                        'split here will be the same as the split you use later.')
        
        self._run_lda = False
        if self._settings.lda_topics is not None:
            if self._settings.lda_topics < 2:
                self._out_err.add_error('The number of LDA topics must be two or more.')
                
            self._run_lda = True
        
        self._out_err.finalize()
    
    @property
    def run_lda(self):
        return self._run_lda
        
    @property
    def source_data(self):
        return self._settings.source_data
        
    @property
    def max_df(self):
        return self._settings.max_df
    
    @property
    def min_df(self):
        return self._settings.min_df
    
    @property
    def max_features(self):
        return self._settings.max_features
    
    @property
    def stop_words(self):
        return self._settings.stop_words
    
    @property
    def tf_idf(self):
        return self._settings.tf_idf
    
    @property
    def lda_topics(self):
        return self._settings.lda_topics
    
    @property
    def seed(self):
        return self._settings.seed
    
    @property
    def eval_perc(self):
        return self._settings.eval_perc
        
    @property
    def bigrams(self):
        return self._settings.bigrams
    
    @property
    def stem(self):
        return self._settings.stem
    
    @property
    def output_code(self):
        return self._settings.output_code
    
    @property
    def sparse_output(self):
        return self._settings.sparse_output
    
    @property
    def max_lda_iter(self):
        return self._settings.max_lda_iter

class PredictiveAddinInstance(AddinInstance):
    
    def __init__(self, excel_connector, out_err, udf_server):
        AddinInstance.__init__(self, excel_connector, out_err, PREDICTIVE_CONFIG, udf_server)
        
    def load_settings(self):
        self._read_settings_string()
        
        self._out_err.finalize()
        
        # Model type
        try:
            self._model_name = MODELS.key(self._settings.model)
        except AttributeError:
            self._out_err.add_error(f'You are attempting to fit a model that is not supported by the addin. '
                                          f'The model you provided was {self._settings.model}', critical=True)
        
        # Evaluation percentage
        if self._settings.evaluation_perc is not None:
            if (self._settings.evaluation_perc <= 0) or (self._settings.evaluation_perc >= 100):
                self._out_err.add_error('Error parsing the evaluation percentage. This number needs to be '
                                  'greater than 0 and less than 100, but you provided '
                                     f'{self._settings.evaluation_perc}')
            else:
                self._evaluation_perc = self._settings.evaluation_perc*1.0/100
        else:
            self._evaluation_perc = None
        
        # Split formula parameter
        self._formula = [i.strip() for i in self._settings.formula.split(EXCEL_INTERFACE.list_splitter)]
        
        # Replace "and" with &
        self._formula = [i.replace(' and ', ' & ') for i in self._formula]
        
        # Handle best subset
        if (self._model_name == LINEAR_REGRESSION) and (self._settings.param1.lower() == 'bs'):
            self._best_subset = True
            self._settings.param1 = ''
            
            if len(self.formula) > 1:
                self._out_err.add_error('Error setting up best-subset selection. You entered '
                                          'multiple formulas; I can only do best-subset selection '
                                            'with one formula.')
            else:
                y_var = self.formula[0].split('~')[0].strip()
                x_vars = remove_dupes([x.strip() for x in self.formula[0]
                                                 .split('~')[1].strip().split('+')])
                
                if len(x_vars) > 10:
                    self._out_err.add_error('You are trying to do best-subset selection with more '
                                                'than 10 variables. Totally get why you\'re trying to do '
                                                'this, but this would result in over 1000 competing models, '
                                                'so you might want to consider something a little more robust '
                                                'than an Excel-based addin! Even better, consider using the '
                                                'Lasso.')
                elif '-1' in x_vars:
                    self._out_err.add_error('You are trying to do best-subset selection with an intercept-'
                                                'suppressing term (-1). The right thing to do here is slightly '
                                                'ambiguous, so XLKitLearn does\'t support this.')
                else:
                    self._n_x_terms = len(x_vars)
                    formulas = []
                    for n in range(1, len(x_vars)+1):
                        formulas.extend([y_var + ' ~ ' + ' + '.join(these_x)
                                             for these_x in itertools.combinations(x_vars, n)])
                    self._formula = formulas
                        
        else:
            self._best_subset = False
        
        # Clean parameters
        self._params = D()
        for param_id, param_spec in MODELS[self._model_name].params.zip_entries():
            this_param = D(english_key=param_spec.english_key)
            if ('list_default' in param_spec) and (self._settings[param_id] == ''):
                this_param['vals'] = [param_spec.list_default]
            else:
                this_param['vals'] = [self._parse_setting(i, param_spec)
                                        for i in self._settings[param_id].split(EXCEL_INTERFACE.list_splitter)]
            
            self._params[param_spec.sklearn_name] = this_param
            
        # Create the parameter grid, and determine if tuning is
        # needed
        param_grid = itertools.product(*([self.formula] + [self._params[i].vals for i in self._params]))
        grid_names = ['formula'] + list(self.params)
        self._param_grid = [D({j:k for j,k in zip(grid_names,i)}) for i in param_grid]
        
        self._optimal_params = self._param_grid[0]
        self._needs_tuning = False
        if (self._model_name == BOOSTED_DT) or (len(self._param_grid) > 1):
            self._needs_tuning = True
            self._optimal_params = None
            
        # Ensure K is provided
        if (len(self.tuning_grid) > 0) and ((self.K is None) or (self.K < 2)):
            self._out_err.add_error('Error with the K in k-fold cross-validation. You must provide a K, and it '
                                f'has to be greater or equal to 2. You provided {self.K}.')

    @property
    def model_name(self):
        return self._model_name
    
    @property
    def english_model_name(self):
        return MODELS[self._model_name].english_key
        
    @property
    def formula(self):
        return self._formula
        
    @property
    def params(self):
        return self._params
        
    @property
    def training_data(self):
        return self._settings.training_data
    
    @property
    def K(self):
        return self._settings.K
        
    @property
    def evaluation_perc(self):
        return self._evaluation_perc
    
    @property
    def evaluation_data(self):
        return self._settings.evaluation_data
    
    @property
    def prediction_data(self):
        return self._settings.prediction_data
    
    @property
    def output_model(self):
        return self._settings.output_model
    
    @property
    def output_evaluation_details(self):
        return self._settings.output_evaluation_details
    
    @property
    def best_subset(self):
        return self._best_subset
    
    @property
    def tuning_grid(self):
        if self.needs_tuning:
            return self._param_grid
        else:
            return []
    
    @property
    def needs_tuning(self):
        return self._needs_tuning
        
    @property
    def optimal_params(self):
        return self._optimal_params
    
    @optimal_params.setter
    def optimal_params(self, value):
        self._optimal_params = value
    
# =======================
# =   Dataset Classes   =
# =======================
      
class SparseDataset():
    
    def init_from_pandas(self, data, base_dataset=None):
        # Ensure rows are sorted consistently. This is in case we use output
        # from the text add-in here
        unique_rows = data.ROW.unique()
        self._row_to_key = ( list(sorted([i for i in unique_rows if type(i) in [int, float]]))
                               + list(sorted([i for i in unique_rows if type(i) not in [int, float]])) )
                                
        # Create maps from keys to rows and rows to keys
        self._key_to_row = {j:i for i, j in enumerate(self._row_to_key)}
        
        if base_dataset is None:
            self._col_to_key = list(data.COLUMN.unique())
        else:
            # Inherit column data
            self._col_to_key = list(base_dataset.columns)
            
            # Remove any columns that don't occur in the base dataset
            data = data[data.COLUMN.isin(self._col_to_key)]
        
        self._key_to_col = {j:i for i, j in enumerate(self._col_to_key)}
        
        # Replace row and columns keys with numbers
        data = (data.rename(columns={'ROW':'old_ROW', 'COLUMN':'old_COLUMN'})
                    .merge(pd.DataFrame({'old_ROW':list(self._key_to_row.keys()), 'ROW':list(self._key_to_row.values())}),
                           how='left',
                           validate='many_to_one')
                    .merge(pd.DataFrame({'old_COLUMN':list(self._key_to_col.keys()), 'COLUMN':list(self._key_to_col.values())}),
                           how='left',
                           validate='many_to_one'))[['ROW', 'COLUMN', 'VALUE']]
        #data.ROW = data.ROW.replace(self._key_to_row).astype(int)
        #data.COLUMN = data.COLUMN.replace(self._key_to_col).astype(int)
        
        # Create the sparse matrix
        self._data = sparse.csr_matrix((data.VALUE, (data.ROW, data.COLUMN)))
        
    def init_from_sparse(self, data, col_to_key, row_to_key):
        self._row_to_key = row_to_key
        self._col_to_key = col_to_key
        
        self._key_to_row = {j:i for i, j in enumerate(self._row_to_key)}
        self._key_to_col = {j:i for i, j in enumerate(self._col_to_key)}
        
        self._data = data
    
    def __len__(self):
        return self._data.shape[0]
        
    def __getitem__(self, key):
    
        if type(key) == list:
            cols = [self._key_to_col[i] for i in key]
            return self._data[:, cols]
        elif type(key) == str:
            cols = self._key_to_col[key]
            return self._data[:, cols].todense().transpose()
        else:
            raise AssertionError('Attempting to access item from SparseDataset in a non-supported way.')
        
        
    
    @property
    def columns(self):
        return self._col_to_key

    @property
    def index(self):
        return self._row_to_key
    
    @property
    def matrix(self):
        return self._data
    
class Datasets:
    '''
    This class stores all datasets for a given run of the addin.
    '''

    def __init__(self, addin, out_err, excel_connector, ensure_numeric):
        
        # Keep track of the error and model object
        self._out_err = out_err
        self._addin = addin
        self._excel_connector = excel_connector
        
        self._ensure_numeric = ensure_numeric
        
        # Load the datasets
        # -----------------
        self._raw_data = D()

        self._is_sparse = None
        for dataset_name in ['training_data', 'evaluation_data', 'prediction_data']:
            dataset = self._load_dataset(data_range=getattr(self._addin, dataset_name),
                                            dataset_name=dataset_name,
                                            base_dataset=self._raw_data.training_data if 'training_data' in self._raw_data else None)
            self._raw_data[dataset_name] = dataset

            if self._is_sparse is None:
                self._is_sparse = (type(dataset) is SparseDataset)
            elif (dataset is not None) and (self._is_sparse != (type(dataset) is SparseDataset)):
                self._out_err.add_error('It seems you are mixing sparse and dense datasets. Please make sure every ' 
                                        'dataset is of the same type.')

        # Ensure we have a training dataset
        if (self._raw_data.training_data is None) or len(self._raw_data.training_data) == 0:
            self._out_err.add_error('You seem to have selected an empty training set. Please select '
                                        'a training dataset and re-run.', critical=True)

        # If we need to split the training set, do it
        do_split = lambda x : sk_ms.train_test_split(x,
                                           train_size = 1-self._addin.evaluation_perc,
                                           test_size = self._addin.evaluation_perc,
                                           random_state = self._addin.seed)
                                           
        if (self._raw_data.evaluation_data is None) and (self._addin.evaluation_perc is not None):
            if self._is_sparse:
                training_mat, evaluation_mat = do_split(self._raw_data.training_data.matrix)
                training_rows, evaluation_rows = do_split(self._raw_data.training_data.index)
                columns = self._raw_data.training_data.columns
                
                self._raw_data.training_data = SparseDataset()
                self._raw_data.training_data.init_from_sparse(training_mat, columns, training_rows)
                
                self._raw_data.evaluation_data = SparseDataset()
                self._raw_data.evaluation_data.init_from_sparse(evaluation_mat, columns, evaluation_rows)
                
            else:
                self._raw_data.training_data, self._raw_data.evaluation_data = (
                                            do_split(self._raw_data.training_data) )

        self._binary_data = None
        self._current_formula = None
        self._outcome = None
        self._data = D()
        self._stored_formulas = []

    def _load_dataset(self, data_range, dataset_name, base_dataset = None):
        if data_range is None: return
        
        if 'cell' in data_range:
            try:
                raw_data = self._excel_connector.wb.sheets(data_range.sheet).range(data_range.cell).options(ndim=2).value
            except:
                self._out_err.add_error(f'Error reading the {dataset_name} dataset in sheet {data_range.sheet}, cells '
                                            f'{data_range.cell}. This usually happens because the range you selected '
                                            "doesn't exist in your spreadsheet. Could you double-check?")
                return
        else:
            try:
                file_path = self._excel_connector.wb.sheets('code_text').range(EXCEL_INTERFACE.path_cell).value
                delim = '/' if '/' in file_path else '\\'
                file_path = file_path + delim
                
                if data_range.file.split('.')[1] in ['xls', 'xlsx']:
                    # If we use header=None and there are headers, the whole column will
                    # be read as a string
                    raw_data = pd.read_excel(file_path+data_range.file, header=0, na_values=['', ' '], keep_default_na=False)
                    raw_data = [list(raw_data.columns)] + raw_data.values.tolist()
                elif data_range.file.split('.')[1] == 'csv':
                    raw_data = pd.read_csv(file_path+data_range.file, header=0, na_values=['', ' '], keep_default_na=False)
                    raw_data = [list(raw_data.columns)] + raw_data.values.tolist()
                else:
                    self._out_err.add_error(f'Error reading the {dataset_name} dataset in file {data_range.file}. XLKitLearn '
                                                'is only able to accept Excel and CSV files. Please ensure you file has the '
                                                'extension xls, xlsx, or csv.')
            except FileNotFoundError:
                error_message = (f'The filename you provided for the {dataset_name} does not exist, or could not be found '
                                    'in the same directory as this Excel spreadsheet. Here is some info that could help '
                                    'with debugging:\n'
                                    '  - The file name you entered was\n'
                                   f'        {data_range.file}\n'
                                    '  - I\'ve detected the path this spreadsheet sits in as\n'
                                   f'        {file_path}')
                                   
                try:
                    close_files = [i for i in os.listdir(file_path)
                                    if levenshtein_ratio_and_distance(i.split('.')[0], data_range.file.split('.')[0]) <= 2]
                    if len(close_files) > 1:
                        error_message += '\n\nThe following files are in the same directory and have similar names. Perhaps you meant '
                        error_message += f'to use one of those: {", ".join(close_files)}'
                except:
                    pass
                
                self._out_err.add_error(error_message, critical=True)
                                            
        # Check whether we have a sparse dataset
        if raw_data[0] == ['ROW', 'COLUMN', 'VALUE']:
            # We have a sparse dataset
            out = SparseDataset()
            
            raw_df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
            raw_df.ROW = raw_df.ROW.astype(str)
            raw_df.COLUMN = raw_df.COLUMN.astype(str)
            
            try:
                out.init_from_pandas(raw_df, base_dataset)
            except TypeError:
                self._out_err.add_error('It looks like you provided a sparse matrix in which some entries are '
                                            'not numeric. XLKitLearn only supports numeric entries in a sparse '
                                           f'matrix. The error occured in the following dataset: {dataset_name}.', critical=True)
                                           
        else:
            # We have a dense dataset
            
            # If every column in this new dataset existed in the base dataset,
            # take the header columns in this new dataset as the headers
            if base_dataset is None:
                headers = raw_data[0]
                raw_data = raw_data[1:]
                
                def robust_strip(txt):
                    try:
                        return txt.strip()
                    except:
                        return ' '
                
                blank_headers = []
                for this_header_i, this_header in enumerate(headers):
                    if robust_strip(this_header) == '':
                        blank_headers.append((this_header_i+1, this_header))
                
                if len(blank_headers) > 0:
                    self._out_err.add_error(f'Some of the column names in your training data are blank. Please check your data. The problematic headers and their positions were {blank_headers}.', critical=True)
            else:
                if len(set(raw_data[0]) - set(base_dataset.columns)) == 0:
                    headers = raw_data[0]
                    raw_data = raw_data[1:]
                elif len(raw_data[0]) == len(base_dataset.columns):
                    # The first row in our dataset doesn't seem to match the original
                    # dataset, but it has the same number of columns. Use the original
                    # columns as the columns in this dataset
                    headers = list(base_dataset.columns)
                else:
                    self._out_err.add_error(f'Error reading the {dataset_name} dataset in sheet {data_range.sheet}, cells '
                                                  f'{data_range.cell}. Evaluation and prediction datasets must either contain '
                                                  'headers that also exist in the original training dataset, OR can have '
                                                  'no headers but exactly the same number of columns as the training '
                                                  'dataset.')
                    return
                
            out = pd.DataFrame(raw_data, columns=headers)
            
            if self._ensure_numeric != False:
                try:
                    out[[i for i in out.columns if i != self._ensure_numeric]].astype(float)
                except:
                    self._out_err.add_error('You are trying to fit a y~. formula (one that includes every '
                                              'column), but you provided a dataset in which some of the columns '
                                              'are not numeric. Unfortunately, XLKitLearn is not able to automatically '
                                              'create categorical factors when the columns are not specified directly '
                                              'in the formula.')
                                              
            # Shift the dataset so the index starts from 1 to match Excel rows
            out.index += 1
        
        # Ensure the dataset is not empty
        if len(out) == 0:
            self._out_err.add_error(f'Error reading the {dataset_name} dataset in sheet {data_range.sheet}, cells '
                                            f'{data_range.cell}. The dataset was empty.')
            return

        return out

    def _set_binary(self, y_values = None):
        
        y_binary = (set(y_values) == {0, 1})

        if self._binary_data is None:
            self._binary_data = y_binary
        else:
            if self._binary_data != y_binary:
                self._out_err.add_error('It seems you are mixing binary and continuous prediction problems. Please '
                                        'ensure the outcomes in all your datasets and all your formulas are either '
                                        'all continuous or all binary.', critical=True)
    
    def _set_outcome(self, formula):
        
        outcome = formula.split('~')[0].strip()
        
        if self._outcome is None:
            self._outcome = outcome
        elif (self._outcome != outcome):
            self._out_err.add_error('It seems you are trying to fit two models with different outcome variables with '
                                    "a sparse dataset. This doesn't make much sense; the outcome variable in one "
                                    'instance will be included in the independent variables of the other model. Please '
                                    'review your formulas.', critical=True)
                                    
    @property
    def binary(self):
        return self._binary_data

    @property
    def is_sparse(self):
        return self._is_sparse

    def _check_headers(self, headers):

        invalid_headers = [header for header in headers
                           if keyword.iskeyword(header) or (not header.isidentifier())
                           or (header == 'Intercept')]
        duped_headers = pd.Series(headers).pipe(lambda x: x[x.duplicated()]).unique().tolist()

        if len(invalid_headers) > 0:
            self._out_err.add_error(f'Your dataset contains some columns with invalid names. Column names '
                                    f'must be valid Python variable names; they cannot, for example, start '
                                    f'with numbers, or be reserved Python keywords. The problematic columns '
                                    f'were {", ".join(invalid_headers)}.')
            return False

        if len(duped_headers) > 0:
            self._out_err.add_error(f'Your dataset contains some duplicated columns. The problematic columns '
                                    f'were {", ".join(duped_headers)}.')
            return False

        return True

    def _log_patsy_error(self, e, dataset_name):
        try:
            error_parts = str(e).split('\n')
            problem_factor = ''.join([error_parts[1][i] for i, j in enumerate(error_parts[2]) if j == '^'])

            if 'missing values' in error_parts[0]:
                self._out_err.add_error(f'One of variables in your formula resulted in a missing value in the '
                                        f'{dataset_name}. The offending term was {problem_factor}. Please double '
                                        f'check this column has a value in every rows.')
            elif 'NameError' in error_parts[0]:
                if dataset_name == 'training set':
                    self._out_err.add_error(f'One of the terms in your formula does not exist in the '
                                            f'{dataset_name}. The offending term was {problem_factor}. This '
                                            f'usually happens because of a typo in a variable name; remember '
                                            f'that all variables names are case sensitive.')
                else:
                    self._out_err.add_error(f'One of the terms in your formula does not exist in the '
                                            f'{dataset_name}. The offending term was {problem_factor}. This '
                                            f'usually happens because there is a column in the training set '
                                            f'that is missing from the {dataset_name}.')
            elif 'match any of the expected levels':
                new_value = str(e).split('observation with value ')[1].split(' ')[0]
                self._out_err.add_error(f'The categorical variable {problem_factor} contains a value in the '
                                        f'{dataset_name} that does not exist in the training set, so the model '
                                        f"can't interpret it. The offending value was {new_value}. This could "
                                         'be because the value is rare, and the train/test split put every row '
                                         'with that value in the evaluation set. You might want to combine less '
                                         'frequent values of this categorical variable, or simply remove rows '
                                         'with this rare value from your dataset.')
            else:
                self._out_err.add_error(f'There was an unknown error in evaluating your formula on the '
                                        f'{dataset_name}. The specific error was\n\n{e}')
        except:
            self._out_err.add_error(f'There was an unknown error in evaluating your formula on the '
                                        f'{dataset_name}. The specific error was\n\n{e}')
    
    def _fix_design_matrix(self, design_matrix):
        '''
        This function will take a Patsy design_matrix object and return a dictionary with an
        X and y value. If the object has two dimensions, the first is assumed to be y, and
        the second to be X. If it has one dimension, we assume we only have an X.

        It will then attempt to remove any intercept term from the matrix, because we'll be
        using sklearn's own intercept functionality
        '''
        
        if len(design_matrix) == 2:
            out = D()
            out['X'] = np.asarray(design_matrix[1])
        
            if design_matrix[0].shape[1] == 1:
                # We have a single y-value
                out['y'] = np.asarray(design_matrix[0]).transpose()[0]
            elif ((design_matrix[0].shape[1] == 2)
                        and (design_matrix[0].design_info.column_names[0][-7:] == '[False]')
                        and (design_matrix[0].design_info.column_names[1][-6:] == '[True]')):
                # We have a True/False y-value
                out['y'] = np.asarray(design_matrix[0])[:,1]
            else:
                self._out_err.add_error('It looks like your outcome column (the bit to the left of the '
                                        "~) contains some text values. Unfortunately, XLKitLearn doesn't "
                                        'quite know what to do with that; how is it supposed to predict a '
                                        'text value? This sometimes happens because some of the rows contain '
                                        'text like N/A, or NA, to denote missing values. Make sure you remove '
                                        'any such rows when you clean your data before using '
                                        'XLKitLearn.', critical=True)
        
        else:
            out = D(X=np.asarray(design_matrix[0]))

        try:
            intercept_index = design_matrix[-1].design_info.column_names.index('Intercept')
            out.X = np.delete(out.X, intercept_index, axis=1)
        except ValueError:
            pass

        return out

    @staticmethod
    def _fix_col_name(col_name):
        out = []
        for formula_part in col_name.split(':'):
            if ('[T.' in formula_part) and (formula_part[-1] == ']'):
                try:
                    feature_name = formula_part.split('[T.')[0]
                    feature_val = formula_part.split('[T.')[1][:-1]
                    
                    if feature_name[:2] == 'C(' and feature_name[-1] == ')':
                        feature_name = feature_name[2:].split(',')[0].strip()
                        
                        if feature_name[-1] == ')':
                            feature_name = feature_name[:-1]
                    
                    out.append(f'{feature_name} = {feature_val}')
                except:
                    out.append(formula_part)
            else:
                out.append(formula_part)

        return ':'.join(out)

    def set_formula(self, formula):
        if self.is_sparse:
            self._set_outcome(formula)
        
        self._current_formula = formula

        if formula not in self._data:
            formula_no_space = ''.join([i for i in formula if i != ' '])

            # Figure out the outcome column
            out_col = formula_no_space.split('~')[0].strip()
            
            # Ensure the outcome column is the same as previous ones, if this is not the first
            # formula
            if (len(self._data) > 0):
                existing_out_col = ''.join([i for i in list(self._data.keys())[0].split('~')[0] if i != ' ']).strip()
                if (out_col != existing_out_col):
                    self._out_err.add_error('When entering multiple formulas, every formula must use the SAME outcome variable.', critical=True)

            # Add an entry to the data dictionary
            is_first_formula = (len(self._data) == 0)
            self._stored_formulas.append(formula)
            self._data[formula] = D()

            # Check whether the outcome column has brackets and an equal sign.
            if ('=' in out_col):
                if out_col[0] == '(' and out_col[-1] == ')':
                    out_col = out_col[1:-1]
                
                target_val = out_col.split('=')[1].strip()
                out_col = out_col.split('=')[0].strip()
                
                # If this is our first formula, modify the outcome column
                if is_first_formula:
                    def is_target_val(i):
                        try:
                            if str(i) == target_val: return True
                            if i == int(i) and str(int(i)) == target_val: return True
                        except:
                            return False
                            
                    # Modify the data to reflect the new target column; because we insist
                    # all formulas have the same outcome, this won't cause trouble anywhere
                    # else
                    self._raw_data.training_data[out_col] = [1 if is_target_val(i) else 0 for i in self._raw_data.training_data[out_col]]
                    
                    # Ensure we don't have all zeroes
                    if self._raw_data.training_data[out_col].sum() == 0:
                        self._out_err.add_error('You are asking xlkitlearn to fit a binary prediction model to predict whether column '
                                                    f'{out_col} is equal to {target_val}, but not a single one of the values in that '
                                                     'column is equal to that! Please check...', critical=True)
                    
                    # Modify evaluation data if necessary (prediction data won't have an outcome column)
                    if self._raw_data.evaluation_data is not None:
                        if out_col not in  self._raw_data.evaluation_data:
                            self._out_err.add_error(f'Your manual evaluation set does not contain a column called {out_col}, which you use in your formula.', critical=True)
                        
                        self._raw_data.evaluation_data[out_col] = [1 if is_target_val(i) else 0 for i in self._raw_data.evaluation_data[out_col]]
                    
                    
                    
                # Create an internal version of this formula for processing purposes that
                # includes the column name itself [i.e., y instead of (y=1)]. Note that
                # out_col is already updated above
                formula_no_space = out_col + '~' + formula_no_space.split('~')[1]
                
            if (formula_no_space[-1] == '.') or (formula_no_space[-3:] == '.-1'):
                other_cols = [i for i in self._raw_data.training_data.columns if i != out_col]

                # Get datasets
                for dataset_name in ['training_data', 'evaluation_data']:
                    if self._raw_data[dataset_name] is not None:
                        this_X = self._raw_data[dataset_name][other_cols]
                        this_X = np.asarray(this_X) if type(this_X) is pd.DataFrame else this_X
                        
                        try:
                            self._data[formula][dataset_name] = D(X=this_X,
                                                                  y=np.asarray(self._raw_data[dataset_name][out_col]).ravel())
                        except KeyError:
                            self._out_err.add_error(f'It looks like {out_col} is not one of the columns in your '
                                                        'dataset, please double check.', critical=True)
                        
                        self._set_binary(self._data[formula][dataset_name].y)
                        
                if self._raw_data.prediction_data is not None:
                    self._data[formula]["prediction_data"] = D(X=self._raw_data.prediction_data[other_cols])
                    
                #
                self._data[formula]['columns'] = other_cols

                #
                self._data[formula]['intercept'] = (formula_no_space[-2:] != '-1')

            else:
                # Ensure we don't have a sparse dataset
                if self.is_sparse:
                    self._out_err.add_error('You are attempting to use a formula on a sparse dataset. This add-in only '
                                            'supports sparse datasets with formulas of the form y~. which use every '
                                            'column in the data.', critical=True)

                # Check for invalid columns. Only check the training set; those are the only
                # columns that will be used
                if not self._check_headers(self._raw_data.training_data.columns):
                    self._out_err.finalize()
                
                # Process the training set using patsy
                try:
                    patsy_train = pt.dmatrices(formula_no_space, self._raw_data.training_data, NA_action='raise')
                    self._data[formula]['training_data'] = self._fix_design_matrix(patsy_train)
                    self._set_binary(self._data[formula].training_data.y)
                except pt.PatsyError as e:
                    self._log_patsy_error(e, 'training set')
                    self._out_err.finalize()

                # Process the evaluation set using patsy
                if self._raw_data.evaluation_data is not None:
                    try:
                        self._data[formula]['evaluation_data'] = self._fix_design_matrix(pt.build_design_matrices(
                                                                                    [patsy_train[i].design_info for i in [0,1]],
                                                                                    self._raw_data.evaluation_data,
                                                                                    NA_action = 'raise'))
                        self._set_binary(self._data[formula].evaluation_data.y)
                    except pt.PatsyError as e:
                        self._log_patsy_error(e, 'evaluation set')

                # Process the prediction data using patsy
                if self._raw_data.prediction_data is not None:
                    try:
                        self._data[formula]['prediction_data'] = self._fix_design_matrix(pt.build_design_matrices([patsy_train[1].design_info],
                                                                                                             self._raw_data.prediction_data,
                                                                                                             NA_action = 'raise'))
                    except pt.PatsyError as e:
                        self._log_patsy_error(e, 'prediction data')

                # Save the column names
                self._data[formula]['columns'] = [self._fix_col_name(i) for i in patsy_train[1].design_info.column_names
                                                                                                   if i != 'Intercept']

                # Check whether we need an intercept
                self._data[formula]['intercept'] = ('Intercept' in patsy_train[1].design_info.column_names)

        if len(self._stored_formulas) > 5:
            del self._data[self._stored_formulas.pop(0)]
        
        self._out_err.finalize()
            
    @property
    def intercept(self):
        return self._data[self._current_formula].intercept
    
    @property
    def X_train(self):
        return self._data[self._current_formula].training_data.X
        
    @property
    def X_eval(self):
        if self._raw_data.evaluation_data is not None:
            return self._data[self._current_formula].evaluation_data.X
        return None
        
    @property
    def X_pred(self):
        if self._raw_data.prediction_data is not None:
            return self._data[self._current_formula].prediction_data.X
        return None
        
    @property
    def y_train(self):
        return self._data[self._current_formula].training_data.y
        
    @property
    def y_eval(self):
        if self._raw_data.evaluation_data is not None:
            return self._data[self._current_formula].evaluation_data.y
        return None
        
    @property
    def y_pred(self):
        if self._raw_data.prediction_data is not None:
            return self._data[self._current_formula].prediction_data.y
        return None
        
    @property
    def eval_rows(self):
        return self._raw_data.evaluation_data.index
        
    @property
    def pred_rows(self):
        return self._raw_data.prediction_data.index
        
    @property
    def columns(self):
        return self._data[self._current_formula].columns

# =====================
# =   Model Classes   =
# ===================== 
    
class AddinModel:

    def __init__(self, model_name, params, binary_data, seed, out_err):
        self._out_err = out_err
        self._binary_data = binary_data
        self._model_name = model_name
        self._params = params
        self._seed = seed

        if binary_data:
            self._eval_func = roc_auc_score
        else:
            self._eval_func = r2_score

        if model_name == LINEAR_REGRESSION:
            if binary_data:
                if params.alpha == 0:
                    self._model = LogisticRegression(penalty="none", max_iter=MAX_LR_ITERS, solver='newton-cg')
                else:
                    self._model = LogisticRegression(C=1.0 / params.alpha, penalty="l1", solver='liblinear', max_iter=MAX_LR_ITERS)
            else:
                if params.alpha == 0:
                    self._model = LinearRegression()
                else:
                    self._model = Lasso(alpha=params.alpha, normalize=True)
                
        elif model_name == NEAREST_NEIGHBORS:
            if binary_data:
                if params.weights == "d" or params.weights == "distance":
                    self._model = KNeighborsClassifier(n_neighbors=params.n_neighbors, weights="distance")
                else:
                    self._model = KNeighborsClassifier(n_neighbors=params.n_neighbors, weights="uniform")
            else:
                if params.weights == "d" or params.weights == "distance":
                    self._model = KNeighborsRegressor(n_neighbors=params.n_neighbors, weights="distance")
                else:
                    self._model = KNeighborsRegressor(n_neighbors=params.n_neighbors, weights="uniform")

        elif model_name == DECISION_TREE:
            if binary_data:
                self._model = DecisionTreeClassifier(max_depth=params.max_depth, random_state=seed)
            else:
                self._model = DecisionTreeRegressor(max_depth=params.max_depth, random_state=seed)

        elif model_name == BOOSTED_DT:
            if binary_data:
                self._model = GradientBoostingClassifier(max_depth=params.max_depth, random_state=seed,
                                                   n_estimators=params.n_estimators,
                                                   learning_rate=params.learning_rate)
            else:
                self._model = GradientBoostingRegressor(max_depth=params.max_depth, random_state=seed,
                                                  n_estimators=params.n_estimators,
                                                  learning_rate=params.learning_rate)
                                                            
        elif model_name == RANDOM_FOREST:
            if binary_data:
                self._model = RandomForestClassifier(max_depth=params.max_depth, random_state=seed,
                                               n_estimators=params.n_estimators)                            
            else:
                self._model = RandomForestRegressor(max_depth=params.max_depth, random_state=seed,
                                              n_estimators=params.n_estimators)       
                                                                
        else:
            raise
    
    @staticmethod
    def _get_model_string(model_name, params, binary_data, seed):
        import_string = ''
        model_string = ''
        
        if model_name == LINEAR_REGRESSION:
            import_string = 'import sklearn.linear_model as sk_lm'
            if binary_data:
                if params.alpha == 0:
                    model_string = 'sk_lm.LogisticRegression(penalty="none", solver="newton-cg")'    
                else:
                    model_string = f'sk_lm.LogisticRegression(C=1/{params.alpha}, penalty="l1", solver="liblinear")'
            else:
                if params.alpha == 0:
                    model_string = 'sk_lm.LinearRegression()'
                else:
                    model_string = f'sk_lm.Lasso(alpha={params.alpha}, normalize=True)'

        elif model_name == NEAREST_NEIGHBORS:
            import_string = 'import sklearn.neighbors as sk_n'
            if binary_data:
                model_string = AddinModel._make_model_string('sk_n.KNeighborsClassifier', params, ['n_neighbors', 'weights'], None)
            else:
                model_string = AddinModel._make_model_string('sk_n.KNeighborsRegressor', params, ['n_neighbors', 'weights'], None)
                    
        elif model_name == DECISION_TREE:
            import_string = 'import sklearn.tree as sk_t'
            if binary_data:
                model_string = AddinModel._make_model_string('sk_t.DecisionTreeClassifier', params, ['max_depth'], seed)
            else:
                model_string = AddinModel._make_model_string('sk_t.DecisionTreeRegressor', params, ['max_depth'], seed)

        elif model_name == BOOSTED_DT:
            import_string = 'import sklearn.ensemble as sk_e'
            if binary_data:
                model_string = AddinModel._make_model_string('sk_e.GradientBoostingClassifier', params,
                                                                ['max_depth', 'n_estimators', 'learning_rate'], seed)
            else:
                model_string = AddinModel._make_model_string('sk_e.GradientBoostingRegressor', params,
                                                                ['max_depth', 'n_estimators', 'learning_rate'], seed)
                                                            
        elif model_name == RANDOM_FOREST:
            import_string = 'import sklearn.ensemble as sk_e'
            if binary_data:
                model_string = AddinModel._make_model_string('sk_e.RandomForestClassifier', params,
                                                                ['max_depth', 'n_estimators'], seed)                               
            else:
                model_string = AddinModel._make_model_string('sk_e.RandomForestRegressor', params,
                                                                ['max_depth', 'n_estimators'], seed)   
                                                                
        else:
            raise    
    
        return D({'import_string':import_string, 'model_string':model_string})
    
    @staticmethod
    def _make_model_string(f_name, params, potential_params, seed):
        existing_params = []
        for p in potential_params:
            if p in params:
              if isinstance(params[p], str):
                existing_params.append(f'{p}=\"{params[p]}\"')
              else: 
                existing_params.append(f'{p}={params[p]}')
        
        if seed is not None: existing_params.append(f'random_state={seed}')
        
        return f_name + '(' + ', '.join(existing_params) + ')'
    
    @property
    def model_string(self):
        return self._model_string
    
    @property
    def import_string(self):
        return self._import_string
    
    def fit(self, X, y, intercept):
        if self._model_name == LINEAR_REGRESSION:
            self._model.set_params(fit_intercept=intercept)
        
        try:
            self._model.fit(X, y)
        except Exception as e:
            self._out_err.add_error('An unidentified error has occurred while fitting your model. This happened '
                                            f'during the sklearn stage. The full error text was\n\n{e}', critical=True)
        
        # If the model in question involves a number of iterations, check that
        # we haven't hit the maximum number
        self._max_iter_reached = False
        try:
            if self._model.n_iter_ == self._model.max_iter:
                self._max_iter_reached = True
        except:
            pass
        
        self._X = X 
        self._y = y
        self._intercept = intercept
        
        # sklearn is really bad at identifying numerical errors. If we have a linear
        # regression with alpha=0, fit it using statsmodels too
        self._fit_lr_using_statsmodels()
        
        return self

    def predict(self, X):
        if self._binary_data:
            return [i[1] for i in self._model.predict_proba(X)]
        else:
            return self._model.predict(X)

    def staged_predict(self, X):
        if self._model_name == BOOSTED_DT:
            if self._binary_data:
                return [[i[1] for i in j] for j in self._model.staged_predict_proba(X)]
            else:
                return list(self._model.staged_predict(X))

        elif self._model_name == RANDOM_FOREST:
            stand_matrix = lambda x : (np.cumsum(x, axis=0) / np.array(range(1,len(x)+1))[:,None]).tolist()
            if self._binary_data:
                return stand_matrix(np.array([[i[1] for i in j.predict_proba(X)] for j in self._model.estimators_]))
            else:
                return stand_matrix([i.predict(X) for i in self._model.estimators_])
        
        else:
            return [self.predict(X)]

    def evaluate(self, y_true, y_pred=None, X=None):
        if X is not None:
            y_pred = self.predict(X)

        return self._eval_func(y_true, y_pred)

    def staged_evaluate(self, y_true, y_staged_pred=None, X=None):
        if X is not None:
            y_staged_pred = self.staged_predict(X)
        
        return [self._eval_func(y_true, i) for i in y_staged_pred]
    
    def n_nonzero_coefs(self):
        if self._model_name == LINEAR_REGRESSION:
            if self._binary_data:
                return sum([i != 0 for i in self._model.coef_[0]])
            else:
                return sum([i != 0 for i in self._model.coef_])
        
        return None
    
    def print_roc_curve(self, y_true, y_pred=None, X=None, out=None):
        if self._binary_data:
            if X is not None:
                y_pred = self.predict(X)
            
            fpr, tpr, _ = roc_curve(y_true, y_pred)
            
            fig, ax = plt.subplots(1, 1, figsize=(9, 6))
            ax.plot(fpr, tpr)
            plt.plot([0, 1], [0, 1], linestyle = "--")
            plt.xlabel("False Positive Rate", fontsize = 18)
            plt.ylabel("True Positive Rate", fontsize = 18)
            sns.despine()
        
            out.add_graph(fig, 2)
    
    def _fit_lr_using_statsmodels(self):
        if (self._model_name == LINEAR_REGRESSION) and (self._params.alpha == 0) and (type(self._X) is not sparse.csr_matrix):
            if self._intercept:
                this_X = np.hstack([np.ones((len(self._X),1)), self._X])
            else:
                this_X = self._X
            
            try:
                if self._binary_data:
                    self._sm_result = sm.GLM(self._y, this_X, family=sm.families.Binomial()).fit(maxiter=MAX_LR_ITERS)
                    #self._sm_result = sm.Logit(self._y, this_X).fit(maxiter=MAX_LR_ITERS)
                else:
                    self._sm_result = sm.OLS(self._y, this_X).fit(maxiter=MAX_LR_ITERS)
            except np.linalg.LinAlgError:
                self._out_err.add_error('Something\'s gone wrong - a linear algebra error has occurred. '
                                            'this usually happens because two columns are perfectly correlated. '
                                            'Consider removing one of the columns or using a penalized model.', critical=True)
            except Exception as e:
                self._out_err.add_error('An unidentified error has occurred while fitting your linear regression. This happened '
                                            f'during the statsmodels stage. The full error text was\n\n{e}', critical=True)
            
            self._good_condition_number = True
            if (not self._binary_data) and (self._sm_result.condition_number > 1E12):
                self._good_condition_number = False
            
            self._sm_converged = True
            if (self._binary_data) and (not self._sm_result.converged): # (not self._sm_result.mle_retvals['converged']):
                self._sm_converged = False                  
    
    def print_reg_results(self, columns, get_p_vals, out):
        if self._model_name == LINEAR_REGRESSION:
            
            coef_table = D({'':[], 'Coefficient':[]})
            
            if self._intercept:
                coef_table[''].append('Intercept')
                coef_table.Coefficient.append(self._model.intercept_[0] if self._binary_data
                                                                        else self._model.intercept_)
            
            for column, coef in zip(columns, self._model.coef_[0] if self._binary_data else self._model.coef_):
                if ((self._params.alpha == 0) and get_p_vals) or coef != 0:
                    coef_table[''].append(column)
                    coef_table.Coefficient.append(coef)
            
            if (self._params.alpha == 0) and get_p_vals:
                if not self._binary_data:
                    coef_table['t stat'] = self._sm_result.tvalues
                
                coef_table['[2.5% CI'] = self._sm_result.conf_int()[:, 0]
                coef_table['97.5% CI]'] = self._sm_result.conf_int()[:, 1]
                
                coef_table['p-value'] = list(self._sm_result.pvalues)
                                
                def pval_to_star(x):
                    if x <= 0.001:
                        return "***"
                    elif x <= 0.01:
                        return "**"
                    elif x <= 0.05:
                        return "*"
                    elif x <= 0.1:
                        return "."
                    else:
                        return ''
                
                coef_table['Significance'] = [pval_to_star(i) for i in self._sm_result.pvalues]
            
            if (self._model_name == LINEAR_REGRESSION) and (self._params.alpha == 0) and (type(self._X) is not sparse.csr_matrix):
                if not self._sm_converged:
                    out.add_header('WARNING', 2, 'The algorithm XLKit learn used to find p-values did not converge. This could be nothing or it could be a sign your results are unreliable. Please seek help.')
                    out.add_blank_row()
                
                if not self._good_condition_number:
                    out.add_header('WARNING', 2, 'The condition number for your linear regression is too large - this typically means two '
                                                    'of your columns are highly correlated. Consider removing one of the columns.')
                    out.add_blank_row()
            
            if self._max_iter_reached:
                out.add_header('WARNING', 2, 'The algorithm that fit your model did not converge. This could be nothing or it could be a sign your results are unreliable. Please seek help.')
                out.add_blank_row()
            
            out.add_table(pd.DataFrame(coef_table))
            out.add_blank_row()
            
            if (self._params.alpha == 0) and get_p_vals:
                out.add_row( ["Significance codes: 0 `***` 0.001 `**` 0.01 `*` 0.05 `.` 0.1 ` ` 1"],
                             ["italics"])
                out.add_blank_row()
    
    def print_tree(self, columns, out):
        if self._model_name == DECISION_TREE:
            tree_printer = TreePrinter(self._model, columns)
            
            if tree_printer.tree_depth <= 5:
                out.add_graph(tree_printer.plot_tree(), 2)
            else:
                out.add_row(['Graphical representations are not printed for tree depths greater than 5.'])
                out.add_blank_row()
                
            out.add_rows([[i] for i in tree_printer.tree_text()], ['courier'])
            out.add_blank_row()
    
    def print_var_importance(self, columns, n=10, out=None):
        if self._model_name != LINEAR_REGRESSION:
            if len(columns) <= EXCEL_INTERFACE.variable_importance_limit:
                # Create a scorer
                scorer = make_scorer(self._eval_func, greater_is_better=True, needs_proba=self._binary_data)
                
                # Find the importances
                importances = sk_i.permutation_importance(self._model, self._X, self._y, scoring=scorer, n_repeats=n, random_state=self._seed)
                
                # Make a table
                df_imp = pd.DataFrame({'Variable':columns})
                df_imp['Importance'] = importances.importances_mean
                df_imp['Importance ' + chr(963)] = importances.importances_std/np.sqrt(n)
                
                df_imp = df_imp.sort_values('Importance', ascending=False)
                
                # Add the table; if it is successfully printed, print the variable importance
                # graph
                if out.add_table(df_imp):
                    # Get the ones we care about
                    relevant_var_ids = df_imp.index
                    
                    fig = plt.figure(figsize=(7,(len(relevant_var_ids)+2)/EXCEL_INTERFACE.graph_line_per_inch), facecolor='blue')
                    ax = fig.add_axes((0, 0, 1, 1))
                    
                    for var_n, var_id in enumerate(reversed(relevant_var_ids)):
                        ax.plot( [df_imp.loc[var_id, 'Importance']],
                                    [var_n],
                                    #[df_imp.loc[var_id, 'Variable']],
                                    marker='*',
                                    color='black',
                                    markersize=10)
                                    
                        ax.plot( importances.importances[var_id, :],
                                    [var_n]*n,
                                    #[df_imp.loc[var_id, 'Variable']]*n,
                                    marker='.',
                                    color='red',
                                    markersize=5,
                                    linewidth=0 )
                        
                    ax.set_xlabel('Decrease in in-sample score', fontsize=18)
                    ax.set_yticks(range(len(df_imp)))
                    ax.set_yticklabels(reversed(df_imp.Variable.tolist()), fontsize=18)
                    sns.despine()
                    
                    out.add_graph(fig, 2, manual_shift=(3,-len(df_imp)))
                    
                    out.add_blank_row()
                    out.add_blank_row()
            else:
                out.add_row('To save time, variable importance measures are not calculated for models with '
                              f'more than {EXCEL_INTERFACE.variable_importance_limit} variables.')
                out.add_blank_row()
                
"""
def listify_tree(estimator, node_id):
    '''
    Convert a tree into a list for output
    '''
    round_sf = lambda x : 0 if x == 0 else round(x, -int(np.floor(np.log(abs(x))/np.log(10))) + 3)
    
    out_list = [estimator.tree_.feature[node_id], round_sf(estimator.tree_.threshold[node_id])]
    
    if estimator.tree_.children_left[node_id] == estimator.tree_.children_right[node_id]:
        if len(estimator.tree_.value[0][0]) == 1:
            # We have a regressor
            return round_sf(estimator.tree_.value[node_id][0][0])
        else:
            # We have a classifier
            return round_sf(estimator.tree_.value[node_id][0][1] / sum(estimator.tree_.value[node_id][0]))
    else:
        out_list.append( listify_tree(estimator, estimator.tree_.children_left[node_id] ) )
        out_list.append( listify_tree(estimator, estimator.tree_.children_right[node_id] ) )
        
        return out_list
"""

class TreePrinter:
    '''
    This class provides additional utilities for displaying trees - in both
    textual and graphical format
    '''
    
    def __init__(self, estimator, col_names = None):
        '''
        This function accepts the following arguments
          - estimator : a sklearn tree estimator object
          - col_names : the names of the features in the model, in the order
                        they were provided to the estimator. If blank, column
                        numbers will be used
        
        It will parse the tree to make it ready for printing or plotting as
        required
        '''
        
        # We will use the tree structure in model.tree_ . This is structured as follows:
        #   - .node_count contains the number of notes, including leaf nodes
        #   - .feature contains the feature each node splits at
        #   - .threshold contains the threshold at which each node splits at
        #   - .children_left and .children_right contain the left and right children. Both
        #     will be equal to -1 for leaf nodes
        #   - .value contains one list per node. For a regressor, that list will have the
        #     form [[x]] were x is the predicted value. For a predictor, it will have the
        #     form [[x,y]] where x is the number of points with outcome 0, and y is the
        #     number of points with the outcome 1
        #   - .impurity contains the impurity
        #   - .n_node_samples contains the number of samples at that node
        # The nodes are ordered depth first in the order we would naturally want to print
        # them
        
        # Save the number of nodes
        self._node_count     = estimator.tree_.node_count
        
        # Save the threshold at each node
        self._threshold      = [round_to_string(i, 3) for i in estimator.tree_.threshold]
        
        # Find the predictions at each node, depending on whether the outcome is binary or
        # continuous
        self._node_preds     = [round_to_string(i[0][0],3) if len(i[0])==1
                                    else round_to_string(i[0][1]/sum(i[0]), 3)
                                                  for i in estimator.tree_.value]
        self._n_node_samples = [str(i) for i in estimator.tree_.n_node_samples]
        self._impurities     = [round_to_string(i, 3) for i in estimator.tree_.impurity]
        
        # For each node, save the left and right child
        self._children_left  = estimator.tree_.children_left
        self._children_right = estimator.tree_.children_right
        
        self._node_feature = ['' if i < 0 else (col_names[i] if col_names is not None else str(i))
                                                    for i in estimator.tree_.feature]
        
        # Traverse the tree to find additional node information
        self._traverse_tree()
        
    def _traverse_tree(self):
        '''
        For each node, this function will work out the node depth, whether it is
        a leaf, whether it is to the left of the node above it (<=) or to the
        right (>), and its parent's ID
        '''
        
        # Create placeholders
        self._node_depth  = [None] *self._node_count
        self._is_leaf     = [False]*self._node_count
        self._node_path   = ['']   *self._node_count
        self._node_parent = [None] *self._node_count
        self._node_index  = [None] *self._node_count
        
        # Create a stack that will contain all nodes left to traverse. Each tuple in the
        # stack will contain the node ID and its depth. Seed this stack with the
        # root node
        stack = [(0, 0)]
        self._node_index[0] = 1
        
        while len(stack) > 0:
            node_id, self._node_depth[node_id] = stack.pop()
                        
            left_child, right_child = self._children_left[node_id], self._children_right[node_id]
            
            if left_child == right_child:
                self._is_leaf[node_id] = True
            else:
                self._node_path[left_child],   self._node_path[right_child]   = '<=',    '>'
                
                self._node_parent[left_child], self._node_parent[right_child] = node_id, node_id
                
                self._node_index[left_child]  = 2*self._node_index[node_id] - 1
                self._node_index[right_child] = 2*self._node_index[node_id]
                
                stack.append((left_child,  self._node_depth[node_id]+1))
                stack.append((right_child, self._node_depth[node_id]+1))
        
        # Save the depth of the tree
        self._tree_depth = max(self._node_depth)
    
    def tree_text(self):
        '''
        This function will return a text-based version of the tree
        
        The structure of the tree will look something like this
        
                                                               n     imp    pred    
            Split on population ------------------------------ 2177  0.025  0.634  
            -    <= 200395.5 : Split on population ----------- 1952  0.02   0.658  
            -    -    <= 6407.5 : Split on income ------------ 297   0.018  0.747  
            -    -    -    <= 31206.5 : Split on income ------ 15    0.031  0.482  
            -    -    -    -    <= 21051.5 : LEAF NODE ------- 1     0.0    0.838  
            -    -    -    -    > 21051.5 : LEAF NODE -------- 14    0.024  0.456  
        
        Notice there are four sections to this printout:
          - S1: the structure
          - S2: The n column
          - S3: The imp column
          - S4: the val column
        
        The function will return a list, with one line per element.
        '''
        
        # The nodes are ordered depth first in the order we would naturally want to print
        # them.
        
        # Begin with section S1. Start with the root node
        out = [f'Split on {self._node_feature[0]}']
        
        # Add the remaining nodes
        for node_id in range(1, self._node_count):
            out.append('-    '*self._node_depth[node_id]
                          + self._node_path[node_id] + ' '
                          + self._threshold[self._node_parent[node_id]]
                          + ' : Split on '
                          + self._node_feature[node_id])
                          
        # Find the width of each section
        S1_width = max([len(i) for i in out]) + 2
        S2_width = max([len(i) for i in self._n_node_samples])
        S3_width = max([len(i) for i in self._impurities])
        S4_width = max([len(i) for i in self._node_preds])
        
        # Add the remaining sections
        for node_id in range(self._node_count):
            out[node_id] = (   pad_string(out[node_id],                  S1_width, '-')
                             + pad_string(self._n_node_samples[node_id], S2_width, ' ')
                             + pad_string(self._impurities[node_id],     S3_width, ' ')
                             + pad_string(self._node_preds[node_id],     S4_width, ' ') )
        
        # Add a header row
        out.insert(0,    pad_string('',     S1_width, ' ')
                       + pad_string('n',    S2_width, ' ')
                       + pad_string('imp',  S3_width, ' ')
                       + pad_string('pred', S4_width, ' ') )
       
        # Return
        return out
    
    def plot_tree(self):
        '''
        This function will create a matplotlib visualization of the tree that I find
        rather prettier than the default sklearn one. It returns a matplotlib figure
        with the visualization        
        '''
        
        fig, ax = plt.subplots(1, 1, figsize=(22,6))
        ax.invert_yaxis()
        ax.set_frame_on(False)
        ax.get_xaxis().set_visible(False)
        ax.get_yaxis().set_visible(False)
        
        for node_id in range(self._node_count):
            if self._is_leaf[node_id]:
                label = self._node_preds[node_id]
            else:
                label = (self._node_feature[node_id]
                            + (f' > {self._threshold[node_id]}' if '=' not in
                                        self._node_feature[node_id] else ''))
                
            self._place_node(ax,
                              label,
                              self._node_depth[node_id],
                              self._node_index[node_id],
                              self._tree_depth,
                              boxed = (self._is_leaf[node_id]))
        
        return fig
            
    @staticmethod
    def _get_node_pos(node_depth, node_index, tree_depth):
        '''
        This function will return the (x,y) position of a given node on a tree plot.
        See the get_node_pair_pos function for a desription of how nodes are
        specified        
        '''
        
        # Define the margin around our entire plot 
        margin = 0.05        
        
        # Find the height of each tree level
        level_height = (1 - 2*margin)/tree_depth
        
        # Find the width of each node at this specific level
        entry_width = (1 - 2*margin)/np.power(2, node_depth)
        
        # Return the coordinates
        return (margin + (node_index-0.5)*entry_width,
                       margin + node_depth*level_height)
    
    @staticmethod
    def _get_node_pair(node_depth, node_index, tree_depth):
        '''
        This function will return
            (x, y, parent_x, parent_y)
        where (x, y) is the coordinate of the node given, and (parent_x, parent_y)
        is the coordinate if its direct parent, on a canvas with width 1 and
        height 1.
        
        The node is given in terms of a depth and an index, determined as follows
        
                depth                   index
                  0                       1
                  1                 1           2
                  2               1   2       3   4
        
        The function also needs the tree's depth
        '''
        
        # Get the node's position
        out = TreePrinter._get_node_pos(node_depth, node_index, tree_depth)
    
        if node_depth == 0:
            return out + (None, None)
            
        else:
            # Find the parent node
            node_index = np.round((node_index/np.power(2,node_depth))
                                        *np.power(2,node_depth-1)+0.000001)
            node_depth -= 1
            
            # Return the combined coordinates
            return out + TreePrinter._get_node_pos(node_depth, node_index, tree_depth)
            
    def _place_node(self, ax, node_text, node_depth, node_index, tree_depth, boxed=False):
        '''
        This function will place a single node on the tree canvas, given the node
        text, its depth, and index. Nodes indexes are determined as follows
        
                                 1
                           1           2
                         1   2       3   4
                         
        If the node is not the top node, an arrow will be drawn connecting it to the
        levels above.
                         
        The function also expects the axis on which the node should be placed, and
        whether the node should be boxed
        '''
        
        # Get the node coordinates
        x, y, prev_x, prev_y = self._get_node_pair(node_depth, node_index, tree_depth)
    
        # Print the text
        if boxed:
            ax.text( x, y, node_text, ha='center', va='center', size=10,
                       bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        else:
            ax.text( x, y, node_text, ha='center', va='center', size=10)
        
        # If this is not the top node, draw an arrow
        if node_depth > 0:
            # Find the point between this node and the parent node
            mid_x, mid_y = (x+prev_x)/2, (y+prev_y)/2
            
            # Figure out if this node was going down to the left, or to the right
            is_left = (x < prev_x)
            
            # Draw the arrow
            ax.annotate('',
                         xy         = (mid_x + (0.005 if is_left else -0.005), mid_y),
                         xytext     = (prev_x, prev_y+0.04),
                         arrowprops = dict(arrowstyle='-', connectionstyle="angle3, angleA=90, angleB=0"))
            ax.annotate('n' if is_left else 'y',
                         va         = 'center',
                         ha         = 'center',
                         size       = 12,
                         xy         = (x, y-0.04),
                         xytext     = (mid_x, mid_y),
                         arrowprops = dict(arrowstyle='->', connectionstyle="angle3, angleA=0, angleB=90"))
    
    @property
    def tree_depth(self):
        '''
        Returns the tree depth. A tree with a single question has a depth of 1
        '''
        
        return self._tree_depth
        
# =======================
# =  Validation Object  =
# =======================

class TuningObject:

    def __init__(self, addin, datasets, out_err):
        self._out_err = out_err
        self._addin = addin
        self._datasets = datasets

        self._data = D()

        # Add basic columns
        self._data['row']             = D(english_key = 'Row',
                                           vals        = list(range(1,len(addin.tuning_grid)+1)))

        self._data['in_score']        =  D(english_key = 'In-sample score',
                                            vals       = [])

        self._data['out_score']       = D(english_key = 'Out-of-sample score',
                                           vals       = [])

        self._data['out_se']          = D(english_key = 'Out-of-sample SE',
                                           vals       = [])

        self._data['formula']         = D(english_key = 'Formula',
                                           vals       = [])

        # Add columns for each of the parameters
        for param in addin.params:
            self._data[param]         = D(english_key  = addin.params[param].english_key,
                                           vals        = [])

        # Add columns for specific models
        if addin.model_name == LINEAR_REGRESSION:
            self._data['n_coefs']     = D(english_key  = 'Nonzero coefficients',
                                           vals        = [])
        elif addin.model_name in [BOOSTED_DT, RANDOM_FOREST]:
            self._data['opt_n_trees'] = D(english_key  = 'Best number of trees',
                                           vals        = [])
            self._data['staged_graphs'] = D(english_key  = '',
                                             vals        = [])

        # Create a column for iteration errors
        self._data['max_iter_reached'] = D(english_key = '', vals=[])

    def tune_model(self):
        if len(self._addin.tuning_grid) > 0:
            # Create the folds
            self._folds = sk_ms.KFold(n_splits = self._addin.K,
                                        shuffle=True, random_state=self._addin.seed)

            for grid_id, grid_point in enumerate(self._addin.tuning_grid):
                message = ('Tuning the model using K-fold cross validation. Testing parameter '
                                f'combination {grid_id+1} of {len(self._addin.tuning_grid)}.')
                self._addin.update_status(message)
                self._run_grid_point(grid_point, message)

            self._addin.optimal_params = (
                        self._addin.tuning_grid[np.argmax(self._data.out_score.vals)])
                        
            if self._addin.model_name == BOOSTED_DT:
                self._addin.optimal_params['n_estimators'] = self._data.opt_n_trees.vals[np.argmax(self._data.out_score.vals)]

    def _run_grid_point(self, grid_point, message):
        # Save the grid point
        for setting in grid_point:
            self._data[setting].vals.append(grid_point[setting])

        # No iteration problems
        self._data.max_iter_reached.vals.append(False)

        # Get the datasets
        self._datasets.set_formula(grid_point.formula)
        X, y, intercept = self._datasets.X_train, self._datasets.y_train, self._datasets.intercept
        
        # Store evaluation details
        eval = D({'in_sample':[], 'out_sample':[]})
        nonzero_coefs = []
        
        fold_n = 1
        for train_index, test_index in self._folds.split(X):
            self._addin.update_status(f'{message} Running fold {fold_n} of {self._addin.K}.')
        
            # Get the model, re-instantiating for every fold
            model = AddinModel(self._addin.model_name, grid_point,
                                    self._datasets.binary, self._addin.seed, self._out_err)

            # Fit it
            model.fit(X[train_index], y[train_index], intercept)
            nonzero_coefs.append(model.n_nonzero_coefs())

            # Evaluate it
            eval.in_sample.append (model.staged_evaluate(y_true=y[train_index], X=X[train_index]))
            eval.out_sample.append(model.staged_evaluate(y_true=y[test_index],  X=X[test_index]))

            # Check whether we ran out of iterations
            self._data.max_iter_reached.vals[-1] = (self._data.max_iter_reached.vals[-1] or model._max_iter_reached)

            # Iterate the fold number
            fold_n += 1

        # Convert evaluation details to arrays; each row will contain a folds
        # and each column a certain number of trees. Then save the staged graphs
        # and find the optimal number of trees
        eval = D({i:np.array(eval[i]) for i in eval})
        staged_graph = eval.out_sample.mean(axis=0)
        if self._addin.model_name == BOOSTED_DT:
            opt_n_trees = staged_graph.argmax()
        else:
            opt_n_trees = len(staged_graph) - 1

        # Store the results
        se = lambda x : x.std()/np.sqrt(len(x))

        self._data.in_score.vals.append  (   eval.in_sample [:, opt_n_trees].mean())
        self._data.out_score.vals.append (   eval.out_sample[:, opt_n_trees].mean())
        self._data.out_se.vals.append    (se(eval.out_sample[:, opt_n_trees]))

        # Store additional details if required
        if self._addin.model_name == LINEAR_REGRESSION:
            self._data.n_coefs.vals.append(np.mean(nonzero_coefs))
        elif self._addin.model_name in [BOOSTED_DT, RANDOM_FOREST]:
            self._data.opt_n_trees.vals.append(opt_n_trees+1)
            self._data.staged_graphs.vals.append(staged_graph)

    def output_tuning_table(self, out):

        df_tuning = pd.DataFrame({self._data[i].english_key
                                        :self._data[i].vals
                                           for i in self._data
                                             if self._data[i].english_key != ''})

        if any(self._data.max_iter_reached.vals):
            df_tuning['Iteration issues'] = ['*' if i else '' for i in self._data.max_iter_reached.vals]

        out.add_table(df_tuning)
        out.add_blank_row()
        
        if any(self._data.max_iter_reached.vals):
            out.add_header('WARNING', 2, 'During tuning, some of the algorithms that trained your models did not converged within the allocated number of iterations. These are indicated by a *. Seek help!')
            out.add_blank_row()

    def output_tuning_graph(self, out):

        if self._addin.model_name == LINEAR_REGRESSION:
            fig, ax = plt.subplots(1, 2, figsize=(18,6))

            df_n_coefs = (pd.DataFrame({'n_coefs':self._data.n_coefs.vals,
                                        'out_score':self._data.out_score.vals,
                                        'out_se':self._data.out_se.vals})
                            .sort_values('n_coefs'))

            df_n_coefs_g = df_n_coefs[df_n_coefs.out_score == df_n_coefs.groupby('n_coefs').out_score.transform('max')]

            ax[1].errorbar( df_n_coefs_g.n_coefs,
                               df_n_coefs_g.out_score,
                               marker='x',
                               color='black',
                               markersize=10,
                               yerr=1.96*df_n_coefs_g.out_se,
                               capsize=5 )

            ax[1].plot(df_n_coefs.n_coefs,
                           df_n_coefs.out_score,
                           linewidth=0,
                           marker='.',
                           color='red',
                           markersize=12)

            ax[1].set_xlabel('Non zero coefficients', fontsize=18)
            ax[1].set_ylabel('')
            sns.despine()

            ax = ax[0]

        else:
            fig, ax = plt.subplots(1, 1, figsize=(9,6))
            
        if len(self._data.row.vals) >= 200:
            ax.text(0, 0.55, 'This plot is omitted when you are cross-', fontsize=20)
            ax.text(0, 0.45, 'validating more than 100 rows.', fontsize=20)
            ax.axis('off')
            
        else:
            df_rows = pd.DataFrame({'row':self._data.row.vals,
                                        'out_score':self._data.out_score.vals,
                                        'out_se':self._data.out_se.vals})

            ax.bar(      df_rows.row,
                         df_rows.out_score,
                         width=0.4,
                         align='center' )
            ax.errorbar( df_rows.row,
                         df_rows.out_score,
                         yerr=1.96*df_rows.out_se,
                         ls='none',
                         capsize=5,
                         color='red' )

            ax.set_xlabel('Row number', fontsize=18)
            ax.set_ylabel('Out-of-sample score', fontsize=18)
            
            min_point = min(df_rows.out_score - 1.96*df_rows.out_se)
            max_point = max(df_rows.out_score + 1.96*df_rows.out_se)
            margin = max(0.005, 0.1*(max_point-min_point))
            ax.set_ylim((min_point-margin), (max_point+margin))
            if len(df_rows.row) < 20:
                ax.set_xticks(df_rows.row)
                ax.set_xticklabels(df_rows.row)
            sns.despine()

        out.add_graph(fig, indent_level=2)

    def output_staging_graph(self, out):

        if self._addin.model_name in [BOOSTED_DT, RANDOM_FOREST]:
            out.add_row('Ensemble model fitting path:')
            out.add_blank_row()

            fig, ax = plt.subplots(1, 1, figsize=(18,6))

            for row, staged_path in zip(self._data.row.vals, self._data.staged_graphs.vals):
                ax.plot(range(1, len(staged_path)+1), staged_path, label=f'Row {row}')

            ax.legend()
            ax.set_xlabel("Number of trees", fontsize = 18)
            ax.set_ylabel('Out-of-sample score', fontsize=18)

            ax.set_ylim((0, ax.get_ylim()[1]))

            sns.despine()

            out.add_graph(fig, indent_level=2)

    @property
    def tuned(self):
        return len(self._data.row.vals) > 0

# ===================================
# =   Predictive Analytics Add-in   =
# ===================================

def run_addin(function_name, sheet_name, udf_server, workbook=None):
   
    excel_connector = ExcelConnector(workbook)
    
    # Create an error tracking object
    out_err = AddinErrorOutput(sheet_name, excel_connector, function_name)
    
    try:
        # Get the PID and save it
        excel_connector.wb.sheets(EXCEL_INTERFACE.run_id_sheet).range(EXCEL_INTERFACE.pid_cell).value = os.getpid()
    
        if function_name == PREDICTIVE_CONFIG.english_key:
            run_predictive_addin(out_err, sheet_name, excel_connector, udf_server)
            
        elif function_name == TEXT_CONFIG.english_key:
            run_text_addin(out_err, sheet_name, excel_connector, udf_server)
        
        else:
            out_err.add_error("What, what?! How on earth did you get here - please contact "
                                "xlkitlearn@guetta.com to report this has happened!", critical=True)
                                
    except AddinError:
        # If we get an AddinError, we've already handled it
        pass
    
    except BaseException as e:
        if workbook is None:
            try:
                out_err.add_error('An unidentified error has occurred. Please email xlkitlearn@guetta.com '
                                    'with the full text below to report it\n\n Full'
                                    'error:\n\n' + traceback.format_exc(), critical=True)
            except AddinError:
                pass
        else:
            raise
            
def run_predictive_addin(out_err, sheet, excel_connector, udf_server):
    # Step 1 - Setup
    # --------------
    
    # Create output objects to print out results
    out = AddinOutput(sheet, excel_connector)

    # Create a model interface
    addin = PredictiveAddinInstance(excel_connector, out_err, udf_server)
    
    # Step 2 - Validate
    # -----------------
    '''
    addin.update_status('Validating add-in.')
    out.log_event('Validation time')
    verify_addin(out_err, out, settings_cell.split('!')[1])
    out_err.finalize()
    '''
    
    # Step 3 - Load Settings
    # ----------------------
    addin.update_status('Parsing parameters.')
    out_err.add_error_category('Parameter parsing')
    
    addin.load_settings()
    
    out.log_event('Parsing time')
    out_err.finalize()
    
    # Step 3 - Load Datasets
    # ----------------------
    addin.update_status('Loading data.')
    out_err.add_error_category( "Data Reading" )

    ensure_numeric = False
    if any([i[-1]=='.' or i[-3:]=='.-1' for i in addin.formula]):
        y_eq_cols = set([i.split('=')[0].replace('(','').strip() for i in addin.formula if '=' in i.split('~')[0]])
        if len(y_eq_cols) != 1:
            ensure_numeric = True
        else:
            ensure_numeric = y_eq_cols.pop()
    
    datasets = Datasets(addin, out_err, excel_connector, ensure_numeric)

    out.log_event('Data loading')
    out_err.finalize()

    # Step 4 - Tune our model
    # -----------------------
    addin.update_status('Checking if K-fold cross validation is needed.')
    
    # If we're doing best-subset selection, make sure we don't have categoricals
    if addin.best_subset:
        datasets.set_formula(addin.formula[-1])
        if datasets.X_train.shape[1] != addin._n_x_terms:
            out_err.add_error('You are trying to do best-subset selection, but one of your '
                                'variables is a categorical variable with more than two levels. '
                                'In these instances, it is unclear what you want to do - '
                                'individually add/remove each level of the variable, or the whole '
                                'variable itself. Please manually define each level you want to '
                                'add/remove in your formula.', critical=True)
    
    tuning = TuningObject(addin, datasets, out_err)
    tuning.tune_model()
        
    out.log_event('Tuning time')
    out_err.finalize()
    
    # Step 5 - Fit the full model
    # ---------------------------
    addin.update_status('Fitting the full model.')
    
    datasets.set_formula(addin.optimal_params.formula)
    
    model = AddinModel(addin.model_name,
                        addin.optimal_params,
                        datasets.binary,
                        addin.seed,
                        out_err).fit(datasets.X_train, datasets.y_train, datasets.intercept)
    
    out.log_event('Model fit time')
    out_err.finalize()
    
    # Step 6 - Perform the Model Evaluation
    # -------------------------------------
    if datasets.X_eval is not None:
        addin.update_status('Performing model evaluation.')
    
        y_eval_pred = model.predict(datasets.X_eval)
        score_eval = model.evaluate(y_true=datasets.y_eval, y_pred=y_eval_pred)
        
        out.log_event('Model evaluation time')
        out_err.finalize()
    
    # Step 7 - Perform any required predictions
    # -----------------------------------------
    if datasets.X_pred is not None:
        addin.update_status('Performing predictions on new data.')
    
        y_pred = model.predict(datasets.X_pred)
    
        out.log_event('Prediction time')
        out_err.finalize()
    
    # ====================
    # =  Output result   =
    # ====================
    
    # Step 1 - Title
    # --------------
    
    if addin._v_message != '':
        split_message = wrap_line(addin._v_message, EXCEL_INTERFACE.output_width)
        for mess_line in split_message.split('\n'):
            out.add_header(mess_line, 0)
        out.add_blank_row()
        out.add_blank_row()
    
    out.add_header('XLKitLearn Output', 0)
    out.add_header('Model', 2, addin.english_model_name)
    out.add_header('Outcome type', 2, 'Binary' if datasets.binary else 'Continuous')
    out.add_header('Seed', 2, addin.seed)
    out.add_header('Dataset type', 2, 'Sparse' if datasets.is_sparse else 'Dense')
    out.add_blank_row()
    
    # Step 2 - Tuning
    # ---------------
    
    if tuning.tuned:
        addin.update_status('Reporting on model tuning')
        
        out.add_header('Parameter tuning', 1)
        out.add_row(f'Cross validation carried out with {addin.K} folds')
        out.add_blank_row()
        
        tuning.output_tuning_table(out)
        tuning.output_tuning_graph(out)
        tuning.output_staging_graph(out)
    
    # Step 3 - Model
    # --------------
    
    if addin.output_model:
        addin.update_status('Reporting on final model')
    
        out.add_header('Model', 1)
        
        for param in addin.optimal_params:
            if param == 'formula':
                out.add_header('Formula', 2, addin.optimal_params[param])
            else:
                out.add_header(addin.params[param].english_key, 2, addin.optimal_params[param])
        
        out.add_blank_row()
        out.add_header('In-sample score', 2, str(model.evaluate(datasets.y_train, X=datasets.X_train)))
        out.add_blank_row()
        
        model.print_reg_results(datasets.columns, get_p_vals=(not tuning.tuned) and (not datasets.is_sparse), out=out)
                
        model.print_tree(datasets.columns, out)
        
        if datasets.is_sparse and (addin.model_name != LINEAR_REGRESSION):
            out.add_row('Variable importances are not available for models fit using sparse datasets.')
            out.add_blank_row()
        else:
            model.print_var_importance(datasets.columns, n=10, out=out)
        
    # Step 4 - Evaluation
    # -------------------
    
    if (datasets.X_eval is not None):
        addin.update_status('Reporting on model evaluation')
        
        out.add_header('Model Evaluation', 1)
        
        if addin.evaluation_perc is None:
            out.add_row('Evaluation set provided was used.')
        else:
            out.add_row(f'Evaluation set created with {int(addin.evaluation_perc*100)}% of training data.')
        
        out.add_blank_row()
        
        out.add_header('Out-of-sample score', 2, score_eval)
        
        out.add_blank_row()
        
        model.print_roc_curve(y_true=datasets.y_eval, y_pred=y_eval_pred, out=out)
        
        if addin.output_evaluation_details:
            out.add_table(pd.DataFrame({'Original row':datasets.eval_rows,
                                           'True outcome' : datasets.y_eval,
                                           'Predicted outcome' : y_eval_pred}))
                                        
        out.add_blank_row()
    
    # Step 4 - Prediction
    # -------------------

    if datasets.X_pred is not None:
        addin.update_status('Reporting on model predictions')
        
        out.add_header('Model Predictions', 1)
        
        out.add_table(pd.DataFrame({'Row' : datasets.pred_rows,
                                        'Predicted outcome' : y_pred}))
        out.add_blank_row()
    
    # Step 5 - Code
    # -------------
    if addin.output_code:
        addin.update_status('Preparing code')
        
        out.add_header('Equivalent Python code', 1)
        
        out.add_row(PredictiveCode(addin, datasets, excel_connector).code_text, format='courier', split_newlines=True)
        
        out.add_blank_row()
    
    addin.update_status('Pushing results to Excel')
    
    out.finalize(addin.settings_string)

class TextCode:
    def __init__(self, addin):
        pass
    
        self._addin = addin
        self._import_statements = ['import pandas as pd']
        
        self.code_text = self._load_data() + self._vectorize()
        
        if addin.run_lda:
            self.code_text += self._run_lda()
        
        self.code_text = self._import() + self.code_text

    def _import(self):
        o = ''
        o +=                       '# ======================'                                                          +'\n'
        o +=                       '# =  Import packages   ='                                                          +'\n'
        o +=                       '# ======================'                                                          +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '\n'.join(remove_dupes(self._import_statements))                                    +'\n'
        o +=                       ''                                                                                  +'\n'
        
        return o
        
    def _load_data(self):
        o = ''
        
        o +=                       '# Store the name of the file the text is stored in'                                +'\n'
        o +=                      f'file_name = "{self._addin.source_data}"'                                           +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# ==============='                                                                 +'\n'
        o +=                       '#   Load Data   ='                                                                 +'\n'
        o +=                       '# ==============='                                                                 +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       'with open(file_name) as f:'                                                        +'\n'
        o +=                       '    # Read the entire file'                                                        +'\n'
        o +=                       '    data = f.read()'                                                               +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '    # Split it by newline'                                                         +'\n'
        o +=                       '    data = data.split("\\n")'                                                      +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '    # Keep the entire file except for blank lines'                                 +'\n'
        o +=                       '    data = [i for i in data if i.strip() != ""]'                                   +'\n'
        o +=                       ''                                                                                  +'\n'

        return o
    
    def _vectorize(self):
        o = ''
        
        o +=                       '# ========================'                                                        +'\n'
        o +=                       '#   Vectorize the Text   ='                                                        +'\n'
        o +=                       '# ========================'                                                        +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Create a function to tokenize text - this function will take a piece of text'
        if self._addin.stem:
            o +=                                                                                                 ','   +'\n'
            o +=                   '# split it into words, and stem each word.'                                        +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'stemmer = nltk.SnowballStemmer("english")'                                         +'\n'
            o +=                   'tokenize = lambda x : [stemmer.stem() for i in x.split()]'                         +'\n'
            self._import_statements.append('import nltk')
        else:
            o +=                                                                                                 ' and'+'\n'
            o +=                   '# split it into words.'                                                            +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'tokenize = lambda x : [i for i in x.split()]'                                      +'\n'
        o +=                       ''                                                                                  +'\n'
        
        vect_e = 'TF-IDF' if self._addin.tf_idf else 'count'
        vect = 'TfidfVectorizer' if self._addin.tf_idf else 'CountVectorizer'
        
        o +=                       '# Initialize a {vect_e} vectorizer.'
        if self._addin.bigrams:
            o +=                                                      ' The argument n_gram_range = (1, 2) means we'   +'\n'
            o +=                   '# want to select groups of 1 and 2 words. An argument of (1, 3) would pick groups' +'\n'
            o +=                   '# of 1, 2, and 3 words. An argument of (2, 3) would pick groups of 2 and 3 words.' +'\n'
        else:
            o +=                                                    ''                                                 +'\n'
        
        vectorizer_settings = [f'max_features={self._addin.max_features}']
        
        if self._addin.bigrams:
            vectorizer_settings.append('ngram_range=(1, 2)')
        
        if self._addin.stop_words:
            vectorizer_settings.append('stop_words="english"')
        
        if self._addin.max_df != 1.0:
            vectorizer_settings.append(f'max_df={self._addin.max_df}')
        
        if self._addin.min_df != 1:
            vectorizer_settings.append(f'min_df={self._addin.min_df}')
            
        o +=                      f'vectorizer = sk_t.{vect}('
        o +=                               ',\n                                   '.join(vectorizer_settings) + ')'    +'\n'
        self._import_statements.append('from sklearn.feature_extraction import text as sk_t')
        
        if self._addin.eval_perc == 0:
            o +=                   '# We just want to vectorize the whole dataset - no need to worry about future'     +'\n'
            o +=                   '# test sets'                                                                       +'\n'
            o +=                   'X = vectorizer.fit_transform(data)'                                                +'\n'
            
        else:
            self._import_statements.append('import sklearn.model_selection as sk_ms')
            self._import_statements.append('import scipy as sp')
            
            o +=                   '# We need to split our data into a training and test set before vectorizing. Begin'+'\n'
            o +=                   '# by creating a list of document numbers (1 through number of documents) and'      +'\n'
            o +=                   '# splitting *that* list using train_test_split.'                                   +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'doc_nums = list(range(len(data)))'                                                 +'\n'
            o +=                   'train, evaluation = sk_ms.train_test_split( doc_nums,'                             +'\n'
            o +=                  f'                                            train_size = 1 - {self._addin.eval_perc},'+'\n'
            o +=                  f'                                            test_size = {self._addin.eval_perc},'+'\n'
            if self._addin.seed is not None:
                o +=              f'                                            random_state = {self._addin.seed},'    +'\n'
            o +=                   '                                            shuffle = True)'                       +'\n'
            o +=                   ''
            o +=                   '# Fit the vectorizer on the training data'                                         +'\n'
            o +=                   'X_train = vectorizer.fit_transform( [data[i] for i in train] )'                    +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '# Now, use that vectorizer on the test data'                                       +'\n'
            o +=                   'X_evaluation = vectorizer.transform([data[i] for i in evaluation])'                +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '# We now have vectorized data, but we need to make sure we return it in the'       +'\n'
            o +=                   '# *original* order that data was in, so that when we later do a train/test'        +'\n'
            o +=                   '# split with the same seed, we get the same train/test sets'                       +'\n'
            o +=                   'correct_order = list(enumerate(train + evaluation))'                               +'\n'
            o +=                   'correct_order.sort(key = lambda i : i[1])'                                         +'\n'
            o +=                   'correct_order = [i[0] for i in correct_order]'                                     +'\n'
            o +=                   'X = sp.sparse.vstack([X_train, X_evaluation])[correct_order, ]'                    +'\n'
            
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Keep track of the dictionary in the order of the columns in X'                   +'\n'
        o +=                       'vocab = sorted(vectorizer.vocabulary_, key = lambda i : vectorizer.vocabulary_[i])'+'\n'
        o +=                       ''                                                                                  +'\n'
        
        return o
        
    def _run_lda(self):
        o = ''
        
        self._import_statements.append('from sklearn import decomposition as sk_d')
        o +=                       '# ====================================='                                           +'\n'
        o +=                       '# =  Run Latent Dirichlet Allocation  ='                                           +'\n'
        o +=                       '# ====================================='                                           +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Run LDA'                                                                         +'\n'
        
        lda_args = [f'n_components={self._addin.lda_topics}']
        if self._addin.seed is not None:
            lda_args.append(f'random_state={self._addin.seed}')
        
        o +=                       'lda = sk_d.LatentDirichletAllocation(' + ','.join(lda_args) + ')'                  +'\n'
        o +=                       'lda.fit(X)'                                                                        +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Get the list of words in each topic; the default matrix LDA spits out'           +'\n'
        o +=                       '# includes *every* word, and indexes them by word number. We want to keep'         +'\n'
        o +=                       '# the top 15 words only, and see the actual words themselves, not the word'        +'\n'
        o +=                       '# number'                                                                          +'\n'
        o +=                       'topics_matrix = lda.components_'                                                   +'\n'
        o +=                       'topics = []'                                                                       +'\n'
        o +=                       'for topic in topics_matrix:'                                                       +'\n'
        o +=                       '    topics.append([ vocab[i] for i in topic.argsort()[ : -15 : -1] ])'             +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       'topics = pd.DataFrame(np.array(topics).transpose(),'                               +'\n'
        self._import_statements.append('import numpy as np')
        o +=                       '                           columns = ["Topic "+str(i)'                             +'\n'
        o +=                      f'                                for i in range({self._addin.lda_topics})])'        +'\n'
        o +=                       'topics = topics.reset_index().rename(columns = {"index":""})'                      +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Get the topic distribution by document'                                          +'\n'
        o +=                       'doc_topics = lda.transform(X)'                                                     +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       'doc_topics = pd.DataFrame(doc_topics,'                                             +'\n'
        o +=                       '                           columns=[f"Topic {i}" for i in range(' + str(self._addin.lda_topics) + ')])'+'\n'
        o +=                       'doc_topics = doc_topics.reset_index().rename(columns={"index":"Doc Number"})'      +'\n'
        
        return o
        
class PredictiveCode:

    def _get_load_dataset_code(self, output_file, range):
        o = ''
        if 'cell' in range:
            self._import_statements.append('import xlwings as xw')
            o +=                   output_file + ' = ( xw.Book("' + self._excel_connector.wb.name + '")'               +'\n'
            o +=                   '                                     .sheets("' + range.sheet + '")'               +'\n'
            o +=                   '                                     .range("' + range.cell + '")'                 +'\n'
            o +=                   '                                     .options(pd.DataFrame, index=0, header=1)'    +'\n'
            o +=                   '                                     .value )'                                     +'\n'
        else:
            if '.csv' in range.file:
                o +=                   output_file + ' = pd.read_csv("' + range.file + '")'                              +'\n'
            else:
                o +=                   output_file + ' = pd.read_excel("' + range.file + '")'                              +'\n'
        return o

    def __init__(self, addin, datasets, excel_connector):
        self._addin = addin
        self._excel_connector = excel_connector
        
        self._sparse_data = datasets.is_sparse
        
        self._training_file = addin.training_data
        self._evaluation_file = addin.evaluation_data
        self._prediction_file = addin.prediction_data
        
        self._evaluation_split = (addin.evaluation_perc is not None)
        
        if self._evaluation_split:
            self._evaluation_perc = addin.evaluation_perc
        
        self._needs_evaluation = ((self._evaluation_file is not None) or self._evaluation_split)
        self._needs_prediction = (self._prediction_file is not None)
        
        if self._sparse_data:
            self._has_intercept = (addin.formula[0].replace(' ', '')[-2:] != '-1')
            self._simple_regression = False
        else:
            self._simple_regression = ((addin.model_name == LINEAR_REGRESSION)
                                                    and (not addin.needs_tuning)
                                                    and (addin.params.alpha.vals[0] == 0))
            
            self._dot_formula = False
            self._non_dot_formula = False
            self._eq_y_value = False
            
            for i in addin.formula:
                if i.replace(' ', '').split('~')[1].strip() in ['.', '.-1']:
                    self._dot_formula = True
                else:
                    self._non_dot_formula = True
                
                if '=' in i.split('~')[0]:
                    self._eq_y_value = True
                    
                    lhs = i.split('~')[0].strip()
                    if (lhs[0] == '(') and (lhs[-1] == ')'):
                        lhs = lhs[1:-1]
                    
                    self._eq_y_var = lhs.split('=')[0].strip()
                    self._eq_target_val = lhs.split('=')[1].strip()
            
        self._formula = addin.formula
        
        self._model_name = addin._model_name
        self._params = addin._params
        self._fixed_params = [i for i in addin._params if len(addin._params[i].vals)==1]
        self._tuning_params = [i for i in addin._params if len(addin._params[i].vals)>1]
        self._tuning_grid = addin.tuning_grid
        self._K = addin.K
        
        self._manual_cv = ((len(self._formula) > 1) or (self._model_name == BOOSTED_DT))

        self._binary = datasets.binary
        
        self._seed = addin.seed
        
        # Create a list of import statements
        self._import_statements = ['import pandas as pd']
        
        # Generate the code minus the important statements, which we'll do at the end
        self.code_text = ''
        self.code_text += self._load_data()
        self.code_text += self._base_estimator()
        if len(addin.tuning_grid) > 0:
            self.code_text += self._tuning()
        else:
            self.code_text += self._fit_no_tuning()
        
        if datasets.X_eval is not None:
            self.code_text += self._evaluate()
        
        if datasets.X_pred is not None:
            self.code_text += self._predict()
    
        self.code_text = self._import() + self.code_text
                
    def _import(self):
        o = ''
        o +=                       '# ======================'                                                          +'\n'
        o +=                       '# =  Import packages   ='                                                          +'\n'
        o +=                       '# ======================'                                                          +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '\n'.join(remove_dupes(self._import_statements))                                    +'\n'
        o +=                       ''                                                                                  +'\n'
        
        return o
    
    def _load_data(self):
        o = ''
        
        if not self._sparse_data:
            # Load the dense dataset
            
            o +=                   '# ======================='                                                         +'\n'
            o +=                   '# =  Load the datasets  ='                                                         +'\n'
            o +=                   '# ======================='                                                         +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'raw_datasets = {}'                                                                 +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   self._get_load_dataset_code('raw_datasets["training_data"]', self._training_file)
            o +=                   ''                                                                                  +'\n'
            
            if self._evaluation_file is not None:
                o +=               '# Load the evaluation data'                                                        +'\n'
                o +=               self._get_load_dataset_code('raw_datasets["evaluation_data"]', self._evaluation_file)
                
            if self._prediction_file is not None:
                o +=               '# Load the prediction data'                                                        +'\n'
                o +=               self._get_load_dataset_code('raw_datasets["prediction_data"]', self._prediction_file)
                o +=               ''                                                                                  +'\n'
            
            # If we have a formula with a (y=...)~ outcome, convert the relevant column
            if self._eq_y_value:
                o +=                                                                                                    '\n'
                o +=               '# Our outcome variable is of the form (y=...). We need to edit the outcome column' +'\n'
                o +=               '# to make it binary.'                                                              +'\n'
                o +=              f'y_col = "{self._eq_y_var}"'                                                        +'\n'
                o +=              f'target_val = "{self._eq_target_val}"'                                              +'\n'
                o +=                                                                                                    '\n'
                o +=               'def is_target_val(i):'                                                             +'\n'
                o +=               '    try:'                                                                          +'\n'
                o +=               '        if str(i) == target_val: return 1'                                         +'\n'
                o +=               '        if i == int(i) and str(int(i)) == target_val: return 1'                    +'\n'
                o +=               '    except:'                                                                       +'\n'
                o +=               '        return 0'                                                                  +'\n'
                o +=               ''                                                                                  +'\n'
                o +=                                                                                                    '\n'      
                o +=               'for this_set in raw_datasets:'                                                     +'\n'
                o +=               '    if y_col in raw_datasets[this_set]:'                                           +'\n'
                o +=               '        raw_datasets[this_set][y_col] = (raw_datasets[this_set][y_col]'            +'\n'
                o +=               '                                                    .apply(is_target_val))'        +'\n'
                o +=                                                                                                    '\n'
                
            if self._evaluation_split:
                self._import_statements.append('import sklearn.model_selection as sk_ms')
                o +=               '# Split the training dataset into a training and evaluation set. We use the'       +'\n'
                o +=               '# sklearn function. Note that we need to tell the function to shuffle the data.'   +'\n'
                o +=               'raw_datasets["training_data"], raw_datasets["evaluation_data"] = ('                +'\n'
                o +=               '                      sk_ms.train_test_split(raw_datasets["training_data"],'       +'\n'
                o +=              f'                                               train_size = 1 - {self._evaluation_perc},'+'\n'
                o +=              f'                                               test_size = {self._evaluation_perc},'+'\n'
                if self._seed is not None:
                    o +=          f'                                               random_state = {self._seed},'       +'\n'
                o +=               '                                               shuffle = True) )'                  +'\n'
                o +=               ''                                                                                  +'\n'

            if (not self._simple_regression) or (self._dot_formula):

                # Add the matrix builder

                o +=               '# ============================='                                                   +'\n'
                o +=               '# =  Define a Matrix Builder  ='                                                   +'\n'
                o +=               '# ============================='                                                   +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               '# When the Python script loads the data, it will be in the form of a simple table' +'\n'
                o +=               '# (for example, one column for y, one column for x1, one column for x2.'           +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               '# We will need to convert these dataframe to a matrix that can be used to fit the' +'\n'
                o +=               '# models.'                                                                         +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               '# The function below will handle all this formatting for us. I\'ve attempted to'   +'\n'
                o +=               '# make it as simple as possible based on the formulas you entered, but there'      +'\n'
                o +=               '# may be ways to simplify it even further - I\'ll leave it as an exercise!'        +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               'def process_datasets(datasets, formula):'                                          +'\n'
                o +=               '    """'                                                                           +'\n'
                o +=               '    This function takes two arguments:'                                            +'\n'
                o +=               '      1. A dictionary called datasets. This dictionary must contain an entry'      +'\n'
                o +=               '         called "training_data", and can contain entries called'                   +'\n'
                o +=               '         "evaluation_data" and "prediction_data", if relevant'                     +'\n'
                o +=               '      2. A string called formula, which contains the formula that should be'       +'\n' 
                o +=               '         used to process the datasets'                                             +'\n'
                o +=               '    This function will process all the datasets required using the formula,'       +'\n'
                o +=               '    and return a list containing the following elements'                           +'\n'
                o +=               '      - columns: the columns in the resulting dataframe'                           +'\n'
                o +=               '      - intercept: whether the formula requires an intercept'                      +'\n'
                o +=               '      - training_data: a dictionary containing two elements - X and y - containing'+'\n'
                o +=               '                       the processes dataset.'                                     +'\n'
                o +=               '      - Similar dictionaries called evaluation_data and prediction_data irrelevant'+'\n'
                o +=               ''                                                                                  +'\n'
                o +=               '    Note that there are two slight issues with this way of doing things:'          +'\n'
                o +=               '      - It processes the training, test, and evaluation datasets, even if only one'+'\n'
                o +=               '        of these datasets is needed - this is computationally wasteful. Given'     +'\n'
                o +=               '        xlkitlearn is designed for Excel, however, it is reasonable to assume'     +'\n'
                o +=               '        the datasets involved will be small enough to make this less of'           +'\n'
                o +=               '        an issue.'                                                                 +'\n'
                o +=               '      - If our data processing involves dataset-wide statistics (for example,'     +'\n'
                o +=               '        finding the mean of a variable to standardize it, or finding a variable\'s'+'\n'
                o +=               '        unique values to create categoricals), and our model will be tuned by'     +'\n'
                o +=               '        cross-validation, we should in theory be finding the dataset-wide'         +'\n'
                o +=               '        statistics for every combination of folds we use for training separately'  +'\n'
                o +=               '        rather than for the training dataset as a whole. One way to do this is'    +'\n'
                o +=               '        is using sklearn pipelines.  To keep the code brief, we by-pass this'      +'\n'
                o +=               '        issue.'                                                                    +'\n'
                o +=               '    """'                                                                           +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               '    output_data = {}'                                                              +'\n'
                o +=               ''                                                                                  +'\n'
                
                dot_translato=''

                dot_translato  +=  '    # We have a formula of the form y~., which means we want to include every'     +'\n'
                dot_translato  +=  '    # column. Unfortunately, patsy cannot handle these kinds of formulas, so'      +'\n'
                dot_translato  +=  '    # we manually construct the matrices'                                          +'\n'
                dot_translato  +=  '    y_col = formula.split("~")[0].strip()'                                         +'\n'
                dot_translato  +=  '    x_cols = [i for i in datasets["training_data"].columns if i != y_col]'         +'\n'
                dot_translato  +=  ''                                                                                  +'\n'
                dot_translato  +=  '    output_data["training_data"] = ('                                              +'\n'
                dot_translato  +=  '                {"y" : np.array(datasets["training_data"][y_col].transpose()),'    +'\n'
                dot_translato  +=  '                 "X" : np.array(datasets["training_data"][x_cols])})'              +'\n'
                dot_translato  +=  ''                                                                                  +'\n'
                dot_translato  +=  '    if "evaluation_data" in datasets:'                                             +'\n'
                dot_translato  +=  '        output_data["evaluation_data"] = ('                                        +'\n'
                dot_translato  +=  '                {"y" : np.array(datasets["evaluation_data"][y_col].transpose()),'  +'\n'
                dot_translato  +=  '                 "X" : np.array(datasets["evaluation_data"][x_cols])})'            +'\n'
                dot_translato  +=  ''                                                                                  +'\n'
                dot_translato  +=  '    if "prediction_data" in datasets:'                                             +'\n'
                dot_translato  +=  '        output_data["prediction_data"] = ('                                        +'\n'
                dot_translato  +=  '                {"X" : np.array(datasets["prediction_data"][x_cols])})'            +'\n'
                dot_translato  +=  ''                                                                                  +'\n'
                dot_translato  +=  '    # If the formula is y ~ . -1, we do not want an intercept, otherwise we do'    +'\n'
                dot_translato  +=  '    output_data["intercept"] = (formula[-2:] != "-1")'                             +'\n'
                dot_translato  +=  ''                                                                                  +'\n'
                if self._simple_regression:
                    dot_translato+='    # We\'ll be using statsmodels to fit our model - it expects an intercept in'   +'\n'
                    dot_translato+='    # the X matrix, so let\'s add it if we need it'                                +'\n'
                    dot_translato+='    if output_data["intercept"]:'                                                  +'\n'
                    dot_translato+='        for d in output_data:'                                                     +'\n'
                    dot_translato+='            if d in ["training_data", "evaluation_data", "prediction_data"]:'      +'\n'
                    dot_translato+='                output_data[d]["X"] = output_data[d]["X"].assign(Intercept=1)'     +'\n'
                    dot_translato+=''                                                                                  +'\n'
                dot_translato  +=  '    # Keep track of the column names'                                              +'\n'
                dot_translato  +=  '    output_data["columns"] = x_cols'                                               +'\n'
                                    
                patsy_translator = '    # We have a patsy formula that transforms some columns and/or combines'        +'\n'
                patsy_translator +='    # them (for example, a formula like y ~ C(x1) + x2*x3, which requires us'      +'\n'
                patsy_translator +='    # to create dummies for x1 and multiply x2 and x3). We could process this'     +'\n'
                patsy_translator +='    # formula manually, but it would be cumbersome. Thankfully, the patsy'         +'\n'
                patsy_translator +='    # library (which we loaded above) can take the original table and a'           +'\n'
                patsy_translator +='    # formula and automatically create all these new matrices for us.'             +'\n'
                patsy_translator +='    #'                                                                             +'\n'
                patsy_translator +='    # Not only that, but it can also seamlessly create evaluation matrices'        +'\n'
                patsy_translator +='    # correctly. For exmaple, if we standardize a variable, patsy knows that'      +'\n'
                patsy_translator +='    # when it prepares the evaluation set, it should substract the mean'           +'\n'
                patsy_translator +='    # from the *training* set, not the evaluation set. It\'s a pretty cool'        +'\n'
                patsy_translator +='    # library.'                                                                    +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    # Process the training set. The last argument means that if any data is'       +'\n'
                patsy_translator +='    # missing, an error will be raised'                                            +'\n'
                patsy_translator +='    p = pt.dmatrices(formula, datasets["training_data"], NA_action="raise")'       +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    # By default, patsy will add a column of ones to the output matrix to act'     +'\n'
                patsy_translator +='    # as an intercept. In theory, we could keep this column, and simply use'       +'\n'
                patsy_translator +='    # set_intercept = False in our sklearn functions. At best, this would'         +'\n'
                patsy_translator +='    # result in an extraneous column (for example in K-NN and tree-based'          +'\n'
                patsy_translator +='    # models) and at worse it would result in the intercept being penalized'       +'\n'
                patsy_translator +='    # like any other column when we fit a penalized regression. We therefore'      +'\n'
                patsy_translator +='    # remove the intercept patsy creates. Later, we will keep track of whether'    +'\n'
                patsy_translator +='    # an intercept was needed, and implement it by passing set_intercept = True'   +'\n'
                patsy_translator +='    # to the sklearn function.'                                                    +'\n'
                patsy_translator +='    def remove_intercept(x):'                                                      +'\n'
                patsy_translator +='        if "Intercept" in x.design_info.column_names:'                             +'\n'
                patsy_translator +='            x=np.delete(x, x.design_info.column_names.index("Intercept"), axis=1)' +'\n'
                self._import_statements.append('import numpy as np')
                patsy_translator +='            return x'                                                              +'\n'
                patsy_translator +='        else:'                                                                     +'\n'
                patsy_translator +='            return x'                                                              +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    output_data["training_data"] = {"y" : p[0].transpose()[0],'                    +'\n'
                patsy_translator +='                                    "X" : remove_intercept(p[1])}'                 +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    if "evaluation_data" in datasets:'                                             +'\n'
                patsy_translator +='        output_data["evaluation_data"] = ('                                        +'\n'
                patsy_translator +='            {"y" : pt.build_design_matrices([p[0].design_info],'                   +'\n'
                patsy_translator +='                                            datasets["evaluation_data"],'          +'\n'
                patsy_translator +='                                            NA_action="raise")[0].transpose()[0],' +'\n'
                patsy_translator +='             "X" : remove_intercept(pt.build_design_matrices([p[1].design_info],'  +'\n'
                patsy_translator +='                                            datasets["evaluation_data"],'          +'\n'
                patsy_translator +='                                            NA_action="raise")[0])})'              +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    if "prediction_data" in datasets:'                                             +'\n'
                patsy_translator +='        output_data["prediction_data"] = ('                                        +'\n'
                patsy_translator +='            {"X" : remove_intercept(pt.build_design_matrices([p[1].design_info],'  +'\n'
                patsy_translator +='                                            datasets["prediction_data"],'          +'\n'
                patsy_translator +='                                            NA_action="raise")[0])})'+'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    # Keep track of whether we need an intercept'                                  +'\n'
                patsy_translator +='    cols = p[1].design_info.column_names'                                          +'\n'
                patsy_translator +='    output_data["intercept"] = True if ("Intercept" in cols) else False'           +'\n'
                patsy_translator +='    '                                                                              +'\n'
                patsy_translator +='    # Keep track of the column names (without the intercept)'                      +'\n'
                patsy_translator +='    output_data["columns"] = [i for i in cols if i != "Intercept"]'                +'\n'
                
                if self._non_dot_formula:
                    # If we have a non-dot formula, we'll need Patsy and numpy
                    self._import_statements.append('import patsy as pt')
                    self._import_statements.append('import numpy as np')
                
                if self._eq_y_value:
                    o +=           '    if "=" in formula.split("~")[0]:'                                              +'\n'
                    o +=           '        y_col = formula.split("~")[0].split("=")[0].strip()'                       +'\n'
                    o +=           '        formula = y_col + "~" + formula.split("~")[1]'                             +'\n'
                    o +=                                                                                                '\n'
                
                if self._dot_formula and self._non_dot_formula:
                    # We have both dot and and non-dot formulas; use the right translator
                    o +=           '    if formula.replace(" ", "").split("~")[1] in [".", ".-1"]:'                    +'\n'
                    o +=           '\n'.join(['    ' + i for i in dot_translato.split('\n')])                          +'\n'
                    o +=           '    else:'                                                                         +'\n'
                    o +=           '\n'.join(['    ' + i for i in patsy_translator.split('\n')])                       +'\n'
                elif self._dot_formula:
                    o +=           dot_translato                                                                       +'\n'
                else:
                    o +=           patsy_translator                                                                    +'\n'
                    
                o +=                    ''                                                                             +'\n'
                o +=                    '    return output_data'                                                       +'\n'
                
                if len(self._formula) == 1:
                    o +=           '# We have a single formula - process our datasets using that formula'              +'\n'
                    o +=          f'datasets = process_datasets(raw_datasets, "{self._formula[0]}")'                   +'\n'
                
            else:
                o +=               'datasets = raw_datasets'                                                           +'\n'
                
            o +=               ''                                                                                      +'\n'
        else:
            # Load the sparse dataset
            self._import_statements.append('import scipy.sparse as sp')
            self._import_statements.append('import numpy as np')
            o +=                   '# =============================='                                                  +'\n'
            o +=                   '# =  Load the sparse datasets  ='                                                  +'\n'
            o +=                   '# =============================='                                                  +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'def load_sparse_dataset(df, x_cols, y_col=None):'                                  +'\n'
            o +=                   '    """'                                                                           +'\n'
            o +=                   '    This function takes a DataFrame with three columns - ROW, COLUMN, and VALUE -' +'\n'
            o +=                   '    describing a sparse dataset, together with a y column and a list of x columns.'+'\n'
            o +=                   '    it will return a dictionary with two elements - y will be a dense array, and'  +'\n'
            o +=                   '    X will be a sparse matrix.'                                                    +'\n'
            o +=                   '    """'                                                                           +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    out = {}'                                                                      +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    # Convert row names to row numbers'                                            +'\n'
            o +=                   '    df = (df.merge(pd.DataFrame({"ROW":df.ROW.unique()})'                          +'\n'
            o +=                   '                         .reset_index()'                                           +'\n'
            o +=                   '                         .rename(columns={"index":"new_ROW"}),'                    +'\n'
            o +=                   '                   how="left")'                                                    +'\n'
            o +=                   '            .drop(columns="ROW")'                                                  +'\n'
            o +=                   '            .rename(columns={"new_ROW":"ROW"}))'                                   +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    if y_col is not None:'                                                         +'\n'
            o +=                   '        # Get the y vector'                                                        +'\n'
            o +=                   '        y = np.zeros(df.ROW.max()+1)'                                                +'\n'
            o +=                   '        y.put(df[df.COLUMN == y_col].ROW, df[df.COLUMN == y_col].VALUE)'           +'\n'
            o +=                   '        out["y"] = y'                                                              +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    # Get the X matrix'                                                            +'\n'
            o +=                   '    df = df[df.COLUMN.isin(x_cols)]'                                               +'\n'
            o +=                   '    df = (df.merge(pd.DataFrame({"COLUMN":x_cols,'                                 +'\n'
            o +=                   '                                 "new_COLUMN":list(range(len(x_cols)))}),'         +'\n'
            o +=                   '                   how="left")'                                                    +'\n'
            o +=                   '            .drop(columns="COLUMN")'                                               +'\n'
            o +=                   '            .rename(columns={"new_COLUMN":"COLUMN"}))'                             +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    out["X"] = sp.csr_matrix((df.VALUE, (df.ROW, df.COLUMN)))'                     +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '    return out'                                                                    +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'datasets = {}'                                                                     +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   self._get_load_dataset_code('training_file', self._training_file)
            o +=                   ''                                                                                  +'\n'
            o +=                  f'y_col = "{self._formula[0].split("~")[0].strip()}"'                                +'\n'
            o +=                   'x_cols = [i for i in training_file.COLUMN.unique() if i != y_col]'                 +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   'datasets["training_data"] = load_sparse_dataset(training_file, x_cols, y_col)'     +'\n'
            o +=                   ''                                                                                  +'\n'
            
            if self._evaluation_file is not None:
                o +=               '# Load the evaluation data'                                                        +'\n'
                o +=               self._get_load_dataset_code('evaluation_file', self._evaluation_file)
                o +=               'datasets["evaluation_data"] = load_sparse_dataset(evaluation_file, x_cols, y_col)' +'\n'
                o +=               ''                                                                                  +'\n'
                
            if self._prediction_file is not None:
                o +=               '# Load the prediction data'                                                        +'\n'
                o +=               self._get_load_dataset_code('prediction_file', self._prediction_file)
                o +=               'datasets["prediction_data"] = load_sparse_dataset(prediction_file, x_cols)'        +'\n'
                o +=               ''                                                                                  +'\n'
            
            o +=                   'datasets["columns"] = x_cols'                                                      +'\n'
            o +=                  f'datasets["intercept"] = {self._formula[0].replace(" ", "")[-2:] != "-1"}'          +'\n'
            o +=                   ''                                                                                  +'\n'
            
            if self._evaluation_split:
                self._import_statements.append('import sklearn.model_selection as sk_ms')
                o +=               '# Split the training dataset into a training and evaluation set. We first'         +'\n'
                o +=               '# generate a list of rows, use the sklearn function train_test_split to  separate' +'\n'
                o +=               '# these rows into a training and evaluation set, and finally use these indices'    +'\n'
                o +=               '# to split the datasets themselves'                                                +'\n'
                o +=               'indices = list(range(len(datasets["training_data"]["y"])))'                        +'\n'
                o +=               'train_indices, test_indices = sk_ms.train_test_split(indices,'                     +'\n'
                o +=              f'                                               train_size = 1 - {self._evaluation_perc},'+'\n'
                o +=              f'                                               test_size = {self._evaluation_perc},'+'\n'
                if self._seed is not None:
                    o +=          f'                                               random_state = {self._seed},'       +'\n'
                o +=               '                                               shuffle = True)'                    +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               'datasets["evaluation_data"] = {"X":datasets["training_data"]["X"][test_indices,:],'+'\n'
                o +=               '                               "y":datasets["training_data"]["y"][test_indices]}'  +'\n'
                o +=               ''                                                                                  +'\n'
                o +=               'datasets["training_data"] = {"X":datasets["training_data"]["X"][train_indices,:],' +'\n'
                o +=               '                             "y":datasets["training_data"]["y"][train_indices]}'   +'\n'
                o +=               ''                                                                                  +'\n'
        
        return o
            
    def _base_estimator(self):
        o = ''
        
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       '# =  Create our Base Estimator  ='                                                 +'\n'
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# In this section, we create a base estimator which we will modify as needed in'   +'\n'
        o +=                       '# later parts of the code.'                                                        +'\n'
        o +=                       ''                                                                                  +'\n'
        
        if self._simple_regression:
            if self._binary:
                data_kind = 'binary'
                reg_type = 'logistic'
                fun = 'Logit' if self._dot_formula else 'logit'
            else:
                data_kind = 'continuous'
                reg_type = 'linear'
                fun = 'OLS' if self._dot_formula else 'ols'
            
            o +=                   '# We are fitting a simple linear model. Unfortunately, sklearn does not have a '   +'\n'
            o +=                   '# function that returns p-values, so we use statsmodels instead. We are fitting '  +'\n'
            o +=                  f'# a model with {data_kind} outcomes, and so we use a function for {reg_type}'      +'\n'
            o +=                   '# regression.'                                                                     +'\n'
            o +=                   ''                                                                                  +'\n'
            
            if self._dot_formula:
                self._import_statements.append('import statsmodels.api as sm')
                o +=              f'final_model = sm.{fun}(endog=datasets["training_data"]["y"],'                      +'\n'
                o +=               '                        exog=datasets["training_data"]["X"])'                      +'\n'
            else:
                self._import_statements.append('import statsmodels.formula.api as sm')
                if self._eq_y_value:
                    o  +=          '# We have a y-value of the form column=value. Let\'s create the corresponding'     +'\n'
                    o  +=          '# binary column.'                                                                  +'\n'
                    o  +=         f'formula = "{self._formula[0]}"'                                                    +'\n'
                    o  +=          'y_part = formula.split("~")[0].replace("(", "").replace(")", "")'                  +'\n'
                    o  +=          'y_col = y_part.split("=")[0].strip()'                                              +'\n'
                    o  +=          'target_val = y_part.split("=")[1].strip()'                                         +'\n'
                    o  +=          ''                                                                                  +'\n'
                    o  +=          'def is_target_val(i):'                                                             +'\n'
                    o  +=          '    try:'                                                                          +'\n'
                    o  +=          '        if str(i) == target_val: return 1'                                         +'\n'
                    o  +=          '        if i == int(i) and str(int(i)) == target_val: return 1'                    +'\n'
                    o  +=          '    except:'                                                                       +'\n'
                    o  +=          '        return 0'                                                                  +'\n'
                    o  +=          ''                                                                                  +'\n'
                    o  +=          'for this_set in datasets:'                                                         +'\n'
                    o  +=          '    if y_col in datasets[this_set]:'                                               +'\n'
                    o  +=          '        datasets[this_set] = datasets[this_set].copy().assign('                    +'\n'
                    o  +=          '                       **{y_col: lambda x: x[y_col].apply(is_target_val)})'        +'\n'
                    o  +=          ''                                                                                  +'\n'
                    o  +=          'formula = y_col + "~" + formula.split("~")[1]'                                     +'\n'
                    o  +=          ''                                                                                  +'\n'
                
                o +=              f'final_model = sm.{fun}(formula=formula,'                                           +'\n'
                o +=               '                             data=datasets["training_data"])'                      +'\n'
            o +=                   ''                                                                                  +'\n'
        
        else:
            if self._model_name == LINEAR_REGRESSION:
                if self._binary:
                    o +=           '# We are fitting a logistic regression. In sklearn, the LogisticRegression'        +'\n'
                    o +=           '# class will do this for us, but it\'s a little complicated because sklearn'       +'\n'
                    o +=           '# can do this using a number of algorithms. The default algorthm - lbfgs -'        +'\n'
                    o +=           '# cannot handle Lasso penalities. When those penalties are needed, we must'        +'\n'
                    o +=           '# switch to other penalty types. In particular, we can fit an unpenalized'         +'\n'
                    o +=           '# logistic regression using'                                                       +'\n'
                    o +=           '#        LogisticRegression(penalty="none")'                                       +'\n'
                    o +=           '# and a Lasso-penalized logistic regression using'                                 +'\n'
                    o +=           '#        LogisticRegression(C=1/alpha, penalty="l1", solver="liblinear")'          +'\n'
                    o +=           '# (notice that the class is defined using C=1/alpha, not alpha itself).'           +'\n'
                    o +=           '# XLKitlearn will automatically figure out the right form of the function to use'  +'\n'
                else:
                    o +=           '# We are fitting a linear regression. In sklearn, the LinearRegression class'      +'\n'
                    o +=           '# is the basic linear regression fitter, but it is not able to handle Lasso'       +'\n'
                    o +=           '# penalties. If such penalties are needed, xlkitlearn uses the sklearn Lasso'      +'\n'
                    o +=           '# class.'                                                                          +'\n'
            else:
                o +=              f'# We are fitting a {MODELS[self._model_name].english_key}.'                        +'\n'
            
            o +=                   ''                                                                                  +'\n'
            
            if len(self._fixed_params) > 0:
                o +=               '# We will be using the following parameters:'                                      +'\n'
                if len(self._formula) == 1:
                    o +=          f'#   * Formula : {self._formula[0]}'                                               +'\n'
                for cur_param in self._fixed_params:
                    o +=          f'#   * {self._params[cur_param].english_key} : {self._params[cur_param].vals}'      +'\n'
            
            if len(self._tuning_params) > 0:
                o +=               '# The following parameters will be tuned using K-Fold cross validation:'           +'\n'
                if len(self._formula) > 1:
                    o +=           '#   * Formulas : ' + '\n#              : '.join(self._formula)                    +'\n'
                for cur_param in self._tuning_params:
                    o +=          ('#   * ' + self._params[cur_param].english_key + ' : '
                                                        + str(self._params[cur_param].vals))                           +'\n'
                        
            o +=                    ''                                                                                 +'\n'
            
            if self._manual_cv:
                if (self._model_name == BOOSTED_DT) and (len(self._formula) > 1):
                    o +=           '# Unfortunately, we won\'t be able to use the standard sklearn cross-validation '  +'\n'
                    o +=           '# functionality for two reasons:'                                                  +'\n'
                    o +=           '#   - We have multiple formulas we want to test against each other, and so we'     +'\n'
                    o +=           '#     need to create new data matrices for every fold.'                            +'\n'
                    o +=           '#   - We are using a boosted decision tree, and we will want to monitor the tree'  +'\n'
                    o +=           '#     as we train it so that we can stop when the tree has reached its best'       +'\n'
                    o +=           '#     performance'                                                                 +'\n'
                elif self._model_name == BOOSTED_DT:
                    o +=           '# Unfortunately, we won\'t be able to use the standard sklearn cross-validation '  +'\n'
                    o +=           '# functionality because we are using a boosted decision tree, and we will want to' +'\n'
                    o +=           '# monitor the as we train it so that we can stop when the tree has reached its'    +'\n'
                    o +=           '# best performance.'                                                               +'\n'
                else:
                    o +=           '# Unfortunately, we won\'t be able to use the standard sklearn cross-validation '  +'\n'
                    o +=           '# functionality because we have multiple formulas we want to test against each'    +'\n'
                    o +=           '# other, and so we need to create new data matrices for every fold.'               +'\n'
                
                o +=               ''                                                                                  +'\n'
                o +=               '# We will therefore create a separate estimator for every step in our cross-'      +'\n'
                o +=               '# validation grid.'                                                                +'\n'
            
            else:
                o +=               '# Create a base estimator'                                                         +'\n'
                
                if self._model_name == LINEAR_REGRESSION:
                    # If we're here, the only kind of tuning we're doing is of the Lasso penalty
                    self._import_statements.append('import sklearn.linear_model as sk_lm')
                    if self._binary:
                        if 0 in self._params.alpha.vals:
                            o +=   'base_estimator = sk_lm.LogisticRegression()'                                       +'\n'
                        else:
                            o +=   'base_estimator = sk_lm.LogisticRegression(solver="liblinear")'                     +'\n'
                    else:
                        o +=       'base_estimator = sk_lm.Lasso(normalize=True)'                                      +'\n'
                    
                    o +=           ''                                                                                  +'\n'
                    o +=           '# Remove the intercept if it is not needed'                                        +'\n'
                    o +=           'if not datasets["intercept"]:'                                                     +'\n'
                    o +=           '    base_estimator.fit_intercept = False'                                          +'\n'
                    o +=           ''                                                                                  +'\n'
                                        
                else:
                    base_model = AddinModel._get_model_string(self._model_name,
                                                                D({i:self._params[i].vals[0]
                                                                        for i in self._fixed_params}),
                                                                        self._binary, self._seed)
                    self._import_statements.append(base_model.import_string)
                    o +=           'base_estimator = ' + base_model.model_string                                       +'\n'
            
            o +=                   ''                                                                                  +'\n'
        
        return o
        
    def _tuning(self):
        o = ''
        
        o +=                        '# =============================='                                                 +'\n'
        o +=                        '# =  Perform cross-validation  ='                                                 +'\n'
        o +=                        '# =============================='                                                 +'\n'
        o +=                        ''                                                                                 +'\n'
        
        if self._manual_cv:
            if len(self._formula) == 1:
                o +=                '# Prepare a list of models we will be tuning'                                     +'\n'
                
                grid_items = []
                for grid_item in self._tuning_grid:
                    this_model = AddinModel._get_model_string(self._model_name, grid_item, self._binary, self._seed)
                    self._import_statements.append(this_model.import_string)
                    grid_items.append(this_model.model_string)
                
                o +=                'param_grid = [' + ',\n              '.join(grid_items)+']'                        +'\n'
                                                    
            else:
                o +=                '# Prepare a list for manual cross validation in which each element is a tuple -'  +'\n'
                o +=                '# the first entry in each tuple will be the model we want to test, and the second'+'\n'
                o +=                '# will be the formula we want to use.'                                            +'\n'
                o +=                ''                                                                                 +'\n'
                o +=                'param_grid = ['
                
                grid_tuples = []
                for grid_item in self._tuning_grid:
                    this_model = AddinModel._get_model_string(self._model_name, grid_item, self._binary, self._seed)
                    self._import_statements.append(this_model.import_string)
                    grid_tuples.append('('
                                        + this_model.model_string
                                        + ','
                                        + '"' + grid_item.formula + '"'
                                        + ')')
                
                o +=                ',\n              '.join(grid_tuples) + ']'                                         +'\n'
                
            o +=                    ''                                                                                 +'\n'
            
            o +=                    '# Create a table to hold the cross-validation results'                            +'\n'
            o +=                    'val_results = [{"mean_test_score" :None,'                                         +'\n'
            if len(self._formula) == 1:
                o +=                '                "model"            :i} for i in param_grid]'                      +'\n'
            else:
                o +=                '                "model"            :i,'                                           +'\n'
                o +=                '                "formula"          :j} for i,j in param_grid]'                    +'\n'
            
            o +=                    ''                                                                                 +'\n'
            o +=                    'for row in val_results:'                                                          +'\n'
            
            if len(self._formula) > 1:
                o +=                '    # Create datasets for this formula'                                           +'\n'
                o +=                '    datasets = process_datasets(raw_datasets, row["formula"])'                    +'\n'
                o +=                ''                                                                                 +'\n'
                if self._model_name == LINEAR_REGRESSION:
                    o +=            '    # Remove the intercept if needed; this will modify the model'                 +'\n'
                    o +=            '    if not datasets["intercept"]:'                                                +'\n'
                    o +=            '        row["model"].fit_intercept = False'                                       +'\n'
                    o +=            ''                                                                                 +'\n'
            
            o +=                    '    # Create a list to store each fold\'s out-of-sample score'                    +'\n'
            o +=                    '    scores = []'                                                                  +'\n'
            o +=                    ''                                                                                 +'\n'
            o +=                    '    # Create folds and loop through them'                                         +'\n'
            o +=                   f'    folds = sk_ms.KFold(n_splits={self._K}, shuffle = True'
            o +=                              (f', random_state={self._seed}' if self._seed is not None else '') + ')' +'\n'
            o +=                    '    for train_index, test_index in folds.split(datasets["training_data"]["X"]):'  +'\n'
            o +=                    '        # Start with a clean model object, and fit it to the training folds'      +'\n'
            o +=                    '        model = clone(row["model"])'                                              +'\n'
            o +=                    '        model.fit(datasets["training_data"]["X"][train_index],'                   +'\n'
            o +=                    '                  datasets["training_data"]["y"][train_index])'                   +'\n'
            o +=                    ''                                                                                 +'\n'
            self._import_statements.append('from sklearn import clone')
            
            self._import_statements.append('import sklearn.metrics as sk_m')
            if self._model_name == BOOSTED_DT:
                o +=                '        # Make predictions on the test set. We\'ll use sklearn\'s staged_predict' +'\n'
                o +=                '        # capabilities. This will tell us what the model would predict with one'  +'\n'
                o +=                '        # tree, then two trees, then three trees, etc... We\'ll compare these'    +'\n'
                o +=                '        # predictions to the true outcomes, to obtain a list of performances as'  +'\n'
                o +=                '        # long as the maximum number of trees we tried.'                          +'\n'
                if self._binary:
                    o +=            '        test_preds = [[i[1] for i in j] for j in model.staged_predict_proba'      +'\n'
                    o +=            '                                  (datasets["training_data"]["X"][test_index])]'  +'\n'
                    eval_func = 'sk_m.roc_auc_score'
                else:
                    o +=            '        test_preds = model.staged_predict(datasets["training_data"]'              +'\n'
                    o +=            '                                                        ["X"][test_index])'       +'\n'
                    eval_func = 'sk_m.r2_score'
                o +=               f'        scores.append([{eval_func}(datasets["training_data"]["y"][test_index],'   +'\n'
                o +=                '                                                        i) for i in test_preds])' +'\n'
                o +=                ''
                
            else:
                o +=                '        # Make predictions on the test set, compare them to the true outcomes'    +'\n'
                o +=                '        # append the performance to the scores list.'                             +'\n'
                if self._binary:
                    o +=            '        test_preds = [i[1] for i in model.predict_proba'                          +'\n'
                    o +=            '                           (datasets["training_data"]["X"][test_index])]'         +'\n'
                    eval_func = 'sk_m.roc_auc_score'
                else:
                    o +=            '        test_preds = model.predict(datasets["training_data"]["X"][test_index])'   +'\n'
                    eval_func = 'sk_m.r2_score'
                o +=               f'        scores.append({eval_func}(datasets["training_data"]["y"][test_index],'    +'\n'
                o +=                '                                                                     test_preds))'+'\n'
            o +=                    ''                                                                                 +'\n'
            o +=                    '    # We\'re done looping through the folds. Average the results over all the'    +'\n'
            o +=                    '    # folds.'                                                                     +'\n'
            
            if self._model_name == BOOSTED_DT:
                o +=                '    scores = np.array(scores).mean(axis=0)'                                       +'\n'
                o +=                ''                                                                                 +'\n'
                o +=                '    # Find the optimal number of trees'                                           +'\n'
                o +=                '    best_n_trees = int(np.argmax(scores) + 1)'                                    +'\n'
                o +=                ''                                                                                 +'\n'
                o +=                '    # Alter the model to reflect this number of trees, and track the final'       +'\n'
                o +=                '    # performance of this set of parameters.'                                     +'\n'
                o +=                '    row["model"].n_estimators = best_n_trees'                                     +'\n'
                o +=                '    row["mean_test_score"] = scores[best_n_trees-1]'                              +'\n'
            else:
                o +=                '    row["mean_test_score"] = np.mean(scores)'                                     +'\n'
                
            o +=                    ''                                                                                 +'\n'
            o +=                    '# Print the validation results'                                                   +'\n'
            o +=                    'print(pd.DataFrame(val_results))'                                                 +'\n'
            o +=                    ''                                                                                 +'\n'
            
            o +=                    '# Get the best combination of parameters'                                         +'\n'
            o +=                    'best_params_n = np.argmax([i["mean_test_score"] for i in val_results])'           +'\n'
            
            if len(self._formula) > 1:
                o +=                '# Process the data with the best formula'                                         +'\n'
                o +=                'datasets = process_datasets(raw_datasets, val_results[best_params_n]["formula"])' +'\n'
                
            o +=                    '# Get the best model, and re-train it on the full dataset'                        +'\n'
            o +=                    'final_model = val_results[best_params_n]["model"].fit('                           +'\n'
            o +=                    '                               datasets["training_data"]["X"],'                   +'\n'
            o +=                    '                               datasets["training_data"]["y"])'                   +'\n'
            
        else:
            if (self._model_name == LINEAR_REGRESSION) and (self._binary) and (0 in self._params.alpha.vals):
                o +=                '# We first need to create a grid of parameters we can pass to GridSearchCV. The'  +'\n'
                o +=                '# function can accept parameters in many forms - here, we will pass a list of'    +'\n'
                o +=                '# dictionaries. Each dictionary will be considered individually, and will contain'+'\n'
                o +=                '# one entry for each parameter to be tuned within that dictionary. Every'         +'\n'
                o +=                '# combination of these parameters will be tested. For example,'                   +'\n'
                o +=                '#        [{"tree_depth":[1,2,3], "n_estimators":[100]},'                          +'\n'
                o +=                '#                {"tree_depth":[7,8], "n_estimators":[50, 60]}]'                  +'\n'
                o +=                '# will test the followings pairs of (tree_depth, n_estimators) parameters:'       +'\n'
                o +=                '#        (1, 100), (2, 100), (3, 100), (7, 50), (8, 50), (8, 50), (8, 60)'        +'\n'
                o +=                'param_grid = [{"penalty":["none"]}, {"penalty":["l1"], "solver":["liblinear"],'   +'\n'
                o +=               f'                 "C":1/{[i for i in self._params.alpha.vals if i != 0]}]'         +'\n'
            else:
                o +=                '# We first need to create a grid of parameters we can pass to GridSearchCV. The'  +'\n'
                o +=                '# function can accept parameters in many forms - here, we will simply pass a'     +'\n'
                o +=                '# dictionary in which every entry corresponds to one of the base_estimator\'s'    +'\n'
                o +=                '# parameters. For each'                                                           +'\n'
                o +=                '#         {"tree_depth":[1,2,3], "n_estimators":[50, 100]}'                       +'\n'
                o +=                '# will test the following pairs of (tree_depth, n_estimators) parameters:'        +'\n'
                o +=                '#         (1, 50), (2, 50), (3, 50), (1, 100), (2, 100), (3, 100)'                +'\n'
                if (self._model_name == LINEAR_REGRESSION) and (self._binary):
                    o +=            'param_grid = {"C":[' + ','.join([f'1/{i}' for i in self._params.alpha.vals]) +']}'+'\n'
                else:
                    o +=            'param_grid = {' + ',\n              '.join([f'"{i}":{self._params[i].vals}'                                                           for i in self._tuning_params]) + '}'    +'\n'
            
            o +=                    ''                                                                                 +'\n'
            o +=                    '# Carry out the cross-validation using the GridSearchCV function. Three notes on' +'\n'
            o +=                    '# function\'s parameters:'                                                        +'\n'
            o +=                    '#   * The argument re-fit=True ensures that once the best set of parameters are'  +'\n'
            o +=                    '#     determined, the model is re-trained on the full training set with those'    +'\n'
            o +=                    '#     parameters.'                                                                +'\n'
            o +=                    '#   * The CV parameter specifies the way folds are split for cross-validation.'   +'\n'
            o +=                    '#     There are simpler ways to specify K-fold cross-validation, but the method'  +'\n'
            o +=                    '#     we use here ensures the results will be identical to those calculated in'   +'\n'
            o +=                    '#     Excel.'                                                                     +'\n'
            o +=                    '#   * The scoring parameter specifies how these models should be compared.'       +'\n'
            o +=                    '#     Xlkitlearn only uses two scorers - roc_auc for binary data, and r2 for'     +'\n'
            o +=                    '#     continuous data.'                                                           +'\n'
            o +=                    'grid_search = sk_ms.GridSearchCV( estimator = base_estimator,'                    +'\n'
            o +=                    '                            param_grid = param_grid,'                             +'\n'
            o +=                   f'                            scoring = "{"roc_auc" if self._binary else "r2"}",'   +'\n'
            o +=                    '                            cv = sk_ms.KFold(n_splits=5,'                         +'\n'
            if self._seed is not None:
                o +=               f'                                             random_state = {self._seed},'        +'\n'
            o +=                    '                                             shuffle=True),'                      +'\n'
            o +=                    '                            refit = True,'                                        +'\n'
            o +=                    '                            return_train_score = True)'                           +'\n'
            o +=                    ''                                                                                 +'\n'
            o +=                    '# Fit on our training data'                                                       +'\n'
            o +=                    'grid_search.fit(X=datasets["training_data"]["X"],'                                +'\n'
            o +=                    '                                 y=datasets["training_data"]["y"])'               +'\n'
            o +=                    ''                                                                                 +'\n'
            o +=                    '# Print the results of cross-validation'                                          +'\n'
            o +=                    'print(pd.DataFrame(grid_search.cv_results_)'                                      +'\n'
            o +=                    '                         [["params", "mean_train_score", "mean_test_score"]])'    +'\n'
            o +=                    ''                                                                                 +'\n'
            o +=                    '# Save the final model (which, remember, will have already been trained on the'   +'\n'
            o +=                    '# full training data)'                                                            +'\n'
            o +=                    'final_model = grid_search.best_estimator_'                                        +'\n'

        return o
           
    def _fit_no_tuning(self):
        o = ''
        
        o +=                       '# ==================='                                                             +'\n'
        o +=                       '# =  Fit our model  ='                                                             +'\n'
        o +=                       '# =================='                                                              +'\n'
        o +=                       ''                                                                                  +'\n'
        
        if self._simple_regression:
            o +=                   'final_model = final_model.fit()'                                                   +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '# Print a summary of the model'                                                    +'\n'
            o +=                   'print(final_model.summary())'                                                      +'\n'
        else:
            o +=                   'final_model = base_estimator'                                                      +'\n'
            o +=                   'final_model.fit(datasets["training_data"]["X"], datasets["training_data"]["y"])'   +'\n'
            
        o +=                       ''                                                                                  +'\n'
        
        return o
        
    def _evaluate(self):
        o = ''
        
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       '# =  Evaluate on a new Dataset  ='                                                 +'\n'
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       ''                                                                                  +'\n'
        o +=                       '# Make predictions for the evaluation set'                                         +'\n'
        
        if self._simple_regression and (not self._dot_formula):
            if self._eq_y_value:
                o +=                  f'y_true = datasets["evaluation_data"][y_col]'                                   +'\n'
            else:
                # We used the dataframe as-is, so extract the y-value
                o +=                  f'y_true = datasets["evaluation_data"]["{self._formula[0].split("~")[0].strip()}"]'  +'\n'
        else:
            # We used our formula builder
            o +=                   'y_true = datasets["evaluation_data"]["y"]'                                         +'\n'
        o +=                       ''                                                                                  +'\n'
        
        if self._simple_regression and (not self._dot_formula):
            # We gave a formula directly to statsmodels, so we can just use it to predict
            o +=                   'y_pred = final_model.predict(datasets["evaluation_data"])'                         +'\n'
        elif self._simple_regression:
            # We have stats models (so predict and not predict_proba) but we used our formula builder
            o +=                   'y_pred = final_model.predict(datasets["evaluation_data"]["X"])'                    +'\n'
        else:
            if self._binary:
                # We have sklearn, use predict_proba
                o +=               'y_pred = [i[1] for i in final_model.predict_proba'                                 +'\n'
                o +=               '                              (datasets["evaluation_data"]["X"])]'                 +'\n'
            else:
                o +=               'y_pred = final_model.predict(datasets["evaluation_data"]["X"])'                    +'\n'                
        o +=                       ''                                                                                  +'\n'
        
        self._import_statements.append('import sklearn.metrics as sk_m')
        if self._binary:
            self._import_statements.append('import matplotlib.pyplot as plt')
            o +=                   '# Find the AUC performance on this new data'                                       +'\n'
            o +=                   'print(f"Out of sample AUC: {sk_m.roc_auc_score(y_true, y_pred)}")'                 +'\n'
            o +=                   ''                                                                                  +'\n'
            o +=                   '# Plot the ROC curve'                                                              +'\n'
            o +=                   'fpr, tpr, _ = sk_m.roc_curve(y_true, y_pred)'                                      +'\n'
            o +=                   'plt.plot(fpr, tpr)'                                                                +'\n'
            o +=                   'plt.plot([0,1], [0,1], linestyle="--")'                                            +'\n'
            o +=                   'plt.xlabel("False Positive Rate", fontsize=18)'                                    +'\n'
            o +=                   'plt.ylabel("True Positive Rate", fontsize=18)'                                     +'\n'
            o +=                   'plt.show()'                                                                        +'\n'
        else:
            o +=                   '# Find the R-squared performance on this new data'                                 +'\n'
            o +=                   'print(f"Out of sample R^2: {sk_m.r2_score(y_true, y_pred)}")'                       +'\n'
        
        o +=                       ''                                       
        
        return o
        
    def _predict(self):
        o = ''
        
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       '# =  Predict for a new Dataset  ='                                                 +'\n'
        o +=                       '# ==============================='                                                 +'\n'
        o +=                       ''                                                                                  +'\n'
        if self._simple_regression:
            o +=                   'final_model.predict(datasets["prediction_data"]["X"])'                             +'\n'
        else:
            if self._binary:
                o +=               '[i[1] for i in final_model.predict_proba(datasets["evaluation_data"]["X"])]'       +'\n'
            else:
                o +=               'final_model.predict(datasets["evaluation_data"]["X"])'                             +'\n'
        
        return o
        
# ===================
# =   Text Add-in   =
# ===================

def run_text_addin(out_err, sheet, excel_connector, udf_server):
    # Step 1 - Setup
    # --------------
    
    # Create an output object to print out the result
    out = AddinOutput(sheet, excel_connector)
    
    # Create a model interface
    addin = TextAddinInstance(excel_connector, out_err, udf_server)

    # Step 2 - Validate
    # -----------------
    
    # Step 3 - Load Settings
    # ----------------------
    addin.update_status('Parsing parameters.')
    out_err.add_error_category('Parameter parsing')
    
    addin.load_settings()
    
    out_err.finalize()    
    out.log_event('Parsing time')
    
    # Step 3 - Load the text data
    # ---------------------------
    addin.update_status('Reading data')
    out_err.add_error_category('Data reading')
    
    try:
        excel_path = excel_connector.wb.sheets('code_text').range(EXCEL_INTERFACE.path_cell).value
        file_path = excel_path
        delim = '/' if '/' in file_path else '\\'
        file_path = file_path + delim
        file_path += addin.source_data
        
        with open(file_path, 'r') as f:
            raw_data = [i for i in f.read().split('\n') if i.strip() != '']
            
    except FileNotFoundError:
        error_message = ('The filename you provided does not exist, or could not be found in the '
                             'same directory as this Excel spreadsheet. Here is some info that could '
                             'help with debugging:\n'
                             '  - The file name your entered was\n'
                            f'        {addin.source_data}\n'
                             '  - I\'ve detected the path this spreadsheet sits in as\n'
                            f'        {excel_path}')
                            
        try:
            close_files = [i for i in os.listdir(excel_path)
                            if levenshtein_ratio_and_distance(i.split('.')[0], addin.source_data.split('.')[0]) <= 2]
            error_message += '\n\nThe following files are in the same directory and have similar names. Perhaps you meant '
            error_message += f'to use one of those: {", ".join(close_files)}'
        except:
            pass
        
        out_err.add_error(error_message, critical=True)
        
    except Exception as e:
        out_err.add_error(f'Unknown error reading the data file. The specific error was\n\n{str(e)}')
    
    if len(raw_data) == 0:
        out_err.add_error('It looks like the file you provided contains no data. Please check and try again.')
    
    out_err.finalize()
    out.log_event('Read time')
    
    # Step 4 - Vectorize the data
    # ---------------------------
    addin.update_status('Converting data to a matrix representation')
    out_err.add_error_category('Vectorization')
    
    if addin.stem:
        stemmer = nltk.SnowballStemmer("english")
        tokenize = lambda x : [stemmer.stem(i) for i in x.split()]
    else:
        tokenize = lambda x : [i for i in x.split()]
        
    ngram_range = (0, 2) if addin.bigrams else (0, 1)
    
    if addin.tf_idf:
        vectorizer = f_e.TfidfVectorizer(tokenizer=tokenize, ngram_range = ngram_range)
    else:
        vectorizer = f_e.CountVectorizer(tokenizer=tokenize, ngram_range = ngram_range)
    
    vectorizer.set_params(strip_accents = "ascii")
    
    if addin.stop_words is not None:
        vectorizer.set_params( stop_words = "english" )
    
    vectorizer.set_params( max_features = addin.max_features )
    
    if addin.max_df != 1:
        vectorizer.set_params( max_df = addin.max_df )
    
    if addin.min_df != 0:
        vectorizer.set_params( min_df = addin.min_df )
    
    # Vectorize the data
    if addin.eval_perc == 0:
        X = vectorizer.fit_transform(raw_data)
        
    else:
        train, evaluation = sk_ms.train_test_split( list(range(len(raw_data))),
                                                 train_size = 1 - addin.eval_perc,
                                                 test_size = addin.eval_perc,
                                                 shuffle = True,
                                                 random_state = addin.seed )
    
        X_train = vectorizer.fit_transform( [ raw_data[i] for i in train ])
        X_evaluation = vectorizer.transform( [ raw_data[i] for i in evaluation ] )
    
        correct_order = list(enumerate(train + evaluation))
        correct_order.sort(key = lambda i : i[1])
        correct_order = [i[0] for i in correct_order]
    
        X = sp.sparse.vstack([X_train, X_evaluation])[correct_order, ]
        
    vocab = sorted( vectorizer.vocabulary_, key = lambda i : vectorizer.vocabulary_[i] )
    
    # Log the fact we've finished vectorizing data
    out_err.finalize()    
    out.log_event("Vectorization time")
            
    # Step 5 - do LDA
    # ---------------
    
    # If LDA is required, fit it
    if addin.run_lda:
        addin.update_status('Running LDA')
        out_err.add_error_category('LDA')
        
        lda = LatentDirichletAllocation(n_components=addin.lda_topics, random_state=addin.seed, max_iter=addin.max_lda_iter)
        lda.fit(X)
        
        topics_matrix = lda.components_
        doc_topics = lda.transform(X)
        
        topics = []
        for topic in topics_matrix:
            topics.append( [ vocab[i] for i in topic.argsort()[ : -15 : -1] ] )
        topics = pd.DataFrame( np.array(topics).transpose(), columns = ["Topic " + str(i) for i in range(addin.lda_topics)] )
        topics = topics.reset_index()
        topics = topics.rename(columns = {topics.columns[0]:""})
        
        doc_topics = pd.DataFrame(doc_topics, columns = ["Topic " + str(i) for i in range(addin.lda_topics)])
        doc_topics = doc_topics.reset_index()
        doc_topics = doc_topics.rename(columns = {doc_topics.columns[0]:""})
    
        # Log the fact we've finished fitting LDA
        out_err.finalize()
        out.log_event("LDA fit time")
    
    # Step 6 - Output
    # ---------------
    addin.update_status('Preparing output to spreadsheet')
    
    if addin._v_message != '':
        split_message = wrap_line(addin._v_message, EXCEL_INTERFACE.output_width)
        for mess_line in split_message.split('\n'):
            out.add_header(mess_line, 0)
        out.add_blank_row()
        out.add_blank_row()
    
    out.add_header('XLKitLearn Output', 0)
    
    # Output features/vectorized text
    if not addin.run_lda:
        if addin.sparse_output:
            out_df = pd.DataFrame(D({'ROW':X.nonzero()[0]+1,
                                        'COLUMN':np.array(vocab)[X.nonzero()[1]],
                                        'VALUE':X.data}))
        else:
            out_df = pd.DataFrame(X.toarray(), columns = vocab)
            out_df = out_df.reset_index()
            out_df = out_df.rename(columns = {out_df.columns[0] : "Doc Number"})
            
        out.add_header( "Text Features", 1 )
        out.add_table( out_df )
        out.add_blank_row()
        
    # Output LDA
    if addin.run_lda:
        out.add_header( "LDA Results", 1 )
        out.add_table( topics )
        out.add_blank_row()
        out.add_table( doc_topics )
        out.add_blank_row()

    # Step 5 - Code
    # -------------
    if addin.output_code:
        addin.update_status('Preparing code')
        
        out.add_header('Equivalent Python code', 1)
        
        out.add_row(TextCode(addin).code_text, format='courier', split_newlines=True)
        
        out.add_blank_row()
    
    out.finalize(addin.settings_string)

def kill_addin():
    pid = xw.Book.caller().sheets('code_text').range('C1').value
    try:
        os.kill(int(pid), signal.SIGTERM)
        xw.Book.caller().macro('format_sheet')()
    except:
        pass