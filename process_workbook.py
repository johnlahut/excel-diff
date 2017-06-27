
# coding: utf-8

# In[5]:

import openpyxl as oxl


# In[11]:

'''
Program Constants
'''

WORKBOOK_PATH = 'test-book.xlsm'

RTE_SHEET = 'RTE Spreadsheet'
SM_SHEET = 'SM Spreadsheet'
MM_SHEET = 'MM Spreadsheet'

RTE_VALIDATION = 'Production MM'
SM_VALIDATION = 'SM Version'
MM_VALIDATION = 'Production MM'

OPERATION_START = 'Full Oper Num'

COL_A = 0
COL_B = 1

RED = 'FFFF0000'


# In[13]:

from . import Route, Operation
def load_route(sheetname):
    '''
    @sheet: 
    '''
    sheet = oxl.load_workbook(WORKBOOK_PATH, read_only=True)[sheetname]
    
    '''
    @report_id:
    @start_reading:
    @start_index:
    '''
    report_id = ''
    start_reading = False
    start_index = -1
    
    '''
    @loop:
    @operation:
    @route:
    '''
    operation = Operation()
    route = Route()
    for index, row in enumerate(sheet.rows):
        
        #Get flow report header store in report_id
        if 'Flow Report' in stringify(row[COL_A]):
            report_id = stringify(row[COL_A])
            
        #If we encounter 'Full Oper Num', set the start index to two rows later
        if stringify(row[COL_B]) == OPERATION_START:
            start_index = index + 2
        
        #First row of operations, set flag to begin reading data
        if index == start_index:
            start_reading = True
            
        #'Read data mode'
        if start_reading:
            
            #If the current operation list isn't empty (accounting for first loop iteration)
            #and the current cell contains an operation number
            #store the operation in the route, clear out temp operation
            if row[COL_B].value != '' and not operation.is_empty():
                if not operation.flagged_for_removal:
                    route.add_operation(operation)
                else:
                    print('Operation flagged for removal. Ignoring operation:', operation)
                operation = Operation()               
                
            #If the current row is not all empty cells, add it to the current operation
            operation.add_row(row)
                
    #Last operation will still be stored but not added
    if not operation.is_empty():
        route.add_operation(operation)  
    route.set_route_id(get_route_id(report_id))
    '''
    Print statements in place of future exception handling. Want to
    throw errors here if the sheets are invalid.
    '''
    
    if sheetname == RTE_SHEET and RTE_VALIDATION not in report_id:
        print('INVALID RTE SHEET')
    elif sheet == SM_SHEET and SM_VALIDATION not in report_id:
        print('INVALID SM SHEET')
    elif sheet == MM_SHEET and MM_VALIDATION not in report_id:
        print('INVALID MM SHEET')
    return route
rte_route = load_route(RTE_SHEET)
sm_route = load_route(SM_SHEET)

for operation in rte_route.operations:
    print(operation)


# In[1]:

#pass in cell, return .value as string with no leading/trailing whitespace
def stringify(cell):
    return str(cell.value).strip()


# In[2]:

def get_route_id(report_id):
    try:
        '''
        Expected report_ids:
        Flow Report (SM Version): [ROUTE_ID], [PRODUCT_ID]
        Flow Report (Production MM): [ROUTE_ID], [PRODUCT_ID]
        '''
        return report_id.split(':')[1].strip().split(',')[0]
    except:
        return 'ErrorReadingRouteID'


# In[3]:

'''
@OrderedDict: Used as main data structure to store operations in the form
    (Key, Value)    ->    (Operation Number, Operation Class) 
'''

from collections import OrderedDict

'''
@Route: Stores all operations for a given route.
'''
class Route():
    
    '''
    @operations: Stores all operations for given route. Incoming values will be Operation type.
    @route_id: ID of the route pulled from the report_id
    @route_type: Will be either RTE (submitted), SM (staged chages), MM (production values)
    '''
    def __init__(self):
        self.operations = OrderedDict()
        self.route_id = None
        self.route_type = None
        
    '''
    @return: returns the route id if not None, else 'NoSetRouteID' as error message
    '''
    def __str__(self):
        if self.route_id != None:
            return self.route_id
        else:
            return 'NoSetRouteID'
    
    '''
    adds an operation to the route
    @param operation: will be of Operation type; operation to be added to route
        assuming operations are in sequential order
    '''
    def add_operation(self, operation):
        self.operations[str(operation)] = operation
        
    '''
    returns the route's operations to the caller
    @return: returns the operations of the route
    '''
    def get_operations(self):
        return self.operations
    
    '''
    sets the route's route ID
    @param route_id: ID of the current route
    '''
    def set_route_id(self, route_id):
        self.route_id = route_id
        
    '''
    sets the route's route type
    @param route_type: type of the current route
    '''
    def set_route_type(self, route_type):
        self.route_type = route_type
    
    '''
    returns total number of operations to the caller
    @return: int of total number of operations
    '''
    def get_num_operations(self):
        return len(self.operations)
    
    


# In[12]:

class Operation():
    
    '''
    constructor
    '''
    def __init__(self):
        self.rows = []
        self.operation_no = None
        self.flagged_for_removal = False
        
    '''
    returns the current operation's full operation number 
        'XXXX.XXXX'
    @return: operation's operation number
    '''
    def __str__(self):
        return self.fix_operation_no(self.rows[0][1].value)
    
    def __eq__(self, other_op):
        
        for index, pairs in enumerate(zip(self.rows, other_op.rows)):
            if (pairs[0] != pairs[1]):
                pairs[0].print_operation()
                pairs[1].print_operation()
    
    '''
    returns true if the current operation has no rows
    @return: True if the operation is empty; False otherwise
    '''
    def is_empty(self):
        return len(self.rows) == 0
    
    '''
    adds a row to the operation,
        detects for None value types (rather than ''),
        ensures the row is now all empty cells
    @param row: row from openpyxl ReadOnlyWorksheet
    '''
    def add_row(self, row):
        if not all(cell.value == None or cell.value == '' for cell in row):
            if any(cell.value == None for cell in row):
                pass
                #print('Detected None cell value.')
            if any(cell.fill.fgColor.rgb == RED for cell in row):
                    self.flagged_for_removal = True
            self.rows.append([cell for cell in row])
            
    '''
    prints the operation to the console
    @param log: log shall be set to None unless called by log_operation
    '''
    def print_operation(self, log=None):
        for row in self.rows:
            print([cell.value for cell in row], file=log)
    
    '''
    logs the operation to the given file
        append mode - if file exists it will be appended to not overwritten
    @param file_name: file_name of the log file
    '''
    def log_operation(self, file_name):
        with open(file_name, 'a+') as log_file:
            self.print_operation(log=log_file)
        log_file.close()
    
    '''
    sets and cleans the operation number incoming from a openpyxl ReadOnlyCell
        needs to cast to string in case Excel converted operation number to float
        ensures the format XXXX.XXXX 
        note that 100.XXXX and 10.XXXX are legal operation numbers
    @param operation_no: incoming cell value from where the operation number was encountered
    '''
    def fix_operation_no(self, operation_no):
        operation_no = str(operation_no)
        
        if '.' not in operation_no:
            operation_no = operation_no + '.0000'
        else:
            while len(operation_no.split('.')[1]) < 4:
                operation_no  = operation_no + '0'
        return operation_no
        
    
    

