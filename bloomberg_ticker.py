# # # # # # # # # # # # # # # #
# Bloomberg Ticker Translator #
# # # # # # # # # # # # # # # #

import openpyxl as xl

# path to excel file containing the dot tickers
file = 'tickers.xlsx'


# exceptions to the algorithm
# format is [[dot ticker, exceptional bbg ticker], ...]:
exceptions = [['XIU 01/01/19 P23', 'XIU 01/01/19 P23 Equity'],
              ['JWEL.CA', 'JWEL CV Equity']]


# components that follow a slash for common stocks:
com_slash = ['A', 'B']

# components that follow a dash for common stocks:
com_dash = ['U', 'I', 'II', 'PB']


# components that follow a slash for preferred stocks:
pfd_slash = ['A']

# components that follow a dash for preferred stocks:
pfd_dash = ['U', 'B']

# components that follow a space for preffered stocks:
pfd_space = ['I', 'II']

################################################################################

# make_bloomberg_ticker(ticker) converts ticker, a dot seperated ticker, to 
#    to a bloomberg ticker that can be used to pull data with bloomberg's
#    excel add-in.
# make_bloomberg_ticker: Str --> Str
def make_bloomberg_ticker(ticker):
    # non-dot ticker instances
    if '.' not in ticker:
        return ticker    
    # 1. separate dot ticker into it's components
    components = []    
    comp_start = 0
    comp_end = 0
    for char in ticker:
        comp_end += 1
        if char == '.':
            components += [ticker[comp_start : comp_end - 1]]
            comp_start = comp_end
        elif '.' not in ticker[comp_start:]:
            components += [ticker[comp_start:]]
            break
    # 2. translate components to bloomberg
    company = components[0]
    country = components[-1]
    if country == 'CA':
        country = 'CN'
    format_core = []    
    core_comps = components[1:-1] 
    # preferred stocks
    if 'PR' in components:
        for comp in core_comps:
            if comp == 'PR':
                continue
            else:
                if comp in pfd_slash:
                    format_core += ['/' + comp]
                elif comp in pfd_dash:
                    format_core += ['-' + comp]
                elif comp in pfd_space:
                    format_core += [' ' + comp]                    
                else:
                    return ticker  
            # 3. piece together the bloomberg ticker
            head = company
            for comp in format_core:
                head += comp
            bbg_ticker = '{} Pfd'.format(head)
            return bbg_ticker                
    # common stocks
    elif ticker not in exceptions:
        for comp in core_comps:
            if comp in com_slash:
                format_core += ['/' + comp]
            elif comp in com_dash:
                format_core += ['-' + comp]
            else:
                return ticker
        # 3. piece together the bloomberg ticker
        head = company
        for comp in format_core:
            head += comp
        bbg_ticker = '{} {} Equity'.format(head, country)
        return bbg_ticker
    # exceptions
    elif ticker in exceptions:
        for override in expections:
            if override[0] == ticker:
                bbg_ticker = override[1]
                return bbg_ticker
    else:
        return ticker

# examples:
# pref is ABC I Pfd
pref = make_bloomberg_ticker('ABC.PR.I.US')
# bee is WOAH/B CN Equity
bee = make_bloomberg_ticker('WOAH.B.CA')


# make_xl_col_bbg(file) replaces dot seperated tickers in file with bbg tickers.
# make_xl_col_bbg: Str --> None
# How to use make_xl_col_bbg:
#    1. copy a column of dot tickers that you are working with in excel
#    2. paste them in an empty xlsx file in col A and save file
#    3. run the file through make_xl_col_bbg
#    4. find the tickers translated to bloomberg tickers in file
#    5. paste the updated tickers in your working file to pull bbg data
# Note: Run one time. Multiple runs will compound the ticker.
# Requires: 
#    - file is an .xlsx file
#    - tickers in file are dot seperated tickers in col A
def make_xl_col_bbg(file):
    wb = xl.load_workbook(filename = file)
    ws = wb.active
    for row in ws.iter_rows(min_row=1, max_col=1):
        for cell in row:
            dot_ticker = cell.value
            bbg_ticker = make_bloomberg_ticker(dot_ticker)
            cell.value = bbg_ticker
    wb.save(file)


# run the program through file file
make_xl_col_bbg(file)