command! -nargs=0 Excel call ParseExcel()

if !has("python")
    echo "excel.vim requires support for python"
    finish
endif

au BufRead,BufNewFile *.xls,*.xlam,*.xla,.xlt :call ParseExcel()
au BufRead,BufNewFile *.xlsb,*.xlsx,*.xlsm,*.xltx,*.xltm :call ParseExcel2010()


function! ParseExcel()
set nowrap
:python import xlrd
python << EOF
import vim

# for non-ascii characters
def getRealLength(string):
    length = len(string)
    for s in string:
        if ord(s) > 256:
            length += 1
    return length

# get current file name
vim.command("let currfile = expand('%:p')")
currfile = vim.eval("currfile")

# parse sheets
excelobj = xlrd.open_workbook(currfile)
for sheet in excelobj.sheet_names():
    shn = excelobj.sheet_by_name(sheet)
    sheet = sheet.replace(" ", "\\ ")
    rowsnum = shn.nrows
    if not rowsnum:
        continue
    cmd = "tabedit %s" % (sheet)
    vim.command(cmd)

    for n in xrange(rowsnum):
        line = ""
        for val in shn.row_values(n):
            try:
                val = val.replace('\n',' ')
            except AttributeError as e:
                val = str(val).replace('\n', ' ')
            val = isinstance(val,  basestring) and val.strip() \
                    or str(val).strip()
            line += val + ' ' * (30 - getRealLength(val))
        vim.current.buffer.append(line)

# close the first tab
for i in xrange(excelobj.nsheets):
    vim.command("tabp")
vim.command("q!")

EOF
endfunction

function! ParseExcel2010()
set nowrap
:python import openpyxl
python << EOF
import vim

# for non-ascii characters
def getRealLength(string):
    length = len(string)
    for s in string:
        if ord(s) > 256:
            length += 1
    return length

# Convert NoneType to empty String
def none2str(s):
    if s is None:
        return ''
    elif isinstance(s, basestring): 
        return s
    else:
        return str(s)

# get current file name
vim.command("let currfile = expand('%:p')")
currfile = vim.eval("currfile")

# parse sheets
excelobj = openpyxl.load_workbook(currfile, data_only=True)
for sheet in excelobj:
    cmd = "tabedit %s" % (sheet.title)
    vim.command(cmd)

    for row in sheet.values:
        strLine = ""
        for value in row:
            strLine += none2str(value) + " | "
        vim.current.buffer.append(strLine)

for i in excelobj:
    vim.command("tabp")
vim.command("q!")

EOF
endfunction
