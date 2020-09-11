#!/usr/bin/env python3

import sys
import os
import re
import csv
import time
from subprocess import Popen, PIPE
from os.path import basename, exists
from shutil import move, copy
import shlex
from PyPDF2 import PdfFileWriter, PdfFileReader

# Check the input csv and message files
def checkfiles(files_passed):

    for key, ifile in files_passed.items():
        # If variable for the file is passed
        if ifile:
            if not exists(ifile):
                print("Required file '%s' not found. Aborting." % ifile)
                sys.exit(1)
        else:
        # If the requried variable is not passed
            print("Required '%s' file not supplied. Aborting." % key)
            sys.exit(1)


## Overriding csv.DictReader to strip whitespace and ignore case
class MyDictReader(csv.DictReader, object):
    @property
    def fieldnames(self):
        return [field.strip() for field in super(MyDictReader,
            self).fieldnames]

# Read csv file
def csv_dict_reader(infile):
    reader = MyDictReader(open(infile), delimiter=',',skipinitialspace=True)
    return reader

def number_of_data(incsv):

    ndata = 0
    with open(csvfile) as infp:
        for line in infp:
            if line.strip():
                ndata += 1

    # Deleting the header
    ndata -= 1
    return ndata

# Get the path of any executable
def which(pgm):
    path = os.getenv('PATH')
    pathext = os.environ.get("PATHEXT", "").split(os.pathsep)
    for p in path.split(os.path.pathsep):
        p = os.path.join(p,pgm)
        for ext in pathext:
            fp = p + ext.lower()
            if exists(fp) and os.access(fp,os.X_OK):
                return fp


# Check for the list of required softwares
def check_software(soft_list):

    myname = sys._getframe().f_code.co_name

    if type(soft_list) != type([]):
        raise TypeError("%s requires a list, but was given %s" 
            % (myname, type(soft_list)))

    for soft in soft_list:
        path = which(soft)
        if not path:
            sys.exit("%s not installed. Aborting." %(soft)) 
    

# Get root of the tex file
def get_root_tex(txfile):
    
    rttex = txfile
    rttex_user = None

    with open (rttex, "r") as rttexfh:
        for i in range(5):
            line = rttexfh.readline().rstrip().strip()
            if re.search("TeX[ \t]*root[ \t]*=",line):
                rttex_user = re.sub(".*=",'',line).strip()

    if rttex_user != None:
        rttex = rttex_user

    if not os.path.isfile(rttex) :
        sys.exit("File '%s' not found." % rttex)

    return rttex

# Get the tex program
def get_tex_program(rttex):
    
    txprog = 'pdflatex'

    # Open tex file
    with open (rttex, "r") as rttexfh:
        for i in range(5):
            line = rttexfh.readline().rstrip().strip()
            if re.search("TeX[ \t]*program[ \t]*=",line):
                txprog = re.sub(".*=",'',line).strip()

    return txprog

# Set the options of the tex program
def set_tex_program(rttex,txprog):
    
    ltxcmp = None
    ltxcmp_opts = (['-src-specials', '-file-line-error', '-synctex=1', 
        '-interaction=nonstopmode', '-shell-escape']) 

    # Remove all whitespace and change to lowercase
    txprog = re.sub(' ','',txprog).lower()
    progs = txprog.split('+')
    check_software(progs)

    if ( txprog == 'pdflatex' ):
        ltxcmp = ("-pdf -pdflatex=\"pdflatex " + ' '.join(ltxcmp_opts) 
            + " %O %S\"")

    elif ( txprog == 'latex+dvipdf' ): 
        ltxcmp = ("-pdfdvi -latex=\"latex " + ' '.join(ltxcmp_opts) 
            + " %O %S\"")

    elif ( txprog == 'latex+dvips+ps2pdf' ):
        sys.exit('Better use "%!TeX program = latex+dvipdf". Aborting.')
#         ltxcmp = ("-pdfps -latex=\"latex " + ' '.join(ltxcmp_opts) 
#             + " %O %S\"")

    elif ( txprog == 'xelatex' ):
        ltxcmp = ("-xelatex -pdflatex=\"xelatex " + ' '.join(ltxcmp_opts) 
            + " %O %S\"")

    elif ( txprog == 'latexmk' ):
        ltxcmp = ("-pdf " + ' '.join(ltxcmp_opts)) 
#             + " %O %S")

    else:
        sys.exit("The TeX program \"%s\" not supported." % txprog) 

    return ltxcmp


def error_msg(ext_stat,ltx,texfl):
    # Exit if return code is not zero.
    if (ext_stat != 0):
        print('Problem in running "%s" on "%s". Aborting.' 
            %(ltx, texfl))
        sys.exit(1)


def clean_aux(filename,exts,rmtex):
    # Clean files
    for ext in exts:
        delfile = filename + ext
        if exists(delfile):  
            os.remove(delfile)

    if rmtex:
        delfile = filename + '.tex'
        if exists(delfile):  
            os.remove(delfile)


def helptext(sname,au,ver):
    print("Script to create mail merge pdf files from tex file")
    print("Author: %s" % au)
    print("Version: %s" % ver)
    print("Usage:", end='')
    print("%s [options] file1.tex file2.csv -o "
        "\"<unique_header_in_the_csv_file>\"" % sname)
    print("Options:")
    print("-h|--help    Show this help and exit.")
    print("-o|--output  <unique_header_in_the_csv_file>")
    print("             To create individual files with filenames being")
    print("             the values of the unique header.")
    return


if __name__ == "__main__" :
   
    scriptname = basename(sys.argv[0])
    author = 'Anjishnu Sarkar'
    version = 0.5
    pdf_dir = 'pdf'
    latexmk = 'latexmk'
    dflt_make_opts = '-g -silent' ;
    
    extensions = (['.aux', '.log', '.out', '.fls', '.synctex.gz',
             '.fdb_latexmk', '.ps', '.dvi'])
    texdel = False

    csvfile = None
    template_texfile = None
    unique_hdr = None

    # Number of arguments supplied via cli
    numargv = len(sys.argv)-1
    # Argument count
    iargv = 1

    # Parse cli options
    while iargv <= numargv:    

        if sys.argv[iargv] == "-h" or sys.argv[iargv] == "--help":
            helptext(scriptname,author,version)
            sys.exit(0)

        elif re.search('.csv$',sys.argv[iargv]):
            csvfile = sys.argv[iargv]

        elif re.search('.tex$',sys.argv[iargv]):
            template_texfile = sys.argv[iargv]

        elif sys.argv[iargv] == '-o' or sys.argv[iargv] == "--output":
            output_template = sys.argv[iargv+1]
            iargv += 1  

        else:
            print ("%s: Unspecified option: '%s'. Aborting." 
                % (scriptname, sys.argv[iargv]))
            sys.exit(1)
        iargv += 1 

    # Check the input files
    allfiles = {'csv': csvfile, 'tex': template_texfile}
    checkfiles(allfiles)

    # If output template is not defined
    if not output_template:
        print("Output template not defined. Aborting.")
        sys.exit(1)

    check_software([latexmk])

    roottex = get_root_tex(template_texfile) 
    texprogram = get_tex_program(roottex)
    ltxcompiler = set_tex_program(roottex,texprogram) 

    # Read csv and store as dictionary
    dict_reader = csv_dict_reader(csvfile)
    csv_headers = dict_reader.fieldnames

    # Open tex file
    with open (template_texfile, "r") as intexfh:
        str_msg = intexfh.read()
    intexfh.close()

    # Create output directory
    if not exists(pdf_dir):
        os.makedirs(pdf_dir)

    # Find the number of data lines
    num_data = number_of_data(csvfile)
    user_num = 0

    # Create the separate tex and then pdf files
    for row_line in dict_reader:

        custom_msg = str_msg        

        user_num += 1

        # The output filename
        outfilename = output_template
        for colheader in csv_headers:
            to_replace = '<' + colheader.strip() + '>'
            replaced_by = row_line[colheader]
            outfilename = re.sub(to_replace,replaced_by,outfilename)

        outfilename = outfilename.replace(".pdf","") 
        outtex = outfilename + '.tex'
        outpdf = outfilename + '.pdf'
        
        print("%s/%s: Creating %s ..." %(user_num,num_data,outpdf))    

        outtexfh = open(outtex,'w')

        # For each row replace the column headers
        for colheader in csv_headers:
            to_replace = '<' + colheader.strip() + '>'
            replaced_by = row_line[colheader]
            custom_msg = re.sub(to_replace,replaced_by,custom_msg)    

        for line in custom_msg.splitlines():
            line = line + '\n'
            outtexfh.write(line)

        outtexfh.close()

        time.sleep(1)

        # Run tex cmd
        cmd = ' '.join([latexmk, dflt_make_opts, ltxcompiler, outtex])
        proc = Popen(shlex.split(cmd), stdout=PIPE, stderr=PIPE)
        stdout, stderr = proc.communicate()
        rtn = proc.returncode
        
        time.sleep(1)

        # Exit if return code is not zero.
        if rtn != 0:
            print('Problem in running "%s" on "%s". Aborting.' 
                %(latexmk, outtex))
            sys.exit(1)

        # Move file
        src = outpdf
        dest = os.path.join(pdf_dir,outpdf)
        move(src,dest)

        # Clean files
        exts = (['.aux', '.log', '.out', '.tex', '.fls', '.synctex.gz',
                 '.fdb_latexmk', '.ps', '.dvi'])
        for ext in exts:
            delfile = outfilename + ext
            if exists(delfile):  
                os.remove(delfile)

