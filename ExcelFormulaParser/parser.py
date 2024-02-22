#!/usr/bin/env python
# this file is part of this GitHub project : https://github.com/ronan-deshays/excel-formula-parser

# TO DO : multiline comments + sub-functions copy ability
# LIMITATIONS : 
#   variable name containing space not supported
#   unable to make difference between formula begin 
#   and "=" sign in formula body

import re
target = open("samples\excel_formula_out.txt", "w")
source = open("samples\excel_formula_in.txt", "r")

t = source.read()
t = re.sub("# [^\n]*","",t) # remove comments
t = t.replace("\n","") # remove line breaks
t = t.replace("=","\n\n=") # add some space between formulas
t = t.replace(" ","") # remove spaces

target.write(t)

target.close()
source.close()