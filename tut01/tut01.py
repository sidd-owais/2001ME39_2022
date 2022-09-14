import csv

import math

octant = [0 for i in range(9)]

# opening octant_input file for reading purpose

data = list(csv.reader(open('octant_input.csv','r'))) 

csv.writer(open('octant_output.csv','w',newline='')).writerows(data)

