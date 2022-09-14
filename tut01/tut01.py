import csv

import math

octant = [0 for i in range(9)]

# opening octant_input file for reading purpose

data = list(csv.reader(open('octant_input.csv','r'))) 

csv.writer(open('octant_output.csv','w',newline='')).writerows(data)

# Calculating the numbers of coordinate of different octant

for i in range (1,len(data)):
    line = data[i] 
    if((float(line[7]) > 0) and (float(line[8]) > 0) and (float(line[9]) > 0)):
        octant[1]+=1
    if((float(line[7]) < 0) and (float(line[8]) > 0) and (float(line[9]) > 0)):
        octant[2]+=1
    if((float(line[7]) < 0) and (float(line[8]) < 0) and (float(line[9]) > 0)):
        octant[3]+=1
    if((float(line[7]) > 0) and (float(line[8]) < 0) and (float(line[9]) > 0)):
        octant[4]+=1
    if((float(line[7]) > 0) and (float(line[8]) > 0) and (float(line[9]) < 0)):
        octant[5]+=1
    if((float(line[7]) < 0) and (float(line[8]) > 0) and (float(line[9]) < 0)):
        octant[6]+=1
    if((float(line[7]) < 0) and (float(line[8]) < 0) and (float(line[9]) < 0)):
        octant[7]+=1
    if((float(line[7]) > 0) and (float(line[8]) < 0) and (float(line[9]) < 0)):
        octant[8]+=1
