# File Name: Load_Balance_Calculator_v4.py
# Author: Aditya Harakare
# Last Modified: May 3, 2020
# NOTE: File "Equipment_and_Power_Consumption.xlsx should be in same folder as this code"
# This code takes the input of the equipment power ratings from an excel sheet and
# outputs the 3-phase load balancing schema
# Output can be seen in "Load_Balance_Results.txt" file and also on the terminal

# import required libraries
from operator import itemgetter
import xlrd

# Specify single phase operating voltage
OPER_VOLT = 220

# initialize the variables and arrays
three_ph_load_power = 0
single_ph_array = []        # 2D array containing the name of equipment and power rating
three_ph_array = []          # 2D array containing the name of equipment and power rating
r_array = []
y_array = []
b_array = []
r_ph_power = []
y_ph_power = []
b_ph_power = []

# Read the data from the Excel file
print('################################## READING DATA ##################################')
loc = "Equipment_and_Power_Consumption.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(1, sheet.nrows):
    inp_name = sheet.cell_value(i, 0)
    phases_req = sheet.cell_value(i, 3)
    inp = 1000*sheet.cell_value(i, 4)     # Power required in Watts
    if (phases_req == 3):
        three_ph_array.append([inp_name, float(inp)])
        if (sheet.cell_value(i, 2)>1):    # If the quantity is > 1, create dummy variables and append them to the array
            for j in range (2, sheet.cell_value(i, 2)+1):
                three_ph_array.append([inp_name+"_"+str(j), float(inp)])
    else:
        single_ph_array.append([inp_name, float(inp)])
        if (sheet.cell_value(i, 2)>1):    # If the quantity is > 1, create dummy variables and append them to the array
            for j in range (2, int(sheet.cell_value(i, 2))+1):
                single_ph_array.append([inp_name+"_"+str(j), float(inp)])

# Sort the array of single phase equipments in descending order
single_ph_array = sorted(single_ph_array, key=itemgetter(1), reverse=True)

# Assign each equipment of single phase array either to R/Y/B phase depending on the existing loads
# The equipment with highest power requirement gets assigned first
# The phase with the minimum load gets assigned the next equipment in the sorted array
for i in range(0, len(single_ph_array)):
    sum_r = sum(r_ph_power)         # Maintain Counter for current phase load
    sum_y = sum(y_ph_power)         # Maintain Counter for current phase load
    sum_b = sum(b_ph_power)         # Maintain Counter for current phase load
    if sum_r < sum_y:
        if sum_r < sum_b:
            r_array.append(single_ph_array[i])
            r_ph_power.append(single_ph_array[i][1])
        else:
            b_array.append(single_ph_array[i])
            b_ph_power.append(single_ph_array[i][1])
    elif (sum_y < sum_b):
        y_array.append(single_ph_array[i])
        y_ph_power.append(single_ph_array[i][1])
    else:
        b_array.append(single_ph_array[i])
        b_ph_power.append(single_ph_array[i][1])

# Open the output file
op_file = open("Load_Balance_Results.txt", "w")
op_file.write("######################## Load Balance Results #################### \n")

# Print the results
print('################################## ANALYSIS ##################################')
print('Equipment on R - Phase:')
op_file.write('Equipment on R - Phase: \n')
for i in range(0, len(r_array)):
    print(r_array[i][0])
    op_file.write(r_array[i][0]+"\n")
for i in range(0, len(three_ph_array)):
    print(three_ph_array[i][0], "(3-phase)")
    op_file.write(three_ph_array[i][0] + "(3-phase) \n")
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on Y - Phase:')
op_file.write('Equipment on Y - Phase: \n')
for i in range(0, len(y_array)):
    print(y_array[i][0])
    op_file.write(y_array[i][0]+'\n')
for i in range(0, len(three_ph_array)):
    print(three_ph_array[i][0], "(3-phase)")
    op_file.write(three_ph_array[i][0] + "(3-phase) \n")
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on B - Phase:')
op_file.write('Equipment on B - Phase: \n')
for i in range(0, len(b_array)):
    print(b_array[i][0])
    op_file.write(b_array[i][0] + '\n')
for i in range(0, len(three_ph_array)):
    print(three_ph_array[i][0], "(3-phase)")
    op_file.write(three_ph_array[i][0] + "(3-phase) \n")

# Calculate and print the load on each phase
print('################################## PHASE LOAD ###############################')
op_file.write("############################# PHASE LOAD ######################### \n")
for i in range(0, len(three_ph_array)):
    three_ph_load_power = three_ph_load_power + three_ph_array[i][1]

sum_r_power = (sum(r_ph_power))+(three_ph_load_power/3)
sum_y_power = (sum(y_ph_power))+(three_ph_load_power/3)
sum_b_power = (sum(b_ph_power))+(three_ph_load_power/3)

print("Power drawn from R-Phase: ", sum_r_power, "Watts")
op_file.write("Power drawn from R-Phase: " + str(sum_r_power) + " Watts\n")
print("Power drawn from Y-Phase: ", sum_y_power, "Watts")
op_file.write("Power drawn from Y-Phase: " + str(sum_y_power) + " Watts\n")
print("Power drawn from B-Phase: ", sum_b_power, "Watts")
op_file.write("Power drawn from B-Phase: " + str(sum_b_power) + " Watts\n")
print("Current drawn from R-Phase: ", sum_r_power/OPER_VOLT, "A")
op_file.write("Current drawn from R-Phase: " + str(sum_r_power/OPER_VOLT) + " A\n")
print("Current drawn from Y-Phase: ", sum_y_power/OPER_VOLT, "A")
op_file.write("Current drawn from Y-Phase: " + str(sum_y_power/OPER_VOLT) + " A\n")
print("Current drawn from B-Phase: ", sum_b_power/OPER_VOLT, "A")
op_file.write("Current drawn from B-Phase: " + str(sum_b_power/OPER_VOLT) + " A\n")
print('##################################    END    ###############################')
op_file.write("#############################    END    ##########################")
op_file.close()
