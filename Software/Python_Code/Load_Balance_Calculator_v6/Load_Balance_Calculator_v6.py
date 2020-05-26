# File Name: Load_Balance_Calculator_v6.py
# Author: Aditya Harakare
# Last Modified: May 6, 2020
# NOTE: File "Equipment_and_Power_Consumption.xlsx should be in same folder as this code"
# This code takes the input of the equipment power ratings and equipment type (Resistive/Inductive) from an excel sheet
# and outputs the 3-phase load balancing schema
# Output can be seen in "Load_Balance_Results.txt" file and also on the terminal

# import required libraries
from operator import itemgetter
import xlrd
import cmath

OPER_VOL = 220

# initialize the variables and arrays
single_ph_array = []  # 3D array containing the name of equipment, power rating and type
three_ph_array = []  # 3D array containing the name of equipment, power rating and type
single_ph_imp_array = []
single_ph_res_array = []
r_array = []
y_array = []
b_array = []
r_imp_array = []
y_imp_array = []
b_imp_array = []
r_ph_power = []
y_ph_power = []
b_ph_power = []
r_ph_ind_power = []
y_ph_ind_power = []
b_ph_ind_power = []

# Read the data from the Excel file
print('################################## READING DATA ##################################')
loc = "Equipment_and_Power_Consumption.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(1, sheet.nrows):
    inp_name = sheet.cell_value(i, 0)
    phases_req = sheet.cell_value(i, 3)
    load_type = sheet.cell_value(i, 5)
    inp = 1000 * sheet.cell_value(i, 4)  # Power required in VA
    if (phases_req == 3):
        three_ph_array.append([inp_name, float(inp), load_type])
        if (sheet.cell_value(i, 2) > 1):  # If the quantity is > 1, create dummy variables and append them to the array
            for j in range(2, int(sheet.cell_value(i, 2) + 1)):
                three_ph_array.append([inp_name + "_" + str(j), float(inp), load_type])
    else:
        single_ph_array.append([inp_name, float(inp), load_type])
        if (sheet.cell_value(i, 2) > 1):  # If the quantity is > 1, create dummy variables and append them to the array
            for j in range(2, int(sheet.cell_value(i, 2)) + 1):
                single_ph_array.append([inp_name + "_" + str(j), float(inp), load_type])

# Divide Single Phase equipments into resistive and inductive
for i in range(0, len(single_ph_array)):
    if (single_ph_array[i][2] == "Resistive"):
        single_ph_res_array.append([single_ph_array[i][0], single_ph_array[i][1]])
    elif (single_ph_array[i][2] == "Inductive"):
        single_ph_imp_array.append([single_ph_array[i][0], single_ph_array[i][1]])
    else:
        print("Invalid Input:", single_ph_array[i])

# Sort the array of single phase resistive equipments in descending order
single_ph_res_array = sorted(single_ph_res_array, key=itemgetter(1), reverse=True)
# Sort the array of single phase inductive equipments in descending order
single_ph_imp_array = sorted(single_ph_imp_array, key=itemgetter(1), reverse=True)

# Assign each equipment of single phase resistive array either to R/Y/B phase depending on the existing loads
# The equipment with highest power requirement gets assigned first
# The phase with the minimum load gets assigned the next equipment in the sorted array
for i in range(0, len(single_ph_res_array)):
    sum_r = sum(r_ph_power)  # Maintain Counter for current phase load
    sum_y = sum(y_ph_power)  # Maintain Counter for current phase load
    sum_b = sum(b_ph_power)  # Maintain Counter for current phase load
    if sum_r < sum_y:
        if sum_r < sum_b:
            r_array.append(single_ph_res_array[i])
            r_ph_power.append(single_ph_res_array[i][1])
        else:
            b_array.append(single_ph_res_array[i])
            b_ph_power.append(single_ph_res_array[i][1])
    elif (sum_y < sum_b):
        y_array.append(single_ph_res_array[i])
        y_ph_power.append(single_ph_res_array[i][1])
    else:
        b_array.append(single_ph_res_array[i])
        b_ph_power.append(single_ph_res_array[i][1])

# Similarly for single phase reactive elements
for i in range(0, len(single_ph_imp_array)):
    sum_r = sum(r_ph_ind_power)  # Maintain Counter for current phase load
    sum_y = sum(y_ph_ind_power)  # Maintain Counter for current phase load
    sum_b = sum(b_ph_ind_power)  # Maintain Counter for current phase load
    if sum_r < sum_y:
        if sum_r < sum_b:
            r_array.append(single_ph_imp_array[i])
            r_ph_ind_power.append(single_ph_imp_array[i][1])
        else:
            b_array.append(single_ph_imp_array[i])
            b_ph_ind_power.append(single_ph_imp_array[i][1])
    elif (sum_y < sum_b):
        y_array.append(single_ph_imp_array[i])
        y_ph_ind_power.append(single_ph_imp_array[i][1])
    else:
        b_array.append(single_ph_imp_array[i])
        b_ph_ind_power.append(single_ph_imp_array[i][1])

# Assign the 3 phase equipments to all the phases
for i in range(0, len(three_ph_array)):
    r_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] / 3])
    y_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] / 3])
    b_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] / 3])
    if three_ph_array[i][2] == "Resistive":
        r_ph_power.append(three_ph_array[i][1])
        y_ph_power.append(three_ph_array[i][1])
        b_ph_power.append(three_ph_array[i][1])
    if three_ph_array[i][2] == "Inductive":
        r_ph_ind_power.append(three_ph_array[i][1])
        y_ph_ind_power.append(three_ph_array[i][1])
        b_ph_ind_power.append(three_ph_array[i][1])

# Open the output file
op_file = open("Load_Balance_Results.txt", "w", encoding='utf-8')
op_file.write("######################## Load Balance Results #################### \n")
# Print the results
print('################################## ANALYSIS ##################################')
print('Equipment on R - Phase:')
op_file.write('Equipment on R - Phase: \n')
for i in range(0, len(r_array)):
    print(r_array[i][0])
    op_file.write(r_array[i][0] + "\n")
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on Y - Phase:')
op_file.write('Equipment on Y - Phase: \n')
for i in range(0, len(y_array)):
    print(y_array[i][0])
    op_file.write(y_array[i][0] + '\n')
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on B - Phase:')
op_file.write('Equipment on B - Phase: \n')
for i in range(0, len(b_array)):
    print(b_array[i][0])
    op_file.write(b_array[i][0] + '\n')

# Calculate total power on all phases
r_power = complex(sum(r_ph_power), sum(r_ph_ind_power))
y_power = complex(sum(y_ph_power), sum((y_ph_ind_power)))
b_power = complex(sum(b_ph_power), sum(b_ph_ind_power))

# Calculate the load on each phase
load_on_r = 0
load_on_y = 0
load_on_b = 0
if r_power != 0:
    load_on_r = OPER_VOL * OPER_VOL / r_power
if b_power != 0:
    load_on_b = OPER_VOL * OPER_VOL / b_power
if y_power != 0:
    load_on_y = OPER_VOL * OPER_VOL / y_power

# Calculate and print the phase currents (Magnitude + Angle)
# For perfectly balanced load the magnitude of currents should be equal and the angles should be 120 apart
print('################################ PHASE CURRENTS #############################')
op_file.write("########################### PHASE CURRENTS ####################### \n")
op_file.write(
    "(For perfectly balanced load the magnitude of currents should be equal and the angles should be 120 apart) \n")
if load_on_r != 0:
    print("Current drawn from R-Phase: ", round(abs(OPER_VOL / load_on_r), 2), "∠",
          round(180 * cmath.phase(OPER_VOL / load_on_r) / 3.142, 2), "A")
    op_file.write("Current drawn from R-Phase: " + str(round(abs(OPER_VOL / load_on_r), 2)) + "∠" + str(
        round(180 * cmath.phase(OPER_VOL / load_on_r) / 3.142, 2)) + " A\n")
if load_on_y != 0:
    print("Current drawn from Y-Phase: ", round(abs(OPER_VOL / load_on_y), 2), "∠",
          (120 + round(180 * cmath.phase(OPER_VOL / load_on_y) / 3.142, 2)) % 360, "A")
    op_file.write("Current drawn from Y-Phase: " + str(round(abs(OPER_VOL / load_on_y), 2)) + "∠" + str(
        (120 + round(180 * cmath.phase(OPER_VOL / load_on_y) / 3.142, 2)) % 360) + " A\n")
if load_on_b != 0:
    print("Current drawn from B-Phase: ", round(abs(OPER_VOL / load_on_b), 2), "∠",
          (240 + round(180 * cmath.phase(OPER_VOL / load_on_b) / 3.142, 2)) % 360, "A")
    op_file.write("Current drawn from B-Phase: " + str(round(abs(OPER_VOL / load_on_b), 2)) + "∠" + str(
        (240 + round(180 * cmath.phase(OPER_VOL / load_on_b) / 3.142, 2)) % 360) + " A\n")
print('##################################    END    ###############################')
op_file.write("#############################    END    ##########################")
op_file.close()
