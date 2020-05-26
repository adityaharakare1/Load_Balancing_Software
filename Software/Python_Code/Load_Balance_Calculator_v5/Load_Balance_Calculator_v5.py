# File Name: Load_Balance_Calculator_v5.py
# Author: Aditya Harakare
# Last Modified: May 4, 2020
# NOTE: File "Equipment_and_Power_Consumption.xlsx should be in same folder as this code"
# This code takes the input of the equipment power ratings and equipment type from an excel sheet and
# outputs the 3-phase load balancing schema
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
three_ph_imp_array = []
single_ph_res_array = []
r_array = []
y_array = []
b_array = []
r_imp_array = []
y_imp_array = []
b_imp_array = []
r_fixed_imp_array = []
y_fixed_imp_array = []
b_fixed_imp_array = []
r_ph_power = []
y_ph_power = []
b_ph_power = []
ans_r_array = []
ans_y_array = []
ans_b_array = []
min_cost = 999999999

# This function finds and returns the base 3 equivalent of the decimal number
def find_ternary(num):
    quotient = num / 3
    remainder = num % 3
    if quotient == 0:
        return ""
    else:
        return find_ternary(int(quotient)) + str(int(remainder))

# This functions appends 0's at the start of number x to make it a n_loads digit number
def append_zeros(x, n_loads):
    if (len(x) < n_loads):
        append_string = ""
        for i in range(0, n_loads - len(x)):
            append_string = append_string + "0"
        return append_string + x
    else:
        return x

# This function returns the cost given inputs z1, z2 and z3
# The cost is |z1-z2|+|z2-z3|
# The lesser the cost the closer z1, z2 and z3 (as z1+z2+z3 is constant)
def check_cost(r_imp_array, y_imp_array, b_imp_array):
    r_imp = 0
    y_imp = 0
    b_imp = 0
    for i in range(0, len(r_imp_array)):
        r_imp = r_imp + 1/(r_imp_array[i][1])
    for i in range(0, len(y_imp_array)):
        y_imp = y_imp + 1/(y_imp_array[i][1])
    for i in range(0, len(b_imp_array)):
        b_imp = b_imp + 1/(b_imp_array[i][1])
    if r_imp !=0:
        r_imp = 1/r_imp
    if y_imp !=0:
        y_imp = 1/y_imp
    if b_imp !=0:
        b_imp = 1/b_imp
    return abs(r_imp - y_imp) + abs(y_imp - b_imp)


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

# Add the single phase non-resistive loads to single_ph_imp_array
for i in range(0, len(single_ph_array)):
    if (single_ph_array[i][2] == "Capacitive"):
        impedance = complex(0, -1 * OPER_VOL * OPER_VOL / single_ph_array[i][1])
        single_ph_imp_array.append([single_ph_array[i][0], impedance])
    elif (single_ph_array[i][2] == "Inductive"):
        impedance = complex(0, OPER_VOL * OPER_VOL / single_ph_array[i][1])
        single_ph_imp_array.append([single_ph_array[i][0], impedance])

# Add all the 3 phase loads to the the three_ph_imp_array
for i in range(0, len(three_ph_array)):
    if (three_ph_array[i][2] == "Resistive"):
        impedance = complex(3 * OPER_VOL * OPER_VOL / three_ph_array[i][1], 0)
    elif (three_ph_array[i][2] == "Capacitive"):
        impedance = complex(0, -3 * OPER_VOL * OPER_VOL / three_ph_array[i][1])
    elif (three_ph_array[i][2] == "Inductive"):
        impedance = complex(0, 3 * OPER_VOL * OPER_VOL / three_ph_array[i][1])
    else:
        print("Error on equipment type of: ", three_ph_array[i][0])
        impedance = 0
    three_ph_imp_array.append([three_ph_array[i][0], impedance])

# Add the single phase resistive loads to the single_ph_res_array
for i in range(0, len(single_ph_array)):
    if (single_ph_array[i][2] == "Resistive"):
        single_ph_res_array.append([single_ph_array[i][0], single_ph_array[i][1]])

# Divide the single phase resistive loads among the 3 phases
# This division algorithmn ensures that the load is divided equally to the extent possible

# Sort the array of single phase resistive equipments in descending order
single_ph_res_array = sorted(single_ph_res_array, key=itemgetter(1), reverse=True)

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

# Maintain the equipment list on r-phase along with their impedance in r_imp_array
# Similarly for other phases
for i in range(0, len(r_array)):
    r_imp_array.append([r_array[i][0], complex(OPER_VOL * OPER_VOL / r_array[i][1], 0)])
for i in range(0, len(y_array)):
    y_imp_array.append([y_array[i][0], complex(OPER_VOL * OPER_VOL / y_array[i][1], 0)])
for i in range(0, len(b_array)):
    b_imp_array.append([b_array[i][0], complex(OPER_VOL * OPER_VOL / b_array[i][1], 0)])

# Add the three phase impedance to all the phases
for i in range(0, len(three_ph_imp_array)):
    r_imp_array.append([three_ph_imp_array[i][0] + " (3-ph)", three_ph_imp_array[i][1]])
    y_imp_array.append([three_ph_imp_array[i][0] + " (3-ph)", three_ph_imp_array[i][1]])
    b_imp_array.append([three_ph_imp_array[i][0] + " (3-ph)", three_ph_imp_array[i][1]])

# This comprises of our fixed equipments on the three phases
r_fixed_imp_array = r_imp_array.copy()
y_fixed_imp_array = y_imp_array.copy()
b_fixed_imp_array = b_imp_array.copy()

# Now we will try all the permutations of the non resistive elements among all the phases
# The permutation for which the cost is minimum i.e. most balanced among others is noted in the ans_x_array(s)
for i in range(0, pow(3, len(single_ph_imp_array))):
    x = find_ternary(i)
    x = append_zeros(x, len(single_ph_imp_array))
    for j in range(0, len(single_ph_imp_array)):
        if x[j] == "0":
            r_imp_array.append([single_ph_imp_array[j][0], single_ph_imp_array[j][1]])
            # print("Appened in r")
        elif (x[j] == "1"):
            y_imp_array.append([single_ph_imp_array[j][0], single_ph_imp_array[j][1]])
            # print("Appened in y")
        else:
            b_imp_array.append([single_ph_imp_array[j][0], single_ph_imp_array[j][1]])
            # print("Appened in b")
    cost = check_cost(r_imp_array, y_imp_array, b_imp_array)
    if cost < min_cost:
        min_cost = cost
        ans_r_array.clear()
        ans_b_array.clear()
        ans_y_array.clear()
        ans_r_array = r_imp_array.copy()
        ans_y_array = y_imp_array.copy()
        ans_b_array = b_imp_array.copy()
    r_imp_array.clear()
    y_imp_array.clear()
    b_imp_array.clear()
    r_imp_array = r_fixed_imp_array.copy()
    y_imp_array = y_fixed_imp_array.copy()
    b_imp_array = b_fixed_imp_array.copy()

# Open the output file
op_file = open("Load_Balance_Results.txt", "w",encoding='utf-8')
op_file.write("######################## Load Balance Results #################### \n")
# Print the results
print('################################## ANALYSIS ##################################')
print('Equipment on R - Phase:')
op_file.write('Equipment on R - Phase: \n')
for i in range(0, len(ans_r_array)):
    print(ans_r_array[i][0])
    op_file.write(ans_r_array[i][0]+"\n")
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on Y - Phase:')
op_file.write('Equipment on Y - Phase: \n')
for i in range(0, len(ans_y_array)):
    print(ans_y_array[i][0])
    op_file.write(ans_y_array[i][0]+'\n')
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on B - Phase:')
op_file.write('Equipment on B - Phase: \n')
for i in range(0, len(ans_b_array)):
    print(ans_b_array[i][0])
    op_file.write(ans_b_array[i][0] + '\n')

# Calculate the load on each phase
load_on_r= 0
load_on_y = 0
load_on_b = 0

for i in range(0, len(ans_r_array)):
    load_on_r = load_on_r + 1/ans_r_array[i][1]
for i in range(0, len(ans_y_array)):
    load_on_y = load_on_y + 1/ans_y_array[i][1]
for i in range(0, len(ans_b_array)):
    load_on_b = load_on_b + 1/ans_b_array[i][1]

if load_on_r != 0:
    load_on_r = 1/load_on_r
if load_on_b != 0:
    load_on_b = 1/load_on_b
if load_on_y != 0:
    load_on_y = 1/load_on_y

# Calculate and print the phase currents (Magnitude + Angle)
# For perfectly balanced load the magnitude of currents should be equal and the angles should be 0, 120 and 240
print('################################ PHASE CURRENTS #############################')
op_file.write("########################### PHASE CURRENTS ####################### \n")
op_file.write("(For perfectly balanced load the magnitude of currents should be equal and the angles should be 0, 120 and 240) \n")
if load_on_r !=0:
    print("Current drawn from R-Phase: ", round(abs(OPER_VOL/load_on_r),2), "∠", round(180*cmath.phase(OPER_VOL/load_on_r)/3.142,2), "A")
    op_file.write("Current drawn from R-Phase: " + str(round(abs(OPER_VOL/load_on_r),2)) +"∠" + str(round(180*cmath.phase(OPER_VOL/load_on_r)/3.142,2)) +" A\n")
if load_on_y !=0:
    print("Current drawn from Y-Phase: ", round(abs(OPER_VOL/load_on_y),2), "∠", (120+round(180*cmath.phase(OPER_VOL/load_on_y)/3.142,2))%360, "A")
    op_file.write("Current drawn from Y-Phase: " + str(round(abs(OPER_VOL/load_on_y),2)) +"∠" + str((120+round(180*cmath.phase(OPER_VOL/load_on_y)/3.142,2))%360) +" A\n")
if load_on_b !=0:
    print("Current drawn from B-Phase: ", round(abs(OPER_VOL/load_on_b),2), "∠", (240+round(180*cmath.phase(OPER_VOL/load_on_b)/3.142,2))%360, "A")
    op_file.write("Current drawn from B-Phase: " + str(round(abs(OPER_VOL/load_on_b),2)) +"∠" + str((240+round(180*cmath.phase(OPER_VOL/load_on_b)/3.142,2))%360) +" A\n")
print('##################################    END    ###############################')
op_file.write("#############################    END    ##########################")
op_file.close()
