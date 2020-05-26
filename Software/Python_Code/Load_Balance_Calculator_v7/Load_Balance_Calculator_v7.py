# import required libraries
from operator import itemgetter
import xlrd
import cmath
import math
OPER_VOL = 220
TRY_NO = 5
# initialize the arrays
input_array = []
real_array = []
img_array = []
complex_array = []
three_ph_array = []
single_ph_array = []
r_array = []
y_array = []
b_array = []
r_full_array = []
y_full_array = []
b_full_array = []

# Read the data from the Excel file
print('################################## READING DATA ##################################')
loc = "Equipment_and_Power_Consumption.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

for i in range(1, sheet.nrows):
    inp_name = sheet.cell_value(i, 0)
    phases_req = sheet.cell_value(i, 3)
    power = 1000*sheet.cell_value(i, 4)
    power_factor = sheet.cell_value(i, 5)
    mod_Z = (OPER_VOL*OPER_VOL)/power
    Z_complex = complex(mod_Z*power_factor,math.sqrt(pow(mod_Z,2)-pow(mod_Z*power_factor,2)))
    if (sheet.cell_value(i,6)=="Leading"):
        Z_complex = complex(Z_complex.real,Z_complex.imag*(-1))
    if (phases_req == 3):
        three_ph_array.append([inp_name, 3/Z_complex])
        if (sheet.cell_value(i, 2) > 1):  # If the quantity is > 1, create dummy variables and append them to the array
            for j in range(2, int(sheet.cell_value(i, 2) + 1)):
                three_ph_array.append([inp_name + "_" + str(j), 1/Z_complex])
    else:
        single_ph_array.append([inp_name, 1/Z_complex])
        if (sheet.cell_value(i, 2) > 1):  # If the quantity is > 1, create dummy variables and append them to the array
            for j in range(2, int(sheet.cell_value(i, 2)) + 1):
                single_ph_array.append([inp_name + "_" + str(j), 1/Z_complex])

# We know that the sum of all admittance is constant and admittance on each phase should tend to avg_value = (this sum)/3
# error_in_x is the difference in the avg_value and current atmittance on x phase.
# The algorithm allocates each impedance to the 3 phases one by one and checks which error is reduced the most.
# That phase is finally allocated the impedance.

single_ph_array = sorted(single_ph_array, key=lambda x: abs(x[1]), reverse=True)
average_value = 0
for i in range(0, len(single_ph_array)):
    average_value = average_value + single_ph_array[i][1]
average_value = average_value/3

for i in range(0, len(single_ph_array)):
    error_in_r = abs(sum(r_array) - average_value)
    error_in_y = abs(sum(y_array) - average_value)
    error_in_b = abs(sum(b_array) - average_value)
    error_reduction_r = error_in_r - abs(sum(r_array) + single_ph_array[i][1] - average_value)
    error_reduction_y = error_in_y - abs(sum(y_array) + single_ph_array[i][1] - average_value)
    error_reduction_b = error_in_b - abs(sum(b_array) + single_ph_array[i][1] - average_value)
    if (error_reduction_r >= error_reduction_b) and (error_in_r >= error_in_y):
        r_array.append(single_ph_array[i][1])
        r_full_array.append(single_ph_array[i])
    elif (error_reduction_y >= error_reduction_b) and (error_in_y >= error_in_r):
        y_array.append(single_ph_array[i][1])
        y_full_array.append(single_ph_array[i])
    else:
        b_array.append(single_ph_array[i][1])
        b_full_array.append(single_ph_array[i])


score1 = abs(sum(r_array) - sum(y_array)) + abs(sum(y_array) - sum(b_array)) + abs(sum(r_array) - sum(b_array))

# After first allotment of elements we again run through all elements one by one.
# We temporarily relocate an element in the other two phases one by one and check the score.
# The lesser the score the better the balance.
# If the new_score is lesser than the previous we shift the element to the other phase.
for j in range(0, TRY_NO):
    for i in range(0, len(r_array)):
        if i >= len(r_array):
            break
        # r to y
        score2 = abs(sum(r_array) - 2 * r_array[i] - sum(y_array)) + abs(sum(y_array) + r_array[i] - sum(b_array)) + abs(
            sum(r_array) - r_array[i] - sum(b_array))
        if (score2 < score1):
            y_array.append(r_array[i])
            r_array.remove(r_array[i])
            y_full_array.append(r_full_array[i])
            r_full_array.remove(r_full_array[i])
            score1 = score2
        # r to b
        score2 = abs(sum(r_array) - r_array[i] - sum(y_array)) + abs(sum(y_array) - r_array[i] - sum(b_array)) + abs(
            sum(r_array) - 2 * r_array[i] - sum(b_array))
        if (score2 < score1):
            b_array.append(r_array[i])
            r_array.remove(r_array[i])
            b_full_array.append(r_full_array[i])
            r_full_array.remove(r_full_array[i])
            score1 = score2

    for i in range(0, len(y_array)):
        # y to r
        if i >= len(y_array):
            break
        score2 = abs(sum(r_array) + 2 * y_array[i] - sum(y_array)) + abs(sum(y_array) - y_array[i] - sum(b_array)) + abs(
            sum(r_array) + y_array[i] - sum(b_array))
        if (score2 < score1):
            r_array.append(y_array[i])
            y_array.remove(y_array[i])
            r_full_array.append(y_full_array[i])
            y_full_array.remove(y_full_array[i])
            score1 = score2
        # y to b
        score2 = abs(sum(r_array) - y_array[i] - sum(y_array)) + abs(sum(y_array) - 2 * y_array[i] - sum(b_array)) + abs(
            sum(r_array) - y_array[i] - sum(b_array))
        if (score2 < score1):
            b_array.append(y_array[i])
            y_array.remove(y_array[i])
            b_full_array.append(y_full_array[i])
            y_full_array.remove(y_full_array[i])
            score1 = score2

    for i in range(0, len(b_array)):
        if i >= len(b_array):
            break
        # b to r
        score2 = abs(sum(r_array) + b_array[i] - sum(y_array)) + abs(sum(y_array) - b_array[i] - sum(b_array)) + abs(
            sum(r_array) + 2 * b_array[i] - sum(b_array))
        if (score2 < score1):
            r_array.append(b_array[i])
            b_array.remove(b_array[i])
            r_full_array.append(b_full_array[i])
            b_full_array.remove(b_full_array[i])
            score1 = score2
        # b to y
        score2 = abs(sum(r_array) - b_array[i] - sum(y_array)) + abs(sum(y_array) + 2 * b_array[i] - sum(b_array)) + abs(
            sum(r_array) - sum(b_array) - b_array[i])
        if (score2 < score1):
            y_array.append(b_array[i])
            b_array.remove(b_array[i])
            y_full_array.append(b_full_array[i])
            b_full_array.remove(b_full_array[i])
            score1 = score2

for i in range (0, len(three_ph_array)):
    r_full_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] ])
    y_full_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] ])
    b_full_array.append([three_ph_array[i][0] + " (3-ph)", three_ph_array[i][1] ])
    r_array.append(three_ph_array[i][1])
    y_array.append(three_ph_array[i][1])
    b_array.append(three_ph_array[i][1])

# Open the output file
op_file = open("Load_Balance_Results.txt", "w",encoding='utf-8')
op_file.write("######################## Load Balance Results #################### \n")
# Print the results
print('################################## ANALYSIS ##################################')
print('Equipment on R - Phase:')
op_file.write('Equipment on R - Phase: \n')
for i in range(0, len(r_full_array)):
    print(r_full_array[i][0])
    op_file.write(r_full_array[i][0] + "\n")
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on Y - Phase:')
op_file.write('Equipment on Y - Phase: \n')
for i in range(0, len(y_full_array)):
    print(y_full_array[i][0])
    op_file.write(y_full_array[i][0] + '\n')
print("-----------------------")
op_file.write('----------------------- \n')
print('Equipment on B - Phase:')
op_file.write('Equipment on B - Phase: \n')
for i in range(0, len(b_full_array)):
    print(b_full_array[i][0])
    op_file.write(b_full_array[i][0] + '\n')



# Calculate and print the phase currents (Magnitude + Angle)
# For perfectly balanced load the magnitude of currents should be equal and the angles should be 120 apart
print('################################ PHASE CURRENTS #############################')
op_file.write("########################### PHASE CURRENTS ####################### \n")
op_file.write("(For perfectly balanced load the magnitude of currents should be equal and the angles should be 120 apart) \n")
print("Current drawn from R-Phase: ", round(abs(OPER_VOL*sum(r_array)), 2), "∠",
          round(180 * cmath.phase(OPER_VOL*sum(r_array)) / 3.142, 2), "A")
op_file.write("Current drawn from R-Phase: " + str(round(abs(OPER_VOL *sum(r_array)), 2)) + "∠" + str(
        round(180 * cmath.phase(OPER_VOL *sum(r_array)) / 3.142, 2)) + " A\n")
print("Current drawn from Y-Phase: ", round(abs(OPER_VOL*sum(y_array)), 2), "∠",
      (120 + round(180 * cmath.phase(OPER_VOL*sum(y_array)) / 3.142, 2)) % 360, "A")
op_file.write("Current drawn from Y-Phase: " + str(round(abs(OPER_VOL*sum(y_array)), 2)) + "∠" + str(
    (120 + round(180 * cmath.phase(OPER_VOL *sum(y_array)) / 3.142, 2)) % 360) + " A\n")
print("Current drawn from B-Phase: ", round(abs(OPER_VOL *sum(b_array)), 2), "∠",
      (240 + round(180 * cmath.phase(OPER_VOL *sum(b_array)) / 3.142, 2)) % 360, "A")
op_file.write("Current drawn from B-Phase: " + str(round(abs(OPER_VOL *sum(b_array)), 2)) + "∠" + str(
    (240 + round(180 * cmath.phase(OPER_VOL *sum(b_array)) / 3.142, 2)) % 360) + " A\n")
print('##################################    END    ###############################')
op_file.write("#############################    END    ##########################")
op_file.close()
