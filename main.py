from openpyxl import load_workbook


def hammingDistance(string1, string2):
    # Start with a distance of zero, and count up
    string1 = string1.zfill(8)
    string2 = string2.zfill(8)
    distance = 0
    # Loop over the indices of the string
    L = len(string1)
    for i in range(L):
        # Add 1 to the distan   ce if these two characters are not equal
        if string1[i] != string2[i]:
            distance += 1
    # Return the final count of differences
    return distance

wb = load_workbook(filename='1.xlsx', data_only=True)
sheet = wb['1']

x_red_1 = str(sheet['W6'].value)
x_red_2 = str(sheet['W7'].value)
x_red_3 = str(sheet['W8'].value)
x_green_1 = str(sheet['X6'].value)
x_green_2 = str(sheet['X7'].value)
x_green_3 = str(sheet['X8'].value)
x_blue_1 = str(sheet['Y6'].value)
x_blue_2 = str(sheet['Y7'].value)
x_blue_3 = str(sheet['Y8'].value)

for j in range(2, 32):
    print('cell' + str(j))
    y_red = str(sheet['F' + str(j)].value)
    y_green = str(sheet['G' + str(j)].value)
    y_blue = str(sheet['H' + str(j)].value)
    sheet['I' + str(j)] = hammingDistance(x_red_1, y_red)
    sheet['J' + str(j)] = hammingDistance(x_green_1, y_green)
    sheet['K' + str(j)] = hammingDistance(x_blue_1, y_blue)
    sheet['L' + str(j)] = hammingDistance(x_red_2, y_red)
    sheet['M' + str(j)] = hammingDistance(x_green_2, y_green)
    sheet['N' + str(j)] = hammingDistance(x_blue_2, y_blue)
    sheet['O' + str(j)] = hammingDistance(x_red_3, y_red)
    sheet['P' + str(j)] = hammingDistance(x_green_3, y_green)
    sheet['Q' + str(j)] = hammingDistance(x_blue_3, y_blue)

wb.save(filename='1.xlsx')

#  pyinstaller -F main.py
