import openpyxl
from openpyxl.styles import Alignment, PatternFill

def main():
    # load the excel file
    file_path = 'test.xlsx'
    workbook = openpyxl.load_workbook(file_path)

    # select a specific worksheet
    worksheet = workbook['Sheet1']

    time = "2023-08-16 08:19:51 "
    print(time[8:10])
    print(time[11:13])

    letters = ["C2", "E2", "G2", "I2", "K2", "M2", "O2", "Q2", "S2", "U2", "W2", "Y2", "AA2", "AC2",
               "AE2", "AG2", "AI2", "AK2", "AM2", "AO2", "AQ2", "AS2", "AU2", "AW2", "AY2", "BA2",
               "BC2", "BE2", "BG2", "BI2", "BK2"]
    # for letter in letters:
    #     #print(type(letter))
    #     print(worksheet[letters[2]].value)
    #     if 22 == worksheet[letters[2]].value:
    #         print("chelseo")
    # print("\nskip\n")

    num = 0
    count = 0
    number = 20
    print(letters[num])
    for a in range(len(letters)):
        if str(number) == worksheet[letters[num]].value:
            print(number)
            number += 1
        else:
            print("not")
            
        if number > 31:
            number = 1
        num += 1
        


    print(count)

    print(len(letters))
    
    # num += 1
    # print(worksheet[letters[num]].value)
        #print(str(worksheet[letters].value))

if __name__ == "__main__":
    main()