import datetime as dt;
import xlrd


# Read and store excel file into variable excel_File
excel_File=xlrd.open_workbook("in/Site_Capacity.xlsx")

# Returns index of a specific value in the header row of the excel file
def getColumnIndex(header,value):
    for i in range(len(header)):

        if(type(header[i].value)==float):
            dateTime=xlrd.xldate_as_datetime(header[i].value,excel_File.datemode)
            if(dateTime==value):
                return i
        
        if(header[i].value==value):
            return i
    return -1
        
# Returns a datetime representation of the given parameters year and month
def getDate(year,month):
    dateTime=dt.datetime(year,month,1)
    return dateTime

# Function required to receive user input
# Returns three values:
#   -   value skill of type string
#   -   value year of type int
#   -   value month of type int
def getInput():

    skill=input("Enter Skill: ")
    year=getInt("Year")
    while(True):
        month=getInt("Month")
        if(controlInt(month)):
            break

    return skill,year,month

# Returns a user input that gets type control before 
# returning
# Type needs to be int 
def getInt(value):
    while(True):
        val=input("Enter "+value+": ")
        try:
            val=int(val)
        except:
            print("Please enter a numeric value.")
        if(isNumeric(val)):
             return val
   
# Returns a summation of all values from a skill at a specific date
# Reads data from excel file excel_File
def getSum(skill,year,month):
    summation=None
    sheet=excel_File.sheet_by_index(0)
    header=sheet.row(0)
    skillIndex=getColumnIndex(header,"Skill")
    dateTime =getDate(year,month)
    dateIndex=getColumnIndex(header,dateTime)
    
    if(dateIndex==-1):
       print("No values on ",dateTime," with skill ",skill)
       return

    for i in range(sheet.nrows):
        row=sheet.row(i)
        
        if(row[skillIndex].value.lower()==skill.lower()):
            if(isNumeric(row[dateIndex].value)):
                if(summation==None):
                    summation=row[dateIndex].value
                else:
                    summation+=row[dateIndex].value

    print("Sum of Values: ",summation)
    return summation

# Returns True if the give parameter is type of int or float
# Else returns False
def isNumeric(value):
    valueType=type(value)
    if(valueType==int or valueType==float ):
        return True
    return False

# Returns true if the value equals to
#   - y or yes or n or no
def matches(value):
    if(value=="y" or value=="n" or value=="yes" or value=="no"):
        return True
    return False

# Returns true if the value is a number from 1 to 12
# Else returns false
def controlInt(value):
    if(value<1 or value>12):
        print("Please enter a number from 1 to 12")
        return False
    return True

# Returns true if user input equals to n or no
# Returns false if user input equals to y or yes
# Else loops until one assignment has been made
def quit():
    answer=""
    while(not matches(answer)):
        answer=input("Get another summation? (y/n) ")
        if(answer.lower()=="n" or answer.lower()=="no"):
            return False
        elif(answer.lower()=="y" or answer.lower()=="yes"):
            return True
        print("Valid answers: y|yes|n|no")
        
    
# Main function 
def main():
    notQuit=True
    while(notQuit):
        skill,year,month=getInput()
        summation=getSum(skill,year,month)
        notQuit=quit()
    print("Program closed")
        

# Setting main function
if __name__ == "__main__":
    main()

