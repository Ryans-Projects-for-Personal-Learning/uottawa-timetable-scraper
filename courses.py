from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

import xlwt
from xlwt import Workbook
import xlsxwriter

# Navigate to the timetable page
url = "https://uocampus.public.uottawa.ca/psc/csprpr9pub/EMPLOYEE/HRMS/c/UO_SR_AA_MODS.UO_PUB_CLSSRCH.GBL?languageCd=ENG&PortalActualURL=https%3a%2f%2fuocampus.public.uottawa.ca%2fpsc%2fcsprpr9pub%2fEMPLOYEE%2fHRMS%2fc%2fUO_SR_AA_MODS.UO_PUB_CLSSRCH.GBL%3flanguageCd%3dENG&PortalContentURL=https%3a%2f%2fuocampus.public.uottawa.ca%2fpsc%2fcsprpr9pub%2fEMPLOYEE%2fHRMS%2fc%2fUO_SR_AA_MODS.UO_PUB_CLSSRCH.GBL&PortalContentProvider=HRMS&PortalCRefLabel=Public%20Class%20Search&PortalRegistryName=EMPLOYEE&PortalServletURI=https%3a%2f%2fuocampus.public.uottawa.ca%2fpsp%2fcsprpr9pub%2f&PortalURI=https%3a%2f%2fuocampus.public.uottawa.ca%2fpsc%2fcsprpr9pub%2f&PortalHostNode=HRMS&NoCrumbs=yes&PortalKeyStruct=yes"

def writeFile(courses,semesters,subjects):
    wb = Workbook()

    times = [
    "Mo 08:30 - 09:50","Mo 10:00 - 11:20","Mo 08:30 - 11:20","Mo 11:30 - 12:50","Mo 13:00 - 14:20","Mo 14:30 - 15:50","Mo 16:00 - 17:20","Mo 14:30 - 17:20","Mo 17:30 - 18:50","Mo 19:00 - 20:20","Mo 17:30 - 20:20","Mo 19:00 - 21:50",
    "Tu 08:30 - 09:50","Tu 10:00 - 11:20","Tu 08:30 - 11:20","Tu 11:30 - 12:50","Tu 13:00 - 14:20","Tu 14:30 - 15:50","Tu 16:00 - 17:20","Tu 14:30 - 17:20","Tu 17:30 - 18:50","Tu 19:00 - 20:20","Tu 17:30 - 20:20","Tu 19:00 - 21:50",
    "We 08:30 - 09:50","We 10:00 - 11:20","We 08:30 - 11:20","We 11:30 - 12:50","We 13:00 - 14:20","We 14:30 - 15:50","We 16:00 - 17:20","We 14:30 - 17:20","We 17:30 - 18:50","We 19:00 - 20:20","We 17:30 - 20:20","We 19:00 - 21:50",
    "Th 08:30 - 09:50","Th 10:00 - 11:20","Th 08:30 - 11:20","Th 11:30 - 12:50","Th 13:00 - 14:20","Th 14:30 - 15:50","Th 16:00 - 17:20","Th 14:30 - 17:20","Th 17:30 - 18:50","Th 19:00 - 20:20","Th 17:30 - 20:20","Th 19:00 - 21:50",
    "Fr 08:30 - 09:50","Fr 10:00 - 11:20","Fr 08:30 - 11:20","Fr 11:30 - 12:50","Fr 13:00 - 14:20","Fr 14:30 - 15:50","Fr 16:00 - 17:20","Fr 14:30 - 17:20","Fr 17:30 - 18:50","Fr 19:00 - 20:20","Fr 17:30 - 20:20","Fr 19:00 - 21:50"
    ]

    sheet1 = wb.add_sheet("Fall Courses")
    sheet2 = wb.add_sheet("Winter Courses")

    sheet3 = wb.add_sheet("Fall Schedule")
    sheet4 = wb.add_sheet("Winter Schedule")

    #Create headers for courses
    sheet1.write(0,0,"Course")
    sheet1.write(0,1,"Time slot 1")
    sheet1.write(0,2,"Time slot 2")
    sheet1.write(0,3,"Subject")
    sheet1.write(0,4,"Semester")

    sheet2.write(0,0,"Course")
    sheet2.write(0,1,"Time slot 1")
    sheet2.write(0,2,"Time slot 2")
    sheet2.write(0,3,"Subject")
    sheet2.write(0,4,"Semester")

    #Filling in rows for schedule
    s1row = 1 
    s2row = 1

    #Create list of courses
    for i in range(1,len(courses)):
        #Write to "Fall Courses" sheet
        if semesters[0] in courses[i]:
            sheet1.write(s1row,0,courses[i][0]) #Name
            sheet1.write(s1row,1,courses[i][1][0]) #Timeslot 1
            sheet1.write(s1row,3,courses[i][2]) #Subject
            sheet1.write(s1row,4,courses[i][3]) #Semester
            try:
                sheet1.write(s1row,2,courses[i][1][1])
                s1row += 1
            except:
                sheet1.write(s1row,2,"")
                s1row += 1
        #Write to "Winter Courses" sheet
        elif semesters[1] in courses[i]:
            sheet2.write(s2row,0,courses[i][0])
            sheet2.write(s2row,1,courses[i][1][0])
            sheet2.write(s2row,3,courses[i][2])
            sheet2.write(s2row,4,courses[i][3]) 
            try:
                sheet2.write(s2row,2,courses[i][1][1])
                s2row += 1
            except:
                sheet2.write(s2row,2," ")
                s2row += 1

    #Row variables
    s3col = 1
    s4col = 1

    #Add times
    for timeRow in range(0,len(times)):
        sheet3.write(timeRow+1,0,times[timeRow])
        sheet4.write(timeRow+1,0,times[timeRow])
    
    #Add subject columns
    for subject in range(0, len(subjects)):
        sheet3.write(0,s3col,subjects[subject])
        sheet4.write(0,s4col,subjects[subject])
        s3col += 1
        s4col += 1
    
    #Write countifs formulas
    row = 1
    col = 1 

    char = 'B'

    for col in range(1,len(subjects)+1):
        subjectCell = chr(ord(char)) + '1'
        for row in range(1,61):
            timeCell = 'A' + str(row+1)
            fallFormula = "=COUNTIFS('Fall Courses'!B2:B300," + timeCell + ",'Fall Courses'!D2:D300," + subjectCell + ") + COUNTIFS('Fall Courses'!C2:C300," + timeCell + ",'Fall Courses'!D2:D300," + subjectCell + ")"
            winterFormula = "=COUNTIFS('Winter Courses'!B2:B300," + timeCell + ",'Winter Courses'!D2:D300," + subjectCell + ") + COUNTIFS('Winter Courses'!C2:C300," + timeCell + ",'Winter Courses'!D2:D300," + subjectCell + ")"
            sheet3.write(row,col,fallFormula)
            sheet4.write(row,col,winterFormula)
        char = chr(ord(char)+1)

    print("Excel file saved as CourseList.xls")
    wb.save("CourseList.xls")

def main():
    semesterSelectId = "CLASS_SRCH_WRK2_STRM$35$"
    
    semesters = ["2019 Fall Term","2020 Winter Term"]
    subjects = ["ECO","DVM","POL","PAP"]
    
    # create a new Firefox session
    driver = webdriver.Firefox()
    driver.implicitly_wait(30)
    driver.maximize_window()

    driver.get(url)

    #Extract courses
    course = []
    courses = []
    times = []

    #iteration variables
    i = 0
    j = 0

    for semester in semesters:
        for subject in subjects:

            #Select term
            Semselect = Select(driver.find_element_by_id(semesterSelectId))
            Semselect.select_by_visible_text(semester)

            #Check off years
            driver.find_element_by_id('UO_PUB_SRCH_WRK_SSR_RPTCK_OPT_01$0').click()
            driver.find_element_by_id('UO_PUB_SRCH_WRK_SSR_RPTCK_OPT_02$0').click()
            driver.find_element_by_id('UO_PUB_SRCH_WRK_SSR_RPTCK_OPT_03$0').click()
            driver.find_element_by_id('UO_PUB_SRCH_WRK_SSR_RPTCK_OPT_04$0').click()

            subjectFieldId = "SSR_CLSRCH_WRK_SUBJECT$0"
            driver.find_element_by_id(subjectFieldId).send_keys(subject)
            driver.find_element_by_id('CLASS_SRCH_WRK2_SSR_PB_CLASS_SRCH').click() #to course results page
            while True:
                try:
                    driver.find_element_by_id("MTG_CLASS_NBR$" + str(i)).click()  #Click on each link to find course details
                    courseName = driver.find_element_by_id("DERIVED_CLSRCH_DESCR200").text #find course name
                    course.append(courseName) #Add course name to course
                    while True:
                        try:
                            time = driver.find_element_by_id("MTG_SCHED$" + str(j)).text #Find time
                            times.append(time) #add to the times list
                            j += 1
                        except:
                            course.append(times) #Add times to course
                            course.append(subject) #Add subject to course
                            course.append(semester) #Add semester to the course
                            courses.append(course) #Add course to overall list of courses
                            course = [] #empty course
                            times = [] #empty time
                            j = 0
                            driver.find_element_by_id("CLASS_SRCH_WRK2_SSR_PB_BACK").click()
                            break
                    i += 1
                    driver.implicitly_wait(2)
                except:
                    print("Course extraction for",subject,"for",semester,"complete with a new total of",len(courses),"courses.")
                    driver.get(url)
                    driver.implicitly_wait(3)
                    i = 0
                    break

    print("Course extraction completed")
    return writeFile(courses,semesters,subjects)

main()