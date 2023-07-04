from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl

driver = webdriver.Chrome()

data = []

pages = int(input("Enter the number of pages there are: "))

for x in range(0, pages):

    url = f'https://music.calendar.utoronto.ca/search-courses?page={x}'

    driver.get(url)

    soup = BeautifulSoup(driver.page_source, 'html.parser')

    courses = soup.find_all("div", class_="views-row")

    for course in courses:
        dct = {}
        course_name = course.find("h3", class_="js-views-accordion-group-header")
        course_desc = course.find("div", class_="views-field views-field-body")
        course_dst = course.find("span", class_="views-field views-field-field-distribution-requirements")
        course_bdt = course.find("span", class_="views-field views-field-field-breadth-requirements")
        course_pre = course.find("span", class_="views-field views-field-field-prerequisite")
        course_core = course.find("span", class_="views-field views-field-field-corequisite")
        course_hours = course.find("span", class_="views-field views-field-field-hours")

        if course_name is not None:
            course_name = course_name.text.strip()
        else:
            course_name = "N/A"
        if course_desc is not None:
            course_desc = course_desc.text.strip()
        else:
            course_desc = "N/A"
        if course_dst is not None:
            course_dst = course_dst.text.strip()
        else:
            course_dst = "N/A"
        if course_bdt is not None:
            course_bdt = course_bdt.text.strip()
        else:
            course_bdt = "N/A"
        if course_pre is not None:
            course_pre = course_pre.text.strip()
        else:
            course_pre = "N/A"
        if course_core is not None:
            course_core = course_core.text.strip()
        else:
            course_core = "N/A"
        if course_hours is not None:
            course_hours = course_hours.text.strip()
        else:
            course_hours = "N/A"

        dct["Course Name"] = course_name
        dct["Course Description"] = course_desc
        dct["Distribution Requirements"] = course_dst
        dct["Breadth Requirements"] = course_bdt
        dct["Prerequisites"] = course_pre
        dct["Corequisites"] = course_core
        dct["Hours"] = course_hours

        data.append(dct)

for i in data:
    if i["Course Name"] == "N/A":
        data.remove(i)

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = input("Enter the name of the sheet: ")

sheet["A1"] = "Course Name"
sheet["B1"] = "Course Description"
sheet["C1"] = "Distribution Requirements"
sheet["D1"] = "Breadth Requirements"
sheet["E1"] = "Prerequisites"
sheet["F1"] = "Corequisites"
sheet["G1"] = "Hours"

for i in range(0, len(data)):
    sheet[f"A{i + 2}"] = data[i]["Course Name"]
    sheet[f"B{i + 2}"] = data[i]["Course Description"]
    sheet[f"C{i + 2}"] = data[i]["Distribution Requirements"]
    sheet[f"D{i + 2}"] = data[i]["Breadth Requirements"]
    sheet[f"E{i + 2}"] = data[i]["Prerequisites"]
    sheet[f"F{i + 2}"] = data[i]["Corequisites"]
    sheet[f"G{i + 2}"] = data[i]["Hours"]

workbook.save(filename=sheet.title + ".xlsx")

driver.quit()
