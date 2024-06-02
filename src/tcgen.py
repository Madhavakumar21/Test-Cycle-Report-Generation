"""
---------------------------
TEST CYCLE REPORT GENERATOR
---------------------------

* Generates a test report for a particular test cycle.
* Uses xlrd module to read the data from a excel file.
* Report is generated as a html file.

"""


import xlrd
import datetime
from matplotlib import pyplot
#import pdfkit
#from ... import ...


class Excel_Content():
    """A specific structure for the excel content"""

    def __init__(self, content):
        self.project_name = content[1][2]
        self.test_report_details = self.get_table(content, 2, 1, 10, 2)
        self.test_cases_summary = self.get_table(content, 14, 1, 2, 6)
        self.bug_details = self.get_table(content, 18, 1, 5, 3)
        self.conclusion = content[25][1]


    def get_table(self, content, s_row, s_col, row_count, col_count):
        """Returns the specified table
        With the starting row & column (starting from zero) (all limits included) and,
            the row count & column count given"""

        table = []
        for row_ind in range(s_row, s_row + row_count):
            row = []
            for col_ind in range(s_col, s_col + col_count):
                row.append(content[row_ind][col_ind])
            table.append(tuple(row))
        return tuple(table)


def fetch_data(file_name):
    """Reads the given excel file and returns the processed data."""

    # open the excel file and the sheet
    location = (file_name)
    excel_file = xlrd.open_workbook(location)
    sheet = excel_file.sheet_by_index(0)

    # read all data
    content = []
    for row in range(26):
        content.append([])
        for col in range(7):
            value = sheet.cell_value(row, col)
            content[row].append(value)

    # process the data and return
    return Excel_Content(content)


def display_content(content):
    """Displays the provided content"""

    for row in range(25):
        for col in range(7):
            print(content[row][col], end = "\t\t")
        print()


def generate_pie_chart(details):
    """Plots a pie chart and saves it as an image file."""

    clrs = ['green', 'yellow', 'red', 'grey']
    lbls = ['Passed', 'Passed with deviation', 'Failed', 'Not Run']
    temp = details[1]
    data = [temp[3], temp[4], temp[5], temp[2]]

    # Removing the zero values (not to be plotted)
    n = 4
    i = 0
    while i < n:
        if data[i] == "0":
            data.pop(i)
            clrs.pop(i)
            lbls.pop(i)
            n -= 1
        else:
            i += 1

    # Plotting and saving
    pyplot.figure(0)
    pyplot.axis("equal")
    pyplot.pie(data, labels = lbls, colors = clrs, autopct = "%3d%%")
    pyplot.savefig("./test_cases_summary.png")


# ------------------------------------------------------------------------------
# HTML Related classes and functions
# ------------------------------------------------------------------------------


class HtmlReportContent(object):
    """Contains methods for easy manipulation of html file (report)"""

    def __init__(self):
        self.cnt = ""
        self.ind = 0

    def get_content(self):
        return self.cnt

    def write(self, data):
        """Inserts the given data (str) into the index point of the content (str).
        if Error: Returns 1
        Else: Returns 0"""
        if type(data) != str: return 1
        else:
            self.cnt = self.cnt[:self.ind] + data + self.cnt[self.ind:]
            self.ind += len(data)
            return 0

    def open_tag(self, tg_name, tg_class = "", tg_id = ""):
        """Inserts the given data (str) into the index point of the content (str).
        Index goes inbetween the tag's opening and closing tags.
        if Error: Returns 1 or 2 or 3
        Else: Returns 0"""
        if type(tg_name) != str: return 1
        elif type(tg_class) != str: return 2
        elif type(tg_id) != str: return 3
        else:
            if tg_class:
                if tg_id: self.write(f"<{tg_name} class = \"{tg_class}\" id = \"{tg_id}\">")
                else: self.write(f"<{tg_name} class = \"{tg_class}\">")
            else:
                if tg_id: self.write(f"<{tg_name} id = \"{tg_id}\">")
                else: self.write(f"<{tg_name}>")
            temp_ind = self.ind
            self.write(f"</{tg_name}>")
            self.ind = temp_ind
            return 0

    def go_front(self):
        """Moves the index after to the next closing tag.
        if Error: Returns 1
        Else: Returns 0"""

        temp_ind = self.ind
        check_1 = False
        check_2 = False
        while not(check_1 and check_2):
            try:
                if (self.cnt[temp_ind] == '<') and (self.cnt[temp_ind + 1] == '/'): check_1 = True
                elif check_1 and (self.cnt[temp_ind] == '>'): check_2 = True
            except IndexError: return 1
            temp_ind += 1
        self.ind = temp_ind
        return 0

    def go_back(self):
        """Moves the index before to the last opening tag.
        if Error: Returns 1
        Else: Returns 0"""

        temp_ind = self.ind
        check_1 = False
        check_2 = False
        while not(check_1 and check_2):
            try:
                if (self.cnt[temp_ind] == '>'): check_1 = True
                elif check_1 and (self.cnt[temp_ind - 1] == '<') and (self.cnt[temp_ind] != '/'):
                    check_2 = True
            except IndexError: return 1
            temp_ind -= 1
        self.ind = temp_ind
        return 0

    def go_into(self):
        """Moves the index after to the next opening tag.
        if Error: Returns 1
        Else: Returns 0"""

        temp_ind = self.ind
        check_1 = False
        check_2 = False
        while not(check_1 and check_2):
            try:
                if (self.cnt[temp_ind] == '<') and (self.cnt[temp_ind + 1] != '/'): check_1 = True
                elif check_1 and (self.cnt[temp_ind] == '>'): check_2 = True
            except IndexError: return 1
            temp_ind += 1
        self.ind = temp_ind
        return 0

    def generate(self):
        """Generates a report by writing the current content into an html file.
        The file name will be 'Report_<time-stamp>.html'."""

        # Generating time-stamp
        temp_ts = str(datetime.datetime.now())
        ts = ""
        for ele in temp_ts:
            if ele.isdigit(): ts += ele
            else: ts += '_'

        # Modifying time-stamp
        ts = ts[2:-6]           # removing first two digits of the year and last three digits of split seconds
        ts = ts.replace("_", "")# removing all the underscores

        # Generating report
        f_name = "Test_Cycle_Report_" + ts + ".html"
        with open(f_name, 'w') as f_h: f_h.write(self.cnt)


def generate_report(info):
    """Generates the HTML report for code metrics software based on the information given."""

    report_content = HtmlReportContent()
    insert_html_default_content(report_content, info)
    #insert_consolidated_table(info, report_content)
    #report_content.write("\n\n        <br><br>\n\n        ")
    #insert_complete_table(info, report_content)

    insert_test_report_details(report_content, info)
    insert_test_cases_summary(report_content, info)
    insert_bug_details(report_content, info)
    insert_conclusion(report_content, info)

    report_content.generate()


def generate_error_report(info, error):
    """Generates an HTML error report for code metrics software based on the error string given."""

    report_content = HtmlReportContent()
    insert_html_default_content(report_content, info)
    report_content.open_tag("h4")
    report_content.write("Sorry, an error had occured:")
    report_content.go_front()
    report_content.write("\n        ")
    report_content.open_tag("h6")
    report_content.write(error)
    report_content.generate()


def insert_html_default_content(html_content, info):
    """Gets an HtmlReportContent and inserts default html tags and css styles.
    Leaves the cursor inside the 'body' block."""

    html_content.write("<!DOCTYPE html>\n\n")
    html_content.open_tag("html")
    html_content.write("\n\n    ")
    html_content.open_tag("head")
    html_content.write("\n\n        ")
    html_content.open_tag("title")
    html_content.write("Test Cycle Report")
    html_content.go_front()
    html_content.write("\n\n        ")
    html_content.open_tag("style")

    #style_content = r'''.............'''
    #html_content.write(style_content.replace("???hash_tag???", "#"))
    styling_file = open("css_styling.txt", "r")
    html_content.write(styling_file.read())
    styling_file.close()

    html_content.go_front()
    html_content.write("\n\n    ")
    html_content.go_front()
    html_content.write("\n\n    ")
    html_content.open_tag("body")
    html_content.write("\n        <!-- NOTE: This is an auto-generated file -->\n\n        ")
    html_content.write('<img id = "logo" src = "./company_logo.jpg">')
    html_content.write("\n\n        ")
    html_content.open_tag("h1")

    #html_content.write("INSTRUMENT CLUSTER")
    html_content.write((info.project_name).upper())

    html_content.go_front()
    html_content.write("\n\n    ")
    html_content.go_front()
    html_content.write("\n\n")
    html_content.go_back()
    html_content.go_front()
    html_content.write("<br>\n        ")
    html_content.open_tag("h2")

    html_content.write("TEST CYCLE REPORT")

    html_content.go_front()
    html_content.write("\n        <br>\n\n        ")


def create_table(html_content, details, t_id = "", indent = ""):
    """Inserts a table into the html content."""

    html_content.write(indent)
    if t_id: html_content.open_tag("table", tg_id = t_id)
    else: html_content.open_tag("table")
    #html_content.write("\n" + indent)
    #html_content.go_back()
    #html_content.go_front()
    #html_content.write("\n")
    html_content.write("\n")
    html_content.write(indent + ("    " * 1))
    html_content.open_tag("thead")
    html_content.write("\n")

    flag = 1
    for row in details:
        html_content.write(indent + ("    " * 2))
        html_content.open_tag("tr")
        html_content.write("\n")
        for ele in row:
            html_content.write(indent + ("    " * 3))
            if flag == 1: html_content.open_tag("th")
            else: html_content.open_tag("td")
            html_content.write(ele)
            html_content.go_front()
            html_content.write("\n")
        html_content.write(indent + ("    " * 2))
        html_content.go_front()
        html_content.write("\n")
        if flag == 1:
            flag = 0
            html_content.write(indent + ("    " * 1))
            html_content.go_front()
            html_content.write("\n")
            html_content.write(indent + ("    " * 1))
            html_content.open_tag("tbody")
            html_content.write("\n")
    
    html_content.write(indent + ("    " * 1))
    html_content.go_front()
    html_content.write("\n")
    html_content.write(indent + ("    " * 0))
    html_content.go_front()


def insert_test_report_details(html_content, info):
    """Inserts the Test Report details including the title."""

    details = info.test_report_details

    html_content.open_tag("h3")
    html_content.write("Test Report Details")
    html_content.go_front()
    html_content.write("\n        <hr>\n\n        ")

    html_content.open_tag("table", tg_class = "section1")
    html_content.write("\n            ")
    html_content.open_tag("tbody")
    html_content.write("\n                ")
    html_content.open_tag("tr")
    html_content.write("\n                    ")
    html_content.open_tag("td", tg_class = "test_report_details")
    html_content.write("\n")
    indent = "                        "

    ################################
    html_content.write(indent)
    html_content.open_tag("table", tg_class = "section1")
    #html_content.write("\n" + indent)
    #html_content.go_back()
    #html_content.go_front()
    #html_content.write("\n")
    html_content.write("\n")
    html_content.write(indent + ("    " * 1))
    html_content.open_tag("tbody")
    html_content.write("\n")

    for row_no in range(len(details)):
        html_content.write(indent + ("    " * 2))
        html_content.open_tag("tr")
        html_content.write("\n")
        for col_no in range(len(details[row_no])):
            html_content.write(indent + ("    " * 3))
            if col_no == 0:
                html_content.open_tag("td", tg_class = "param")
                html_content.write(details[row_no][col_no])
            else:
                html_content.open_tag("td", tg_class = "value")
                html_content.write(": " + details[row_no][col_no])
            html_content.go_front()
            html_content.write("\n")
        html_content.write(indent + ("    " * 2))
        html_content.go_front()
        html_content.write("\n")
        if row_no == 2:
            html_content.write(indent + ("    " * 1))
            html_content.go_front()
            html_content.write("\n")
            html_content.write(indent + ("    " * 0))
            html_content.go_front()
            html_content.write("\n                    ")
            html_content.go_front()
            html_content.write("\n                    ")
            html_content.open_tag("td", tg_class = "test_report_details")
            html_content.write("\n")
            html_content.write(indent + ("    " * 0))
            html_content.open_tag("table", tg_class = "section1")
            html_content.write("\n")
            html_content.write(indent + ("    " * 1))
            html_content.open_tag("tbody")
            html_content.write("\n")
    
    html_content.write(indent + ("    " * 1))
    html_content.go_front()
    html_content.write("\n")
    html_content.write(indent + ("    " * 0))
    html_content.go_front()
    ################################

    html_content.write("\n")
    html_content.write("                    ")
    html_content.go_front()
    html_content.write("\n")
    html_content.write("                ")
    html_content.go_front()
    html_content.write("\n")
    html_content.write("            ")
    html_content.go_front()
    html_content.write("\n")
    html_content.write("        ")
    html_content.go_front()

    html_content.write("\n\n        <br><br>\n\n        ")


def insert_test_cases_summary(html_content, info):
    """Inserts the Test Cases summary including the title."""

    html_content.open_tag("h3")
    html_content.write("Test Cases Summary")
    html_content.go_front()
    html_content.write("\n        <hr>\n\n")
    create_table(html_content, info.test_cases_summary, t_id = "test_cases_summary", indent = "        ")
    html_content.write("\n\n        ")

    #html_content.write('<img src = "./example.png">')
    generate_pie_chart(info.test_cases_summary)
    html_content.write('<img src = "./test_cases_summary.png">')

    html_content.write("\n\n        ")


def insert_bug_details(html_content, info):
    """Inserts the Bug details including the title."""

    html_content.open_tag("h3")
    html_content.write("Bug Details")
    html_content.go_front()
    html_content.write("\n        <hr>\n\n")
    create_table(html_content, info.bug_details, t_id = "bug_details", indent = "        ")
    html_content.write("\n\n        ")
    


def insert_conclusion(html_content, info):
    """Inserts the Conclusion including the title."""

    html_content.open_tag("h3")
    html_content.write("Conclusion")
    html_content.go_front()
    html_content.write("\n        <hr>\n\n")
    html_content.open_tag("div", tg_id = "conclusion")
    html_content.write("\n            ")
    html_content.write(info.conclusion)
    html_content.write("\n        ")
    html_content.go_front()


if __name__ == '__main__':
    data = fetch_data("test_cycle.xls")
    #print(data.test_report_details)
    #print(data.test_cases_summary)
    #print(data.bug_details)
    #print(data.conclusion)

    generate_report(data)
    #generate_error_report(data, "Any error message.")


# END
