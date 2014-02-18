#!/usr/bin/python

"""UCD QR Code Assignment Submission Program
Copyright (c) 2013 Michael Fenton
Hereby licensed under the GNU GPL v3.

This program generates a batch of UCD Assignment
Submission cover sheets with the students' details
already filled out. Each student recieves a personalised
cover sheet for each assignment in each course. All the
student has to do is sign it and staple it to the
front of their assignment when submitting it to
the head office.

The program automatically emails the student with the cover
sheet as an attachment. One email is sent per cover sheet
(to facilitate new assignments being added a priori). If
new students are added to the master course worksheet 
after the fact, the program permits a new batch of cover
sheets to be emailed to the new students without re-sending
emails to all other students.

"""

from xlutils.copy import copy
from os.path import join
from os import path, mkdir, chdir, remove
from xlrd import open_workbook, xldate_as_tuple
from datetime import date
from subprocess import Popen, PIPE
import qrcode
import Submit
from Tkinter import Tk, Label, Entry, Button

DEBUG = False

def generate_codes(given_course):
    """
    Finds the current directory and generates the cover sheets
    """
    current_dir = path.expanduser('~') + '/Dropbox/QR'
    open_excel(given_course, current_dir)

def open_excel(course, current_dir):
    """
    Opens the spreadsheet named under the course code and adds 
    the student details to a list. If student has already been
    given a cover sheet, they are ignored.
    """
    workbook = open_workbook(str(current_dir) + '/grive/' + str(course)
        + ".xls", on_demand = True, formatting_info = True)
    # locates the relevant assignment (the name of the assignment)
    workbook.release_resources()
    if path.exists(str(current_dir) + '/mailcounter.txt'):
        log = open(str(current_dir) + '/mailcounter.txt', "r")
    else:
        log = open(str(current_dir) + '/mailcounter.txt', "w")
        log.write("0")
    figure = log.read()
    if figure:
        counter = float(figure) # set the count from previous emails
        # UCD emails have a sending limit of 2000 emails per day,
        # must continually check that we are under that limit.
    else:
        counter = 0
    print "Daily mail limit:", (2000 - counter),
    print "more mails can be sent today."
    log.close()
    log = open(str(current_dir) + '/mailcounter.txt', "w")
    if counter == 2000: # if the email limit has been previously reached
        email = False
    else: # if the limit has not been reached
        email = True
    for i in workbook.sheet_names():
        print "Assignment: ", str(i)
        workbook = open_workbook(str(current_dir) + '/grive/' + str(course)
            + ".xls", on_demand = True, formatting_info = True)
        student_list = []
        worksheet = workbook.sheet_by_name(str(i))
        deadline = xldate_as_tuple(worksheet.cell_value(1, 1), 0)
        # Skip the first three rows...
        curr_row = 2
        while curr_row < (worksheet.nrows - 1):
            student = []
            curr_row += 1
            curr_cell = -1	    
            while curr_cell < (worksheet.ncols - 1):
                curr_cell += 1
                cell_value = worksheet.cell_value(curr_row, curr_cell)
                student.append(str(cell_value))
            if student[0].strip() == "":
                break
            pupil = {'deadline':date(deadline[0], deadline[1], deadline[2]),
                     'tutor':worksheet.cell_value(0, 1), 'student':student[0],
                     'course':course, 'emailed':student[4],
                     'assignment':str(i), 'row':curr_row, 'email':student[2],
                     'student number':"%08d" % int(float(student[1])),
                     'generated':student[3]}
            if pupil['generated'] and pupil ['emailed']:
                pass
            else:
                student_list.append(pupil)
        work_book = copy(workbook)
        worksheet_2 = Submit.get_sheet_by_name(work_book, str(i))
        # Now enter the "doing stuff" loop where we
        # create cover sheets and email students
        for stud in student_list:
            if not stud['generated']:
            # if a student's cover sheet has not already been generated
                encode(stud, current_dir)
                write_cover_sheet(stud, current_dir)
            if email:
                if not DEBUG:
                    if not stud['emailed']:
                        if counter < 2000:
                            email_cover_sheet(stud, current_dir)
                            worksheet_2.write(int(stud['row']), 4, True)
                            counter += 1
                        else:
                            print "Daily email limit has been reached"
                            log.write("2000")
                            email = False
                            
            # Log that we have generated and emailed a cover sheet
            # to this individual
            worksheet_2.write(int(stud['row']), 3, True)
        work_book.save(join(str(current_dir) + '/grive/' + str(course
                       ) + ".xls"))
    log.write(str(counter))
    log.close()

def encode(student, current_dir):
    """
    Generates a batch of QR codes for a given list of students
    If there is no folder for the current assignment, one is generated.    
    """ 
    encoder = True
    try:
        # There are different QR generating modules. This checks which 
        # one is installed and proceeds accordingly.
        qr_code = qrcode.QRCode()
        encoder = False
    except:
        qr_code = qrcode.Encoder()
    # We create the file paths if they are not there
    if path.isdir(str(current_dir) + '/Courses/'):
        pass
    else:
        mkdir(str(current_dir) + '/Courses/')
    if path.isdir(str(current_dir) + '/Courses/' + str(student['course'])):
        pass
    else:
        mkdir(str(current_dir) + '/Courses/' + str(student['course']))
    if path.isdir(str(current_dir) + '/Courses/' + str(student['course'])
                     + "/Cover Sheets"):
        pass
    else:
        mkdir(str(current_dir) + '/Courses/' + str(student['course'])
                 + "/Cover Sheets")
    if path.isdir(str(current_dir) + '/Courses/' + str(student['course'])
                     + "/Cover Sheets/" + str(student['assignment'])):
        pass
    else:
        mkdir(str(current_dir) + '/Courses/' + str(student['course'])
                 + "/Cover Sheets/" + str(student['assignment']))
    student_data = str(student['course']) + '\n' + str(student['assignment']
            ) + '\n' + str(student['student']) + '\n' + str(
            student['student number']) + '\n' + str(student['email'])
    if not encoder:
        qr_code.add_data(student_data)
        qr_code.make(fit = True)
        image = qr_code.make_image()
    else:
        image = qr_code.encode(student_data)
    image.save(str(current_dir) + '/Courses/' + str(student['course'])
               + '/Cover Sheets/' + str(student['assignment'])
               + '/' + str(student['student number']) + '.png')

def write_cover_sheet(student, current_dir):
    """
    Creates an individual cover sheet for each student, complete with
    course code, assignment name, student name, student number, and
    the QR code, all wrapped up in a LaTeX file.
    """
    chdir(str(current_dir) + '/Courses/' + str(student['course'])
             + "/Cover Sheets/" + str(student['assignment']))
    temp = open(str(student['student number']) + '.tex', 'w')
    temp.write('\NeedsTeXFormat{LaTeX2e}\n\documentclass{article}\n\usepackag')
    temp.write('e[top    = 2.75cm,\nbottom = 0.50cm,\nleft   = 3.00cm,\nright')
    temp.write('  = 2.50cm]{geometry}\n\usepackage{setspace}\n\usepackage{arr')
    temp.write('ay}\n\usepackage{graphicx}\n\usepackage[document]{ragged2e}\n')
    temp.write('\usepackage{hyperref}\n\usepackage{tabulary}\n\usepackage[spa')
    temp.write('ce]{grffile}\n\\addtolength{\\voffset}{-50pt}\n\date{}\n\page')
    temp.write('numbering{gobble}\n\\begin{document}\n\includegraphics[width=')
    temp.write('2.5cm]{' + str(current_dir) + '/Images/Crest}\n\hfill\n\inclu')
    temp.write('degraphics[width=4cm]{' + str(student['student number']) + '}')
    temp.write('\n\n\\vspace{5 mm}\n\n\\textbf{\huge{Assessment Submission Fo')
    temp.write('rm}} \\\\\n\n\\vspace{5 mm}\n\n\\begin{tabular}{ | >{\\bfseri')
    temp.write('es}l | p{11cm}  |}\n  \hline\n  Student Name & ')
    temp.write(str(student['student']) + '\\\\[2.25ex]\n  \hline\n  Student N')
    temp.write('umber & ' + str(student['student number']) + ' \\\\[2.25ex]\n')
    temp.write('  \hline\n  Assessment Title & ' + str(student['assignment']))
    temp.write('\\\\[2.25ex]\n  \hline\n  Course & ' + str(student['course']))
    temp.write('\\\\[2.25ex]\n  \hline\n  Lecturer & ' + str(student['tutor']))
    temp.write('\\\\[2.25ex]\n  \hline\n  Tutor (if applicable) & \\\\[2.25ex')
    temp.write(']\n  \hline\n  OFFICE USE ONLY & \\\\\n  Date Recieved: & ')
    temp.write('\\\\\n  \hline\n  OFFICE USE ONLY & \\\\\n  Grade/Mark & ')
    temp.write('\\\\\n  \hline\n\end{tabular}\n\n\\vspace{5 mm}\n\\textbf{A S')
    temp.write('IGNED COPY OF THIS FORM MUST ACCOMPANY ALL SUBMISSIONS FOR AS')
    temp.write('SESSMENT.}\n\n\\vspace{5 mm}\n\n\\textbf{STUDENTS SHOULD KEEP')
    temp.write(' A COPY OF ALL WORK SUBMITTED.}\n\n\\vspace{5 mm}\n\n\\textbf')
    temp.write('{Procedures for Submission and Late Submission}\n\nEnsure tha')
    temp.write('t you have checked the School\'s procedures for the submissio')
    temp.write('n of assessments.\n\n\\textbf{Note:} There are penalties for ')
    temp.write('the late submission of assessments. ')
    temp.write('For further information please see the University\'s ')
    temp.write('\\textbf{\\textit{Policy on Late Submiss')
    temp.write('ion of Coursework,} (\hyperref[http://www.ucd.ie/registrar/]{')
    temp.write('http://www.ucd.ie/registrar/})}\n\n\\vspace{5 mm}\n\n\\textbf')
    temp.write('{Plagiarism:} the unacknowledged inclusion of another person')
    temp.write('\'s writing or ideas or works, in any formally presented work')
    temp.write(' (including essays, examinations, projects, laboratory repor')
    temp.write('ts or presentations). The penalties associated with plagiaris')
    temp.write('m designed to impose sanctions that reflect the seriousness o')
    temp.write('f University\'s commitment to academic integrity. Ensure that')
    temp.write(' you have read the University\'s \\textbf{\\textit{Briefing f')
    temp.write('or Students on Academic Integrity and Plagiarism}} and the UC')
    temp.write('D \\textbf{\\textit{Plagiarism Statement, Plagiarism Policy a')
    temp.write('nd Procedures,} (\hyperref[http://www.ucd.ie/registrar/]{http')
    temp.write('://www.ucd.ie/registrar/})}\n\n\\vspace{5 mm}\n\n\\begin{tabu')
    temp.write('lar}{| p{15cm} |}\n  \hline\n  \\\\\n  \\textbf{Declaration o')
    temp.write('f Authorship}\n\n  I declare that all material in this assess')
    temp.write('ment is my own work except where there is clear acknowledgeme')
    temp.write('nt and appropriate reference to the work of others.\n  \\\\\n')
    temp.write('  \\\\\n  \\\\\n  \\\\\n  \\\\\n  Signed \underline{\hspace{7')
    temp.write('cm}}\n  Date \underline{\hspace{5cm}}\n  \\\\\n  \\\\\n  \hli')
    temp.write('ne\n\n\end{tabular}\n\end{document}')
    temp.close()
    cmd = 'pdflatex ' + str(student['student number']) + '.tex'
    process = Popen(cmd, shell = True, stdout = PIPE,
                               stdin = PIPE)
    process.communicate()
    remove(str(student['student number']) + '.tex')
    remove(str(student['student number']) + '.log')
    remove(str(student['student number']) + '.aux')
    remove(str(student['student number']) + '.png')
    remove(str(student['student number']) + '.out')

def email_cover_sheet(student, current_dir):
    """
    Sends a formal email to the student with their cover sheet as an attachment
    """
    subject = 'UCD ' + str(student['course']) + ' assignment: ' + str(
              student['assignment']) + ' Cover Sheet'
    name = student['student'].split()
    body = "Dear " + str(name[0]) + ', \n\nPlease find attached the assign'\
           "ment cover sheet necessary for you to submit " + str(
           student['assignment']) + ' for your registered course ' + str(
           student['course']) + ". The deadline for submission of this assig"\
           "nment is " + str(student['deadline']) + " at 15:00. Please downl"\
           "oad this attachment for your future reference as replacements wil"\
           "l not be issued. You are responsible for ensuring that the correc"\
           "t cover sheet is signed and attached to your assignment when subm"\
           "itting.\n\nPlease ensure that you use the QR code on the top righ"\
           "t hand corner of the cover sheet to log your submission using the"\
           " system in the School of Civil, Structural, and Environmental Eng"\
           "ineering office located in Newstead, or else your assignment will"\
           " be registered as overdue and the course coordinator notified.\n\n"
    attachment = str(current_dir) + '/Courses/' + str(student['course']) + '/'\
                 'Cover Sheets/' + str(student['assignment']) + '/'\
                 + str(student['student number']) + '.pdf'
    Submit.send_email(student, subject, body, attachment)

class mainWindow(object):

    """
    Main class for course input popup window.
    """

    def __init__(self, master):
        self.master = master
        self.label = Label(master,
                    text = "Enter a course code (e.g. CVEN10010)")
        self.label.pack()
        self.entry = Entry(master)
        self.entry.pack()
        self.button = Button(master, text = "Generate QR Codes",
                         command = self.cleanup)
        self.button.pack()

    def cleanup(self):
        """
        Runs the program based on the input given in text entry box
        """
        generate_codes(self.entry.get())
        self.master.destroy()

if __name__ == '__main__':
    if DEBUG:
        generate_codes("CVEN20110")
    else:
        ROOT = Tk()
        MAIN = mainWindow(ROOT)
        COURSE = ROOT.mainloop()
    chdir(path.expanduser('~') + "/Dropbox/QR/grive")
    PROCESS = Popen('grive', shell = True, stdout = PIPE, stdin = PIPE)
    PROCESS.communicate()
