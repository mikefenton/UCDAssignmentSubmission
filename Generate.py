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
from __future__ import print_function
from xlutils.copy import copy
from os.path import join
from os import path, mkdir, chdir, remove, getcwd
from xlrd import open_workbook, xldate_as_tuple
from datetime import date
from subprocess import Popen, PIPE
import qrcode
import Submit
from Tkinter import Tk, Label, Entry, Button
import sys


DEBUG = False


def check_email_count(main_dir):
    """
    Checks the email count file.

    :param main_dir: The directory holding the file.
    :return: The current count.
    """

    # Check mail counter.
    counter_file = join(main_dir, 'mailcounter.txt')
    
    # Check mail counter exists.
    if path.exists(counter_file):
        log = open(counter_file, "r")

        # Get current count value.
        counter = int(eval(log.read()))

        log.close()
        
    else:
        # Write mail counter, set value to 0.
        log = open(counter_file, "w")
        log.write("0")
        log.close()
        
        counter = 0

    return counter


def write_email_count(count, main_dir):
    """
    Writes the new count file.

    :param count: The current count to write.
    :param main_dir: The directory holding the file.
    :return: Nothing.
    """

    # Check mail counter.
    counter_file = join(main_dir, 'mailcounter.txt')

    log = open(counter_file, "w")

    # Update the mail counter log.
    log.write(str(count))
    log.close()


def generate_folders(course, assignment, current_dir):
    """
    Check all potential file paths, and generate new folders if they aren't
    there already.
    
    :param course: The module code for the current course.
    :param assignment: The name of the current assignment.
    :param current_dir: The current working directory.
    :return: Nothing.
    """

    # Main courses directory.
    courses_dir = join(current_dir, "Courses")
    
    if not path.isdir(courses_dir):
        # Generate this folder.
        mkdir(courses_dir)

    # Individual course directory.
    course_dir = join(courses_dir, course)
    
    if not path.isdir(course_dir):
        # Generate this folder.
        mkdir(course_dir)
        
    # Assignment directory.
    ass_dir = join(course_dir, assignment)
    
    if not path.isdir(ass_dir):
        # Generate this folder.
        mkdir(ass_dir)


def generate_codes(course, current_dir):
    """
    Opens the spreadsheet named under the course code and adds
    the student details to a list. If student has already been
    given a cover sheet, they are ignored.
    
    :param course: The module code for the current course.
    :param current_dir: The current working directory.
    :return:
    """

    # Check email counter.
    counter = check_email_count(current_dir)

    print("Daily mail limit: %d more mails can be sent today." %
          (2000 - counter))

    if counter >= 2000:
        # if the email limit has been previously reached, we can not send
        # any more emails today.
        email = False

    else:
        # Limit not reached, we can send emails.
        email = True

    # Get course excel file.
    xls_file = join(current_dir, 'grive', course + ".xls")

    # Open workbook.
    workbook = open_workbook(xls_file, on_demand=True, formatting_info=True)

    # locates the relevant assignment (the name of the assignment)
    workbook.release_resources()

    # Loop over the worksheet names.
    for assignment in workbook.sheet_names():
        print("Assignment: %s" % assignment)

        # Ensure the folders are in place.
        generate_folders(course, assignment, current_dir)

        # Re-open workbook every time.
        workbook = open_workbook(xls_file, on_demand=True,
                                 formatting_info=True)

        # Initialise empty list of students.
        student_list = []

        # Get current worksheet.
        worksheet = workbook.sheet_by_name(str(assignment))

        try:
            # Find the deadline.
            deadline = xldate_as_tuple(worksheet.cell_value(1, 1), 0)
        
        except ValueError:
            # Catch error if no date is specified.
            
            print("Error: The deadline date for %s is specified incorrectly "
                  "for %s." % (course, assignment))
            quit()

        # Skip the first three rows...
        curr_row = 2
        while curr_row < (worksheet.nrows - 1):

            # Create new student
            student = []
            curr_row += 1
            curr_cell = -1
            while curr_cell < (worksheet.ncols - 1):
                curr_cell += 1
                cell_value = worksheet.cell_value(curr_row, curr_cell)
                student.append(str(cell_value))

            # No more students.
            if student[0].strip() == "":
                break

            pupil = {'deadline': date(deadline[0], deadline[1], deadline[2]),
                     'tutor': worksheet.cell_value(0, 1),
                     'student': student[0],
                     'course': course,
                     'emailed': student[4],
                     'assignment': assignment,
                     'row': curr_row,
                     'email': student[2],
                     'student number': "%08d" % int(float(student[1])),
                     'generated': student[3]}

            if pupil['generated'] and pupil['emailed']:
                # We have already generated this student's cover sheet and
                # emailed it to them.
                pass

            else:
                # We have to do something with this student.
                student_list.append(pupil)

        # Create a copy of the workbook as we'll be modifying it.
        work_book = copy(workbook)

        # Get the worksheet of the current assignment from the copied book.
        worksheet_2 = Submit.get_sheet_by_name(work_book, str(assignment))

        # Now enter the "doing stuff" loop where we create cover sheets and
        # email students.

        for stud in student_list:
            if not stud['generated']:
                # The student's cover sheet has not already been generated.
                encode(stud, current_dir)
                write_cover_sheet(stud, current_dir)

            if email and not DEBUG and not stud['emailed']:
                # We are allowed email people.

                if counter < 2000:

                    # Email cover sheet.
                    email_cover_sheet(stud, current_dir)

                    # Log that we have emailed a cover sheet to this student.
                    worksheet_2.write(int(stud['row']), 4, True)

                    # Increment our email counter.
                    write_email_count(check_email_count(current_dir) + 1,
                                      current_dir)

                else:
                    print("Daily email limit has been reached. Please try "
                          "again tomorrow.")
                    email = False

                    # Write the max limit emails.
                    write_email_count(2000, current_dir)
                    
                    quit()

            # Log that we have generated a cover sheet for this student.
            worksheet_2.write(int(stud['row']), 3, True)

            # Save the workbook
            work_book.save(xls_file)


def encode(student, current_dir):
    """
    Generates a batch of QR codes for a given list of students
    If there is no folder for the current assignment, one is generated.
    
    :param student: A dictionary describing a student.
    :param current_dir: The current working directory.
    :return: Nothing.
    """
    
    encoder = True
    try:
        # There are different QR generating modules. This checks which
        # one is installed and proceeds accordingly.
        qr_code = qrcode.QRCode()
        encoder = False
    
    except:
        qr_code = qrcode.Encoder()
    
    # Generate string of data to be encoded in QR code.
    student_data = str(student['course']) + '\n' + \
                   str(student['assignment']) + '\n' + \
                   str(student['student']) + '\n' + \
                   str(student['student number']) + '\n' + \
                   str(student['email'])
    
    if not encoder:
        # Use the QRCode module.
        qr_code.add_data(student_data)
        qr_code.make(fit=True)
        
        # Generate the QR image.
        image = qr_code.make_image()
    
    else:
        # Use the Encoder module.
        image = qr_code.encode(student_data)
    
    # Save the QR code image.
    image.save(join(current_dir, 'Courses', student['course'],
                    student['assignment'], str(student['student number']) +
                    '.png'))


def write_cover_sheet(student, current_dir):
    """
    Creates an individual cover sheet for each student, complete with
    course code, assignment name, student name, student number, and
    the QR code, all wrapped up in a LaTeX file. Uses PDFLaTeX to convert
    this LaTeX file into a PDF.
    
    :param student: A dictionary describing a student.
    :param current_dir: The current working directory.
    :return: Nothing.
    """
    
    # Print out useful info.
    print("  ", student['student number'], "\t", student['assignment'],
          "\tGenerating cover sheet.")
    
    # Navigate to correct directory.
    chdir(join(current_dir, 'Courses', student['course'],
               student['assignment']))
        
    # Write latex file for generating cover sheet.
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
    
    # Run pdflatex to generate cover sheet.
    cmd = 'pdflatex ' + str(student['student number']) + '.tex'
    process = Popen(cmd, shell=True, stdout=PIPE, stdin=PIPE)
    process.communicate()
    
    # Delete unnecessary LaTeX log files.
    remove(str(student['student number']) + '.tex')
    remove(str(student['student number']) + '.log')
    remove(str(student['student number']) + '.aux')
    remove(str(student['student number']) + '.png')
    remove(str(student['student number']) + '.out')


def email_cover_sheet(student, current_dir):
    """
    Sends a formal email to the student with their cover sheet as an
    attachment.
    
    :param student: A dictionary defining a particular student.
    :param current_dir: The current working directory.
    :return: Nothing.
    """
    
    print("  ", student['student number'], "\t", student['assignment'],
          "\tEmailing cover sheet.")
    
    subject = 'UCD ' + str(student['course']) + ' assignment: ' + str(
              student['assignment']) + ' Cover Sheet'
    
    name = student['student'].split()
    
    body = "Dear " + name[0] + ", \n\n"\
           "Please find attached the assignment cover sheet necessary for you"\
           "to submit " + student['assignment'] + " for your registered "\
           "course " + student['course'] + ". The deadline for submission of "\
           "this assignment is " + str(student['deadline']) + " at 15:00. " \
           "Please download this attachment for your future reference as "\
           "replacements will not be issued. You are responsible for ensuring"\
           "that the correct cover sheet is signed and attached to your "\
           "assignment when submitting.\n\nPlease ensure that you use the QR "\
           "code on the top right hand corner of the cover sheet to log your "\
           "submission using the system in the Civil Engineering School "\
           "office located in Newstead, or else your assignment will be "\
           "registered as overdue and the course coordinator notified.\n\n"\
           "This is an automated email. Please do not respond to this email."
    
    attachment = join(current_dir, 'Courses', student['course'],
                      student['assignment'], str(student['student number'])
                      + '.pdf')
    
    Submit.send_email(student, subject, body, attachment)


class mainWindow(object):

    """
    Main class for course input popup window.
    """

    def __init__(self, master, main_dir):
        self.main_dir = main_dir
        self.master = master
        self.label = Label(master,
                    text="Enter a course code (e.g. CVEN10010)")
        self.label.pack()
        self.entry = Entry(master)
        self.entry.pack()
        self.button = Button(master, text="Generate QR Codes",
                         command=self.cleanup)
        self.button.pack()

    def cleanup(self):
        """
        Runs the program based on the input given in text entry box
        """
        generate_codes(self.entry.get(), self.main_dir)
        self.master.destroy()


if __name__ == '__main__':

    main_dir = getcwd()

    course = sys.argv[1:][0]

    if course:
        generate_codes(course, main_dir)
    else:
        ROOT = Tk()
        MAIN = mainWindow(ROOT, main_dir)
        COURSE = ROOT.mainloop()

    # Sync google drive folder with online repository.
    chdir(join(main_dir, "grive"))
    PROCESS = Popen('grive', shell=True, stdout=PIPE, stdin=PIPE)
    PROCESS.communicate()
