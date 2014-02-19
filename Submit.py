#!/usr/bin/python

"""UCD QR Code Assignment Submission Program
Copyright (c) 2013 Michael Fenton
Hereby licensed under the GNU GPL v3.

This program allows students to submit their hardcopy assignments
by reading a QR code on the assignment cover and logging the student's
information and the time of submission to a spreadsheet/workbook
for that particular course.

If the submission is successful the student is shown a "success"
image and the program automatically sends them an email notifying them
of their successful submission. If there is a problem, however, 
an email is sent to maintenance with details of the issue."""

from os.path import join
from xlutils.copy import copy
from xlrd import open_workbook, xldate_as_tuple
from sys import path
path.append('Packages/QRTools')
from qrtools import QR
from Tkinter import Tk, Button, Canvas
from datetime import datetime
from subprocess import PIPE, Popen
from itertools import count
from smtplib import SMTP
from email import mime
from ImageTk import PhotoImage
from os import path, chdir
import Image
import email.mime.application


DEBUG = False

def assignment_submission(cur_dir):
    """
    Records the time of submission and updates the spreadsheet
    """   
    values = decode()
    if not values[0]:
        failure(str(values[1]), cur_dir)
        exit()
    else:        
        text = values[0].split('\n')
        student = {'student':text[2], 'course':text[0], 'assignment':text[1],
                   'student number':"%08d" % int(float(text[3])),
                   'email':text[4], 'time':values[1]}
        if len(text) != 5:
            failure("QR Code Error: Insufficient information in QR code",
                    cur_dir)
            exit()
        deadline = save_excel(student, cur_dir)          
    amount = deadline - student['time']
    if student['time'] > deadline:
        expired(cur_dir, abs(amount))
        if not DEBUG:
            subject = "UCD " + str(student['course']) + ' assignment: ' + str(
                      student['assignment']) + " deadline expired"
            body = "The deadline for your submission " + str(student['course']
                   ) + ' ' + str(student['assignment']
                   ) + ' has expired by ' + str(amount
                   )+ '. Your submission at ' + str(student['time']
                   ) + " has been accepted, but the module coordinator has"\
                   "been notified."
            send_email(student, subject, body, None)
    else:   
        success(cur_dir)
        if not DEBUG:
            subject = "UCD " + str(student['course']) + ' assignment: ' + str(
                      student['assignment']) + " submission successful"
            body = "You have successfully submitted " + str(student['course']
                   ) + ' ' + str(student['assignment']) + ' at ' + str(
                   student['time']) + "."
            send_email(student, subject, body, None)

def decode():
    """
    Decodes the QR code and returns a list of information
    """
    my_code = QR()
    result = my_code.decode_webcam()
    if result:
        if result[0]:
            now = datetime.now()
            return [result[1], now]
        else:
            return [None, result[1]]
    else:
        return [None, "Failed to initialise webcam"]

def save_excel(student, cur_dir):
    """
    Opens up the relevant excel workbook and logs that the student
    has submitted Assignment X at time Y.
    """
    if path.exists(str(cur_dir) + '/grive/' + str(student['course']
                      ) + ".xls"):
        pass
    else:
        failure("QR Code Error: Course master spreadsheet " + str(
                cur_dir) + '/grive/' + str(student['course']
                ) + ".xls" + " doesn't exist", cur_dir)
        exit()
    if path.exists(str(cur_dir) + '/grive/' + str(
              student['course']) + "Log.txt"):
        log = open(str(cur_dir) + '/grive/' + str(
              student['course']) + "Log.txt", "a")
    else:
        log = open(str(cur_dir) + '/grive/' + str(
              student['course']) + "Log.txt", "w")
    log.write(str(student['time']) + "\t" + str(student['student number']
              ) + "\t" + str(student['assignment']) + "\n")
    log.close()
    book = open_workbook(join(str(cur_dir) + '/grive/' + str(
        student['course']) + ".xls"), formatting_info = True, on_demand = True)
    worksheet = book.sheet_by_name(str(student['assignment']))
    num_rows = worksheet.nrows - 1
    # Check submission deadline
    student['tutor'] = worksheet.cell_value(0, 1)
    deadline = xldate_as_tuple(worksheet.cell_value(1, 1), 0)
    # Deadline is automatically set at 3pm on the deadline date
    student['deadline'] = datetime(deadline[0], deadline[1],
                                            deadline[2], 18, 0, 0)
    work_book = copy(book)
    worksheet = get_sheet_by_name(work_book, student['assignment'])
    # look through all the student entries, match student number
    for i in range(num_rows-2):        
        if not int(book.sheet_by_name(str(student['assignment'])).cell(
                   i+3, 1).value) == int(student['student number']):
            if i == num_rows-3:
                failure("Student " + str(student['student number']) + " not "\
                       "on course " + str(student['course']) + " master "\
                       "spreadsheet", cur_dir)
                exit()
        else:    
            worksheet.write(i+3, 5, str(student['time']))
            if student['time'] > student['deadline']:
                worksheet.write(i+3, 6, "Deadline Expired")
            break
    work_book.save(join(str(cur_dir) + '/grive/' + str(
            student['course']) + ".xls"))
    return student['deadline']

def get_sheet_by_name(book, name):
    """
    Given a workbook, returns the names of the worksheets within.
    """
    for idx in count():
        sheet = book.get_sheet(idx)
        if sheet.name == name:
            return sheet

def send_email(student, subject, body, attachment):   
    """
    Sends a confirmation of submission email to the student's UCD 
    Connect email address using their student number.
    """
    fromaddr = '***************@*****.com'
    toaddrs = str(student['email'])
    msg = mime.Multipart.MIMEMultipart()
    msg['Subject'] = str(subject)
    msg['From'] = fromaddr
    msg['To'] = toaddrs
    body = mime.Text.MIMEText(str(body))
    msg.attach(body)
    if attachment:
        filename = str(attachment)
        open_attachment = open(filename, 'rb')
        att = mime.application.MIMEApplication(open_attachment.read(),
                                               _subtype = 'pdf')
        open_attachment.close()
        name = str(student['student number']) + ' ' + str(student['course']
                  ) + ' ' + str(student['assignment'])
        att.add_header('Content-Disposition', 'attachment', filename = name)
        msg.attach(att)
    username = '**************'
    password = '**************'
    server = SMTP('smtp.ucd.ie:587')
    server.ehlo()
    server.starttls()
    server.login(username, password)
    server.sendmail(fromaddr, toaddrs, msg.as_string())
    server.quit()

def success(cur_dir):
    """
    Displays a "successful submission" picture
    """
    root = Tk()
    root.focus_set()
    # Get the size of the screen and place the splash screen in the center
    img = Image.open(str(cur_dir) + '/Images/Success.gif')
    width = img.size[0]
    height = img.size[1]
    flog = (root.winfo_screenwidth()/2-width/2)
    blog = (root.winfo_screenheight()/2-height/2)
    root.overrideredirect(1)
    root.geometry('%dx%d+%d+%d' % (width, height, flog, blog))
    # Pack a canvas into the top level window.
    # This will be used to place the image
    success_canvas = Canvas(root)
    success_canvas.pack(fill = "both", expand = True)
    # Open the image
    imgtk = PhotoImage(img)
    # Get the top level window size
    # Need a call to update first, or else size is wrong
    root.update()
    cwidth = root.winfo_width()
    cheight =  root.winfo_height()
    # create the image on the canvas
    success_canvas.create_image(cwidth/2, cheight/2, image = imgtk)
    root.after(4000, root.destroy)
    root.mainloop()

def expired(cur_dir, amount):
    """
    Displays a "deadline expired" picture
    """
    root = Tk()
    root.focus_set()
    # Get the size of the screen and place the splash screen in the center
    img = Image.open(str(cur_dir) + '/Images/Expired.gif')
    width = img.size[0]
    height = img.size[1]
    flog = root.winfo_screenwidth()/2-width/2
    blog = root.winfo_screenheight()/2-height/2
    root.overrideredirect(True)
    root.geometry('%dx%d+%d+%d' % (width*1, height + 44, flog, blog))
    # Pack a canvas into the top level window.
    # This will be used to place the image
    expired_canvas = Canvas(root)
    expired_canvas.pack(fill = "both", expand = True)
    # Open the image
    imgtk = PhotoImage(img)
    # Get the top level window size
    # Need a call to update first, or else size is wrong
    root.update()
    cwidth = root.winfo_width()
    cheight =  root.winfo_height()
    # create the image on the canvas
    expired_canvas.create_image(cwidth/2, cheight/2.1, image = imgtk)
    Button(root, text = 'Deadline Expired by ' + str(amount
          ) + '. Assignment Submitted, time '\
          'noted', width = 80, height = 2, command = root.destroy).pack()
    root.after(5000, root.destroy)
    root.mainloop()
    
def failure(reason, cur_dir):
    """
    Displays a "submission failure" picture and emails
    a bug report to maintenance.
    """
    bugmail = {"email": "michaelfenton1@gmail.com"}
    send_email(bugmail, "QR Code Submission Failure", reason, None)
    root = Tk()
    root.focus_set()
    # Get the size of the screen and place the splash screen in the center
    gif = Image.open(str(cur_dir) + '/Images/Failure.gif')
    width = gif.size[0]
    height = gif.size[1]
    flog = (root.winfo_screenwidth()/2-width/2)
    blog = (root.winfo_screenheight()/2-height/2)
    root.overrideredirect(1)
    root.geometry('%dx%d+%d+%d' % (width*1, height + 44, flog, blog))
    # Pack a canvas into the top level window.
    # This will be used to place the image
    failure_canvas = Canvas(root)
    failure_canvas.pack(fill = "both", expand = True)
    # Open the image
    imgtk = PhotoImage(gif)
    # Get the top level window size
    # Need a call to update first, or else size is wrong
    root.update()
    cwidth = root.winfo_width()
    cheight =  root.winfo_height()
    # create the image on the canvas
    failure_canvas.create_image(cwidth/2, cheight/2.24, image=imgtk)
    Button(root, text = str(
        reason), width = 50, height = 2, command = root.destroy).pack()
    root.after(5000, root.destroy)
    root.mainloop()

if __name__ == '__main__':
    CURRENT_DIR = path.expanduser('~') + "/Dropbox/QR"    
    assignment_submission(CURRENT_DIR)
    # change directory to sync the course master spreadsheets
    chdir(str(CURRENT_DIR) + "/grive")
    PROCESS = Popen('grive', shell = True, stdout = PIPE, stdin = PIPE)
    PROCESS.communicate()
