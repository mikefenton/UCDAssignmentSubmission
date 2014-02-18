OVERVIEW:
    Generate.py
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

    Submit.py
        This program allows students to submit their hardcopy assignments
        by scanning a QR code on the assignment cover. The program logs the
        student's information and the time of submission to a 
        spreadsheet/workbook for that particular course.

        If the submission is successful the student is shown a "success"
        image and the program automatically sends them an email notifying them
        of their successful submission. If there is a problem, however, 
        an email is sent to maintenance with details of the issue.
        
INSTALLATION: 
    software requirements/packages
        zbar-0.10
            Barcode and QR code reader, required to decode QR codes
        Tkinter
            Basic GUI implementation
        qrtools.py
            Library for encoding/decoding QR Codes
        Google Drive
            Requires a Google Drive directory (default name is "grive")
            to be installed. All course workbooks are kept within.
            Every time a student submits an assignment, the Google drive
            is synced and updated. This way lecturers can keep track
            of student submissions in real time.
        Email Account access
            The program sends automatic emails, both to students and
            for maintenance purposes. Access to a working email address
            is required. Generic Google addresses tend to be caught
            by spam filters.
        Dropbox
            Not essential, but useful for version control and maintenance
            across multiple machines.

USEAGE:
    HOW TO GENERATE QR CODES:
        1. Plug in a keyboard to the assignment submission computer
        2. Type "Ctrl + Alt + T" to bring up the command terminal.
        3. Type "cd Dropbox/QR/" to navigate to the folder where everything 
           lives
        4. Type "python Generate.py" to run the program which generates cover 
           sheets.

        Once the program fires up, a pop-up window will appear and will ask
        you for the course code for the particular module you wish to generate
        cover sheets for.
        An example with the correct format is shown in the box that pops up.
        Simply fill in the course code, hit the "Generate QR Codes" button,
        and the program will take care of everything else.
    HOW TO SUBMIT AN ASSIGNMENT:
        The easiest (most student-friendly way) to submit an assignment is
        to create a desktop launcher for the file Submit.py. The file simply
        needs to be executed and the student will do the rest.
        
        See the "Steps.pdf" file in the "Info" folder for more information
        on submission.

NOTES:
    Care must be taken to input the course code exactly as it appears in
    the "give" folder, including spaces or lack thereof. For example, the
    program will not be able to generate course codes for the course
    "CVEN10020" if the course name "CVEN 10020" is given as there is an
    extra space added and they consequently do not match.
    
    See sample worksheet for correct layout of course workbooks.
    
    Students' names should be in ASCII format only. No special characters
    such as á, è, é, ï, ñ, ß, ç, etc. as this will produce error notices.
    
    The "TRUE" fields (under "code generated" and "email sent") are used by
    the program to track which tasks it has completed. These fields are
    filled out automatically by the program itself once a task has been
    completed, so they need to be left blank.
    
     There should be no blank lines in the course master workbook, either
     between line 3 (containing the field "name", "student number", etc) or
     the first line of student data. The name and details of the first
     student on the list should be on line 4 of each worksheet.
     
     Take extra care to ensure that there are no extra rows after the last
     student's details, especially if student lists are copied over from
     other sources.
