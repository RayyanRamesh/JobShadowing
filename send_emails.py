import smtplib
import xlrd

# from ical_invite import create_ical_file
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


STYLE = """<html><head><style TYPE="text/css">
           <!--
           body,td,th {font-family:Calibri;font-size: 12px;}
           table.overview {border:1px solid black;border-collapse: collapse;}
           tbody.th {display:none;}
           table {width:100%;}
           th.overview {border:1px solid black;padding:2px;}
           td.overiew {border:1px solid black;padding:2px;}
           table.border {border:1px solid black;border-collapse: collapse;}
           th.border {width:6.25%;border:1px solid black;padding:2px;}
           td.border {border:1px solid black;padding:2px;}
           table.perf {border:1px solid black;border-collapse: collapse;width:50%;}
           td.neg {border:1px solid black;padding:2px;color:red;font-weight:bold;}
            #background-color:#D8D8D8
           body.positions,td.positions,th.positions {font-family:Courier;font-size: 12px;}
           -->
           </style>
           </head><body>"""


class Match:
    '''
    Each match between an alum and student(s) is saved into an Match object.
    This Class contains a string for email recipient and the email body.
    '''

    def __init__(self, alum_name, alum_email, alum_discipline, alum_job, alum_company, alum_careers, alum_adress, alum_job_description, alum_benefit_long_ans,
                 student1_name, student1_email, student1_year, student1_discipline, student1_option, student1_careers, student1_long_ans,
                 student2_name, student2_email, student2_year, student2_discipline, student2_option, student2_careers, student2_long_ans, avail_date):

        self.alum_name = alum_name
        self.alum_email = alum_email
        self.alum_discipline = alum_discipline
        self.alum_job = alum_job
        self.alum_company = alum_company

        self.student1_name = student1_name
        self.student1_email = student1_email
        self.student1_year = student1_year
        self.student1_discipline = student1_discipline
        self.student1_option = student1_option
        self.student1_long_ans = student1_long_ans

        self.student2_name = student2_name
        self.student2_email = student2_email
        self.student2_year = student2_year
        self.student2_discipline = student2_discipline
        self.student2_option = student2_option
        self.student2_long_ans = student2_long_ans
        self.avail_date = avail_date

    def email_to(self):
        if self.student2_name == 'N/A':
            return self.alum_email + ', ' + self.student1_email
        return self.alum_email + ', ' + self.student1_email + ', ' + self.student2_email

    #TODO: Change the email signature for future emails
    def email_body(self):
        # Alum matched with one student
        if self.student2_name == 'N/A':
            message = 'Hi {Alum_Name} and {Student1_Name},  \n \n' \
                       'Thank you both for signing up for the Alumni Relations Committee’s Job Shadowing program! ' \
                      'We have matched the two of you together.  \n\n'.format(Alum_Name=self.alum_name.split(' ')[0],
                                                                        Student1_Name=self.student1_name.split(' ')[0])
            # Change message if the student is a First Year
            if self.student1_year =='First Year':
                message += '{Student1_Name} is a {Student1_Year} student. '.format(
                                                                          Student1_Name=self.student1_name,
                                                                          Student1_Year=self.student1_year)
            else:
                message += '{Student1_Name} is a {Student1_Year} in {Student1_Discipline}. '.format(
                    Student1_Name=self.student1_name,
                    Student1_Year=self.student1_year,
                    Student1_Discipline=self.student1_discipline)

            message += '{Alum_Name} is a {Alum_Discipline} grad and works at {Alum_Company} as a {Alum_title}. ' \
                          'You both indicated being available on {Avail_date}. Please coordinate between the two of you to arrange the details for your job ' \
                          'shadowing day (ie. work location, meeting time, instructions for entering the building, etc.). ' \
                          'Let me know as soon as possible if one of your availabilities has changed and you need me to arrange a new match.  \n\n' \
                          'We asked all students in their application to write a couple of sentences on what they would be interested in doing ' \
                          'during their job shadowing day. We hope this will help alumni plan what to show the students. The following is ' \
                          '{Student1_Name}’s response:  \n{Student1_Long_Ans}  \n\n' \
                          'I hope you all enjoy your job shadowing day! Feel free to reach out if you have any questions. \n \n' \
                          'Thanks, \n' \
                          'Marissa Matthews \n' \
                          'Alumni Relations Chair | Queen’s Engineering Society \n' \
                          'alumnirelations@engsoc.queensu.ca \n'.format(Alum_Name=self.alum_name,
                                                                          Student1_Name=self.student1_name,
                                                                          Student1_Year=self.student1_year,
                                                                          Student1_Discipline=" ".join(self.student1_discipline),
                                                                          Alum_Discipline=self.alum_discipline,
                                                                          Alum_Company=self.alum_company,
                                                                          Alum_title=self.alum_job,
                                                                          Avail_date=self.avail_date,
                                                                          Student1_Long_Ans=self.student1_long_ans)
        else
            # Alum is matched with two students
            message = 'Hi {Alum_Name}, {Student1_Name}, and {Student2_Name}, \n\n' \
                      'Thank you all for signing up for the Alumni Relations ' \
                      'Committee’s Job Shadowing program! We have matched the three of you together. \n\n'.format(Alum_Name=self.alum_name.split(' ')[0],
                                                                          Student1_Name=self.student1_name.split(' ')[0], Student2_Name=self.student2_name.split(' ')[0])
            # Change message if the student is a First Year
            if self.student1_year == 'First Year':
                message += '{Student1_Name} is a {Student1_Year} student. '.format(
                    Student1_Name=self.student1_name,
                    Student1_Year=self.student1_year)
            else:
                message += '{Student1_Name} is a {Student1_Year} in {Student1_Discipline}. '.format(
                    Student1_Name=self.student1_name,
                    Student1_Year=self.student1_year,
                    Student1_Discipline=self.student1_discipline)
            # Change message if the student is a First Year
            if self.student2_year == 'First Year':
                message += '{Student2_Name} is a {Student2_Year} student. '.format(
                    Student2_Name=self.student2_name,
                    Student2_Year=self.student2_year)
            else:
                message += '{Student2_Name} is a {Student2_Year} in {Student2_Discipline}. '.format(
                    Student2_Name=self.student2_name,
                    Student2_Year=self.student2_year,
                    Student2_Discipline=self.student2_discipline)
            message += '{Alum_Name} is a {Alum_Discipline} grad and works at {Alum_Company}. ' \
                      'You all indicated being available on {Avail_date}. Please coordinate between the three of you to arrange the details for your job ' \
                      'shadowing day (ie. work location, meeting time, instructions for entering the building, etc.). ' \
                      'Let me know as soon as possible if one of your availabilities has changed and you need me to arrange a new match. \n\n' \
                      'We asked all students in their application to write a couple of sentences on what they would be interested in doing ' \
                      'during their job shadowing day. We hope this will help alumni plan what to show the students. The following is ' \
                      '{Student1_Name}’s response: \n{Student1_Long_Ans} \n\n Here is {Student2_Name}’s response: \n{Student2_Long_Ans} \n\n' \
                      'I hope you all enjoy your job shadowing day! Feel free to reach out if you have any questions.\n\n' \
                      'Thanks,\n' \
                      'Marissa Matthews\n' \
                      'Alumni Relations Chair | Queen’s Engineering Society\n' \
                      'alumnirelations@engsoc.queensu.ca\n'.format(Alum_Name=self.alum_name, Student1_Name=self.student1_name, Student2_Name=self.student2_name,
                        Student1_Year=self.student1_year, Student1_Discipline=self.student1_discipline, Student2_Year=self.student2_year, Student2_Discipline=self.student2_discipline,
                        Alum_Discipline=self.alum_discipline, Alum_Company=self.alum_company, Avail_date=self.avail_date,
                        Student1_Long_Ans=self.student1_long_ans, Student2_Long_Ans=self.student2_long_ans)

        return message

def _read_matches(matches_file_path):
    # To open Workbook
    wb = xlrd.open_workbook(matches_file_path)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    match_list = []

    for i in range(1, sheet.nrows):
        curr = Match(
            sheet.cell_value(i, 1),  # alum name
            sheet.cell_value(i, 2),  # alum email
            sheet.cell_value(i, 3),  # alum discipline
            sheet.cell_value(i, 4),  # Alum Job
            sheet.cell_value(i, 5),  # Alum Company
            sheet.cell_value(i, 6),  # Alum Careers
            sheet.cell_value(i, 7),  # Alum Address
            sheet.cell_value(i, 8),  # Alum job description
            sheet.cell_value(i, 9),  # Alum benefit long ans

            sheet.cell_value(i, 10),  # student 1 name
            sheet.cell_value(i, 11),  # student 1 email
            #sheet.cell_value(i, 12),  # student 1 location
            sheet.cell_value(i, 13),  # student 1 year
            sheet.cell_value(i, 14),  # student 1 discipline
            sheet.cell_value(i, 15),  # student 1 option
            sheet.cell_value(i, 16),  # student 1 careers of interest
            sheet.cell_value(i, 17),  # student 1 long ans

            sheet.cell_value(i, 18),  # student 2 name
            sheet.cell_value(i, 19),  # student 2 email
            sheet.cell_value(i, 20),  # student 2 year
            #sheet.cell_value(i, 21), # student 2 location
            sheet.cell_value(i, 22),  # student 2 discipline
            sheet.cell_value(i, 23),  # student 2 option
            sheet.cell_value(i, 24),  # student 2 careers of interest
            sheet.cell_value(i, 25),  # student 2 long ans

            sheet.cell_value(i, 26)   # avail dates
        )
        match_list.append(curr)

    return match_list


class Emailer():
    #TODO: Change stmp information for current person emailing
    def __init__(self):
        self.smtp = smtplib.SMTP(host='smtp.gmail.com', port=587)
        myGmail = 'laure.halabi@gmail.com'
        myGMPasswd ='    '
        self.smtp.starttls()
        self.smtp.login(myGmail, myGMPasswd)
    def send_email(self, email_from, email_to, email_cc, subject, html_body, include_style=False, send_invite=False):
        print("email_to is: {}".format(email_to))
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = email_from
        msg['To'] = email_to
        msg['CC'] = email_cc

        # Include the style can change the html formatting EOD Performance report doesn't require the styling
        if include_style:
            msg.attach(MIMEText(STYLE + html_body, 'html'))
        else:
            msg.attach(MIMEText(html_body, 'html'))

        self.smtp.sendmail(email_from, email_to, msg.as_string())

    def __del__(self):
        self.smtp.quit()

# Function called by main to email out all the matches
def _email_matches(file_path):
    match_list = _read_matches(file_path)
    email_subject = 'Job Shadowing Match'
    _send_email_to_groups(email_subject, match_list)

# Function to save all outgoing emails to a text file
#TODO: create a text file for the rejected participants
def _save_emails_to_file(file_path):
    match_list = _read_matches(file_path)
    f = open('job_shadowing_emails.txt', 'w+')
    for match in match_list:
        f.write('Subject: Job Shadowing Match \n')
        f.write('Email to: {} \n'.format(match.email_to()))
        f.write('Email body: \n{}\n\n\n'.format(match.email_body()))

# Function to send out emails to all the matches
def _send_email_to_groups(email_subject, matches):
    organizer = 'laure.halabi@gmail.com'
    email_subj = email_subject

    for match in matches:
        email_message = match.email_body()
        Emailer().send_email(email_from=organizer, email_to=match.email_to(), email_cc=organizer, subject=email_subj,
                             html_body=email_message)

    # testing for a singular match 
    # match = matches[0]
    # email_message = match.email_body()
    # Emailer().send_email(email_from=organizer, email_to=match.email_to(), email_cc=organizer, subject=email_subj,
    #                   html_body=email_message)


















