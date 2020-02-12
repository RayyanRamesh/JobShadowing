
import xlrd
import xlwt

class Student:
    def __init__(self, name, email, telephone, year, discipline, option, city, avail_dates, transportation, career_fields, same_disp_flag, alum_contact_flag, benefit_long_ans, learn_long_ans, email_long_ans):
        self.name = name
        self.email = email
        self.telephone = str(telephone)
        self.year = year
        self.discipline = discipline
        self.option = option
        self.city = city
        # self.city_detail = city_detail

        for i in range(len(avail_dates)):
            if "Thursday" in avail_dates[i]:
                avail_dates[i] = 'Thursday, February 20, 2020'
            elif "Friday" in avail_dates[i]:
                avail_dates[i] = 'Friday, February 21, 2020'

        self.avail_dates = avail_dates
        self.transportation = transportation
        # self.transportation_time = transportation_time
        self.career_fields = career_fields
        self.same_disp_flag = same_disp_flag
        self.alum_contact_flag = alum_contact_flag
        self.benefit_long_ans = benefit_long_ans
        self.learn_long_ans = learn_long_ans
        self.email_long_ans = email_long_ans

    def profile(self):
        profile = "=====================================================================================================\n"\
                  + "=====================================================================================================\n" \
                  + 'Name: ' + self.name + '\n' \
                  + 'Email: ' + self.email + '\n' \
                  + 'Number: ' + self.telephone + '\n\n' \
                  + 'Academic Year: ' + self.year + '\n' \
                  + 'Discipline: ' + ', '.join(map(str, self.discipline)) + '\n' \
                  + 'Option: ' + self.option + '\n' \
                  + 'Career Fields of Interest: ' + ', '.join(map(str, self.career_fields)) + '\n\n' \
                  + 'City: ' + self.city + '\n' \
                  + 'Transportation: ' + self.transportation + '\n' \
                  + 'Dates Available: ' + '; '.join(map(str, self.avail_dates)) + '\n'\
                  + '_______________________________________________________________________________________________________\n'\
                  + 'How do you think you will benefit from the job shadowing program?\n' + self.benefit_long_ans + '\n' \
                  + '_______________________________________________________________________________________________________\n' \
                  + 'What would you like to learn from the alumni you are shadowing?\n' + self.learn_long_ans + '\n' \
                  + '_______________________________________________________________________________________________________\n' \
                  + 'Introduction email to your alumni match so they can have an idea of how you would like to spend your day with them\n' + self.email_long_ans + '\n'

        return profile


class Alumni:
    def __init__(self, name, email, telephone, discipline, job_title, tasks, company, career_fields, city, avail_dates, student_limit, same_disp_flag, benefit_long_ans):
        self.name = name
        self.email = email
        self.telephone = str(telephone)
        self.discipline = discipline
        self.job_title = job_title
        self.tasks_long_ans = tasks
        self.company = company
        self.career_fields = career_fields
        self.city = city
        # self.city_detail = city_detail

        for i in range(len(avail_dates)):
            if "Thursday" in avail_dates[i]:
                avail_dates[i] = 'Thursday, February 20, 2020'
            elif "Friday" in avail_dates[i]:
                avail_dates[i] = 'Friday, February 21, 2020'

        self.avail_dates = avail_dates
        self.student_limit = student_limit
        self.same_disp_flag = same_disp_flag
        self.benefit_long_ans = benefit_long_ans

    def profile(self):
        profile = "=====================================================================================================\n"\
                  + "=====================================================================================================\n" \
                  + 'Name: ' + self.name + '\n' \
                  + 'Email: ' + self.email + '\n' \
                  + 'Number: ' + self.telephone + '\n' \
                  + 'Student Limit: ' + str(self.student_limit) + '\n\n' \
                  + 'Discipline: ' + self.discipline + '\n' \
                  + 'Job Title: ' + self.job_title + '\n' \
                  + 'Company: ' + self.company + '\n' \
                  + 'Relevant Career Fields: ' + ', '.join(map(str, self.career_fields)) + '\n\n' \
                  + 'City: ' + self.city + '\n' \
                  + 'Dates Available: ' + '; '.join(map(str, self.avail_dates)) + '\n'\
                  + '_______________________________________________________________________________________________________\n' \
                  + 'What tasks does a typical day for you include?\n' + self.tasks_long_ans + '\n' \
                  + '_______________________________________________________________________________________________________\n' \
                  + 'Describe what type of student you feel would benefit the most from shadowing with you\n' + self.benefit_long_ans + '\n'

        return profile

def _read_data(student_file_path, alumni_file_path, matches_file_path):         #, day1, day2):

    # To open Workbook
    wb = xlrd.open_workbook(student_file_path)
    sheet = wb.sheet_by_index(0)

    student_list = []

    for i in range(1, sheet.nrows):
        curr = Student(
            sheet.cell_value(i, 1),  # name
            sheet.cell_value(i, 2),  # email
            sheet.cell_value(i, 3),  # telephone
            sheet.cell_value(i, 4),  # year
            [sheet.cell_value(i, 5)] if sheet.cell_value(i, 4) != 'First Year' else sheet.cell_value(i, 7).split(', '), # dicipline choices TODO:ASK about this
            "N/A" if sheet.cell_value(i, 6) is '' else sheet.cell_value(i, 6),  # option
            sheet.cell_value(i, 8), # + " - " + sheet.cell_value(i, 9),  # city
            sheet.cell_value(i, 10).split(', 2020,'),  # avail_dates
            sheet.cell_value(i, 11) + "\n" + sheet.cell_value(i, 12),  # transportation
            ["Anything"] if "Anything" in sheet.cell_value(i, 13) else sheet.cell_value(i, 13).split(", "),  # career_fields
            True if sheet.cell_value(i, 14) == "Yes" else False,  # same_disp_flag
            True if sheet.cell_value(i, 15) == "Yes" else False,   # alum_contact_flag
            "N/A" if sheet.cell_value(i, 16) is '' else sheet.cell_value(i, 16),    # benefit_long_ans
            "N/A" if sheet.cell_value(i, 17) is '' else sheet.cell_value(i, 17),  # learn_long_ans
            "N/A" if sheet.cell_value(i, 20) is '' else sheet.cell_value(i, 20), # email_long_ans
        )
        student_list.append(curr)

    # To open Workbook
    wb = xlrd.open_workbook(alumni_file_path)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    alumni_list = []

    for j in range(1, sheet.nrows):
        curr = Alumni(
            sheet.cell_value(j, 1),  # name
            sheet.cell_value(j, 2),  # email
            sheet.cell_value(j, 3),  # telephone
            sheet.cell_value(j, 5) if sheet.cell_value(j, 4) == 'Other' else sheet.cell_value(j, 4),  # discipline
            sheet.cell_value(j, 6),  # job title
            sheet.cell_value(j, 7),  # tasks
            sheet.cell_value(j, 8),  # company
            sheet.cell_value(j, 9).split(", "),  # career fields
            sheet.cell_value(j, 10), #+ " - " + sheet.cell_value(j, 11),  # city
            sheet.cell_value(j, 12).split(', 2020,'),  # avail_dates
            2 if "Two" in sheet.cell_value(j, 13) else 1,  # student limit
            True if sheet.cell_value(j, 14) == "Yes" else False,  # same_disp_flag
            "N/A" if sheet.cell_value(j, 15) is '' else sheet.cell_value(j, 15),  # benefit_long_ans
        )
        alumni_list.append(curr)

    wb = xlrd.open_workbook(matches_file_path)
    sheet = wb.sheet_by_index(0)
    good_matches ={}
    for j in range(1, sheet.nrows):
        good_matches[sheet.cell_value(j, 0)] = sheet.cell_value(j, 1)

    wb = xlrd.open_workbook(matches_file_path)
    sheet = wb.sheet_by_index(1)
    bad_matches = {}
    for j in range(1, sheet.nrows):
        bad_matches[sheet.cell_value(j, 0)] = sheet.cell_value(j, 1).split(';')

    return student_list, alumni_list, good_matches, bad_matches


def _save_data(matches, rejected_alumni, rejected_students):

    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('Matches')

    worksheet.write(0, 0, 'Match #')
    worksheet.write(0, 1, 'Alum Name')
    worksheet.write(0, 2, 'Alum Email')
    worksheet.write(0, 3, 'Alum Discipline')
    worksheet.write(0, 4, 'Alum Job Title')
    worksheet.write(0, 5, 'Alum Company')
    worksheet.write(0, 6, 'Alum Career Fields')
    worksheet.write(0, 7, 'Alum Work Address')
    worksheet.write(0, 8, 'Alum Job Description')
    worksheet.write(0, 9, 'Alum Student Benefit Description')
    worksheet.write(0, 10, 'Student 1 Name')
    worksheet.write(0, 11, 'Student 1 Email')
    worksheet.write(0, 12, 'Student 1 City')
    worksheet.write(0, 13, 'Student 1 Year')
    worksheet.write(0, 14, 'Student 1 Discipline')
    worksheet.write(0, 15, 'Student 1 Option')
    worksheet.write(0, 16, 'Student 1 Career Fields of Interest')
    worksheet.write(0, 17, 'Student 1 Long Ans')
    worksheet.write(0, 18, 'Student 2 Name')
    worksheet.write(0, 19, 'Student 2 Email')
    worksheet.write(0, 20, 'Student 2 City')
    worksheet.write(0, 21, 'Student 2 Year')
    worksheet.write(0, 22, 'Student 2 Discipline')
    worksheet.write(0, 23, 'Student 2 Option')
    worksheet.write(0, 25, 'Student 2 Career Fields of Interest')
    worksheet.write(0, 26, 'Student 2 Long Ans')
    worksheet.write(0, 27, 'Matching Details')

    for row in range(len(matches)):
        worksheet.write(row+1, 0, row)
        worksheet.write(row+1, 1, matches[row][0].name)
        worksheet.write(row+1, 2, matches[row][0].email)
        worksheet.write(row+1, 3, matches[row][0].discipline)
        worksheet.write(row+1, 4, matches[row][0].job_title)
        worksheet.write(row+1, 5, matches[row][0].company)
        worksheet.write(row+1, 6, matches[row][0].career_fields)
        worksheet.write(row+1, 7, matches[row][0].city)
        worksheet.write(row+1, 8, matches[row][0].tasks_long_ans)
        worksheet.write(row+1, 9, matches[row][0].benefit_long_ans)

        worksheet.write(row+1, 10, matches[row][1].name)
        worksheet.write(row+1, 11, matches[row][1].email)
        worksheet.write(row+1, 12, matches[row][1].city)
        worksheet.write(row+1, 13, matches[row][1].year)
        worksheet.write(row+1, 14, matches[row][1].discipline)
        worksheet.write(row+1, 15, matches[row][1].option)
        worksheet.write(row+1, 16, matches[row][1].career_fields)
        worksheet.write(row+1, 17, matches[row][1].email_long_ans)

        worksheet.write(row+1, 18, matches[row][2].name if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 19, matches[row][2].email if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 20, matches[row][2].city if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 21, matches[row][2].year if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 22, matches[row][2].discipline if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 23, matches[row][2].option if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 24, matches[row][2].career_fields if len(matches[row]) > 2 else "N/A")
        worksheet.write(row+1, 25, matches[row][2].email_long_ans if len(matches[row]) > 2 else "N/A")

        matching_details = list(set(matches[row][0].avail_dates) & set(matches[row][1].avail_dates))
        worksheet.write(row+1, 26, matching_details[0])

    worksheet = workbook.add_sheet("Rejected Students")
    worksheet.write(0, 0, '#')
    worksheet.write(0, 1, 'Student Name')
    worksheet.write(0, 2, 'Student Email')
    worksheet.write(0, 3, 'Student Year')
    worksheet.write(0, 4, 'Student Discipline')
    worksheet.write(0, 5, 'Student Option')
    worksheet.write(0, 6, 'Student Career Fields of Interest')
    worksheet.write(0, 7, 'Student Long Ans')
    worksheet.write(0, 8, 'Student interested in emailing alumni')
    worksheet.write(0, 9, 'Student Location')

    for row in range(len(rejected_students)):
        worksheet.write(row + 1, 0, row)
        worksheet.write(row+1, 1, rejected_students[row].name)
        worksheet.write(row+1, 2, rejected_students[row].email)
        worksheet.write(row+1, 3, rejected_students[row].year)
        worksheet.write(row+1, 4, rejected_students[row].discipline)
        worksheet.write(row+1, 5, rejected_students[row].option)
        worksheet.write(row+1, 6, rejected_students[row].career_fields)
        worksheet.write(row+1, 7, rejected_students[row].email_long_ans)
        worksheet.write(row+1, 8, rejected_students[row].alum_contact_flag)
        worksheet.write(row+1, 9, rejected_students[row].city)

    worksheet = workbook.add_sheet('Rejected Alumni')

    worksheet.write(0, 0, '#')
    worksheet.write(0, 1, 'Alum Name')
    worksheet.write(0, 2, 'Alum Email')
    worksheet.write(0, 3, 'Alum Discipline')
    worksheet.write(0, 4, 'Alum Job Title')
    worksheet.write(0, 5, 'Alum Company')
    worksheet.write(0, 6, 'Alum Career Fields')
    worksheet.write(0, 7, 'Alum Work Address')
    worksheet.write(0, 8, 'Alum Job Description')
    worksheet.write(0, 9, 'Alum Student Benefit Description')

    for row in range(len(rejected_alumni)):
        worksheet.write(row+1, 0, row)
        worksheet.write(row+1, 1, rejected_alumni[row].name)
        worksheet.write(row+1, 2, rejected_alumni[row].email)
        worksheet.write(row+1, 3, rejected_alumni[row].discipline)
        worksheet.write(row+1, 4, rejected_alumni[row].job_title)
        worksheet.write(row+1, 5, rejected_alumni[row].company)
        worksheet.write(row+1, 6, rejected_alumni[row].career_fields)
        worksheet.write(row+1, 7, rejected_alumni[row].city)
        worksheet.write(row+1, 8, rejected_alumni[row].tasks_long_ans)
        worksheet.write(row+1, 9, rejected_alumni[row].benefit_long_ans)

    workbook.save('job_shadowing_output_data.xls')

    # student_data = open(student_file_path).read().split(',')
    # alumni_data = open(student_file_path).read().split(',')

    # email_database = {}
    #
    # for entry in email_list_file:
    #     (key, val) = entry.split(' <')
    #     email_database[key] = val.replace('>', '')
    # print("email_database_clean: {}".format(email_database))
    #
    # # 2) read in the poll results into a list and clean up the strings
    #
    # poll_results_file = open(participants_file_path, 'r')
    #
    # # store names and their days into a list to iterate over
    # poll_results = [line.split('\t') for line in poll_results_file]
    # print("uncleaned_emails are: {}".format(poll_results))
    #
    # # iterate through and update the values of the list accordingly
    # # i.e. replace the name with their email and replace their day with "day1","day2","both days"
    #
    # weekdays = ['Mond', 'Tues', 'Wedn', 'Thur', 'Frid']
    #
    # # Note: the reason the loop is iterating in reversed order is because if a person did not respond to the poll, we
    # # would delete them off the list. It would mess with the iteration if we didn't reverse it
    # for entry_index in reversed(range(len(poll_results))):
    #
    #     # the first 4 letters of their selection. "Friday: DD/MM/YYYY ..." --> "Frid"
    #     selection = poll_results[entry_index][1][:4]
    #
    #     if selection == weekdays[day1]:
    #         poll_results[entry_index][1] = 'day1'
    #     elif selection == weekdays[day2]:
    #         poll_results[entry_index][1] = 'day2'
    #     elif selection == 'Eith':                   # corresponding to Either
    #         poll_results[entry_index][1] = 'both days'
    #     else:                                       # if someone did not respond to the poll, remove them from the list
    #         del poll_results[entry_index]
    #         continue
    #
    #     # now replace the person's name with their email found in the email dictionary we created
    #     try:
    #         poll_results[entry_index][0] = email_database[poll_results[entry_index][0].replace('\"', '')]
    #
    #     # if the email does't exist in the dictionary (highly unlikely), then input it in using the console
    #     except KeyError:
    #         print("Email for " + clean_emails[index][0] + " not found in database.")
    #         email_database[clean_emails[index][0]] = input("please input email: ")
    #         clean_emails[index][0] = email_database[clean_emails[index][0]]
    #
    # print("cleaned_emails are: {}".format(clean_emails))
    # print("length of list_of_emails is: {}".format(len(clean_emails)))



