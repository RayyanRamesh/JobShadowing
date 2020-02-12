import smtplib
import argparse
import datetime as dt
import pickle
import random

import create_groups as cg
import send_emails as se
import clean_data as cd

def _extract_program_args():
    global args
    parser = argparse.ArgumentParser(description='Handy helper function for matching alumni and students together for '
                                                 'the purposes of the Alumni Relation\'s Job Shadowing program.')
    parser.add_argument('-event_date1', default='2020-02-20', type=str, help='Date for first Job Shadowing day, as yyyy-mm-dd',
                        required=False)
    parser.add_argument('-event_date2', default='2020-02-21', type=str, help='Date for second Job Shadowing day, as yyyy-mm-dd',
                        required=False)
    parser.add_argument('-student_file_path', type=str,
                        help='The path of the excel file containing the Student responses for the Job Shadowing event', required=True)
    parser.add_argument('-alumni_file_path', type=str,
                        help='The path of the excel file containing the Alumni responses for the Job Shadowing event', required=True)
    parser.add_argument('-matches_file_path', type=str,
                        help='The path of the excel file containing the predetermined matches for the Job Shadowing event', required=True)
    # parser.add_argument('-mail_relay', default='mailrelay.cppib.ca', type=str,
    #                     help='Address of mail relay to send emails via')
    # parser.add_argument('-mail_body_message_file_path', type=str,
    #                     help='File containing a message that will be embedded in the body of the invite. '
    #                          'Formatted as HTML')
    # parser.add_argument('-invite_subject', default='Lunch Networking', type=str, help='Title to be used for invite')
    return parser.parse_args()

def main():
    args = _extract_program_args()

    '''
    Stage 1:
        clean_data.py
        1.0 make test data 
        1.1 read in student data and alumni data
        1.2 clean this data, place it in 2 lists made up of Student & Alumni class objects respectively
        1.3 make it so it uses arguments instead of hard coding paths
        
        for  cleaning data, need to have the following arguments: 2 paths to the excel files, 2 event dates
    '''
    create_matches_flag = True

    if create_matches_flag:
        student_list, alumni_list, good_matches, bad_matches = cd._read_data(args.student_file_path, args.alumni_file_path, args.matches_file_path)

        for i in range(len(student_list)):
            print(student_list[i].profile())

        for j in range(len(alumni_list)):
            print(alumni_list[j].profile())

        matches, rejected_alumni, rejected_students = cg._create_matches(alumni_list, student_list, good_matches, bad_matches)

        cd._save_data(matches, rejected_alumni, rejected_students)
        print("Final number of matches: {}". format(len(matches)))

    else:
        se._save_emails_to_file('job_shadowing_output_data.xls')
        #se._email_matches('job_shadowing_output_data.xls')



    # shuffled_groups = [[], []]
    # event_date = [[], []]
    #
    # args = _extract_program_args()
    #
    # event_date[0] = dt.datetime.strptime(args.event_date1, '%Y-%m-%d')
    # event_date[1] = dt.datetime.strptime(args.event_date2, '%Y-%m-%d')
    #
    # email_subject = args.invite_subject
    #
    # message_body = _retrieve_message_body(args.mail_body_message_file_path)
    #
    # first_day_number = event_date[0].weekday()
    # second_day_number = event_date[1].weekday()
    #
    # list_of_emails = _retrieve_emails(args.participants_file_path, args.email_database_path, first_day_number,
    #                                   second_day_number)
    #
    # shuffled_groups[0], shuffled_groups[1] = cg.get_lunch_groups(list_of_emails)
    #
    # # save the groups into a text file for future reference
    # file = open("C:/Users/lhalabi/Documents/lunchclub_demo/Groups_for_" + event_date[0].strftime('%Y-%m-%d') + "_and_" + event_date[1].strftime('%Y-%m-%d')
    #             + ".txt", 'w+')
    # for index in range(2):
    #     file.write('\n\nGroups for: ' + event_date[index].strftime('%Y-%m-%d')+'\n')
    #     for group_number in range(len(shuffled_groups[index])):
    #         file.write('\n'+ str(group_number+1) + '. \t')
    #         for member in shuffled_groups[index][group_number]:
    #             file.write(member + '\t')
    #
    # flag = False
    # testing = True
    # if flag:
    #     if testing:
    #         ce.Emailer(args.mail_relay).sendEmail(email_from="lhalabi@cppib.com", email_to=["lhalabi@cppib.com"],
    #                                            email_cc="lhalabi@cppib.com",
    #                                            subject=email_subject, html_body=message_body,
    #                                            eventYear=event_date[0].year, eventMonth=event_date[0].month,
    #                                            eventDay=event_date[0].day, eventHour=12, eventDuration=1, eventID=22222)
    #     else:
    #         # this is the email to the first day group
    #         ce._send_email_to_groups(ce.Emailer(args.mail_relay), email_subject, shuffled_groups[0], message_body, False,
    #                               event_date[0].day, event_date[0].month, event_date[0].year)
    #         # this is the email for the second day group, note that the event_date.day is incremented
    #         ce._send_email_to_groups(ce.Emailer(args.mail_relay), email_subject, shuffled_groups[1], message_body, False,
    #                               event_date[1].day, event_date[1].month, event_date[1].year)


if __name__ == '__main__':
    main()



    '''    
    create_groups.py
    2.0 initialize matching variables
    2.1 scramble the elements of each list 
    2.2 for each alum, find a student to match with, once their target is filled remove them from the list into a tuple with the student
    2.3 add all rejects into a rejected list 
    2.4 redo the process 100 times, until we end with minimal number of rejects
    2.5 print for approval on to the screen or into text file
    
    send_emails.py
    3.0 create a draft message for the email and substitute in the correct information for each person
    3.1 send out email or calendar invite to the student/alumni pairs 
    
    '''