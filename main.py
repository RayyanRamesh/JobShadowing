import argparse

import create_groups as cg
import send_emails as se
import clean_data as cd

#TODO: need to include sample parameters in README file
def _extract_program_args():
    '''
        Arguments feed into parameters text box in the Run -> Edit Configurations tab will be properly parsed by this function.
        Arguments are all saved into the variable args and can be accessed using the dot operator.
        See ReadMe for example parameters.

    :return: None
    '''
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
    return parser.parse_args()

def main():
    '''
         If we are creating matches, then read in student data and alumni data from the excel files.
         Print all the profiles to the terminal to review. Call create_matches the matches to create matches and return
         three lists: matches, rejected_alumni, rejected_students. Finally, save the three lists into an excel file
         seperated by sheets.

         If matches are already made, read the aforementioned excel file and create list of emails to send out.

         Note: You may be able to send emails out after creating the match list, however it requires a stmp connection
         and some troubleshooting.

    :return: None
    '''

    args = _extract_program_args()
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


if __name__ == '__main__':
    main()

