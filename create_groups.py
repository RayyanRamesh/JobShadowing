'''
**Create Groups
**Laure Halabi
'''

# This class matches students with alumni

import random
GROUP_SIZE = 4


# class Match:
#     def __init__(self, alum, student1, student2=None):
#

DISCIPLINE_WEIGHT = 3
CAREER_WEIGHT = 2
MATCH_SCORE_THRESHOLD = 2

import copy

def _create_matches(alumni_list, student_list, good_matches_orginal, bad_matches):
    '''
    create_groups.py
    2.0 initialize matching variables
    2.1 scramble the elements of each list
    2.2 for each alum, find a student to match with, once their target is filled remove them from the list into a tuple with the student
    2.3 add all rejects into a rejected list
    2.4 redo the process 100 times, until we end with minimal number of rejects
    2.5 print for approval on to the screen
    '''

    iterations = 100
    best_matches = []
    double_flag = False

    for i in range(iterations):
        match_score_sum = 0
        curr_matches = []

        # Makes a copy of the lists of alumni, students, good matches and bad matches
        alumni = copy.deepcopy(alumni_list)
        students = copy.deepcopy(student_list)
        good_matches = copy.deepcopy(good_matches_orginal)

        random.shuffle(alumni)
        random.shuffle(students)
        # This array goes through the list of alumni and students and matches them based on the good and bad matches
        for alum in alumni:
            limit = alum.student_limit
            for student in students:
                if _valid_match(alum, student, bad_matches) and _get_match_score(alum, student) >= MATCH_SCORE_THRESHOLD:
                    if alum.name in good_matches:
                        student = [x for x in students if x.name == good_matches[alum.name]][0]
                        del good_matches[alum.name]
                    elif student.name in good_matches.values():  # dont let a student get matched elsewhere if they're a good match with an another alum
                        break
                        # When a match is found, the alumni's limit for number of students goes down by one
                    limit -= 1

                    if limit > 0:
                        hold_student = copy.deepcopy(student)
                        students.remove(student)

                        alum.avail_dates = list(set(alum.avail_dates) & set(student.avail_dates))
                        double_flag = True

                    if limit == 0:
                        match_score_sum += _get_match_score(alum, student)

                        if double_flag:
                            curr_matches.append([alum, student, hold_student])
                            alumni.remove(alum)
                            students.remove(student)
                            double_flag = False
                        else:
                            curr_matches.append([alum, student])
                            alumni.remove(alum)
                            students.remove(student)
                        break
            if double_flag:
                curr_matches.append([alum, hold_student])
                alumni.remove(alum)
                double_flag = False
        # each alumni and student match is based on a score. The higher the score the better matched the student and alumni are
        print("Average match score: {}".format(match_score_sum/len(curr_matches)))
        print("Number of Matches: {}".format(len(curr_matches)))

        if len(curr_matches) > len(best_matches):
            best_matches = curr_matches
            rejected_students = copy.deepcopy(students)
            rejected_alumni = copy.deepcopy(alumni)

    return best_matches, rejected_alumni, rejected_students     # Sends back the best matches, and alumni and students that didn't match well

# This function determines if a alumni and student are good matches based on their discipline, city, and available dates
def _valid_match(alum, student, bad_matches):
    # check that they haven't been considered a bad match
    if alum.name in bad_matches and student.name in bad_matches[alum.name]:
        return False
    # Check that they're in the same city
    if student.city != alum.city:
        return False
    # Check if they're in the same discipline
    if student.same_disp_flag or alum.same_disp_flag:
        if student.discipline != alum.discipline:
            return False
    # Check if they're available for the same dates
    for i in range(len(student.avail_dates)):
        for j in range(len(alum.avail_dates)):
            if student.avail_dates[i] == alum.avail_dates[j]:
                return True
    return False


def _get_match_score(alum, student):

    match_score = 0
    if student.discipline == alum.discipline:
        match_score += DISCIPLINE_WEIGHT

    if student.career_fields == ['Anything']:
        return MATCH_SCORE_THRESHOLD

    for i in range(len(student.career_fields)):
        for j in range(len(alum.career_fields)):
            if student.career_fields[i] == alum.career_fields[j]:
                match_score += CAREER_WEIGHT

    return match_score

# def get_lunch_groups(list_of_emails):
#     if len(list_of_emails) <= 6:
#         raise ValueError("Unfortunately, there are not enough people to participate.")
#         return 0, 0
#
#     email_list = [[], [], []]
#
#     # Place each participant in their respective list:
#     # 0 - first day, 1 - second day, 2 - either day
#     for email in list_of_emails:
#         email_list[email[1]].append(email[0])
#
#     # Fill the first and second day groups such that they each have a perfect number to form groups of 4
#     for list_index in range(2):
#         amount = (GROUP_SIZE - len(email_list[list_index])) % GROUP_SIZE
#         email_list[list_index].extend(email_list[2][:amount])
#         del email_list[2][:amount]
#
#     # With the remaining either day participants, remove excess groups of 4 and them to the first day participant lists
#     while len(email_list[2]) > GROUP_SIZE - 1:
#         email_list[0].extend(email_list[2][:GROUP_SIZE])
#         del email_list[2][0:GROUP_SIZE]
#
#     # Note that the first and second day groups now have a perfect number to form groups of the set amount,
#     # shuffle each list and section each off into smaller sub groups
#     for list_index in range(2):
#         random.shuffle(email_list[list_index])
#         email_list[list_index] = [email_list[list_index][x:x + GROUP_SIZE] for x in
#                                   range(0, len(email_list[list_index]), GROUP_SIZE)]
#
#     # there is at most 3 participants remaining in the either day list, iteratively add them into other sub groups,
#     # thus making groups of 5. In the case that there aren't enough groups, form a group of 6.
#     for count in range(len(email_list[2])):
#         if count < len(email_list[0]):
#             email_list[0][count].append(email_list[2][count])
#         elif (count-len(email_list[0])) < len(email_list[1]):
#             email_list[0][count-len(email_list[1])].append(email_list[2][count])
#         else:
#             print("Making group of 6")
#             email_list[0][count-len(email_list[0] + email_list[1])].append(email_list[2][count])
#
#     print("email_list[0]: {}".format(email_list[0]))
#     print("email_list[1]: {}".format(email_list[1]))
#     print("email_list[2]: {}".format(email_list[2]))
#     print("list of emails are: {}".format(list_of_emails))
#
#     return email_list[0], email_list[1]
