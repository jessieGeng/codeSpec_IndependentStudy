import csv
import os
import datetime as dt
import sys

# function to check if time1 is before time2
def is_before(time1, time2):
    time_f = "%Y-%m-%d %H:%M:%S"
    t1 = dt.datetime.strptime(time1,time_f)
    t2 = dt.datetime.strptime(time2,time_f)
    return t1 < t2


# function to return the time in seconds from two timestamps
def get_time_in_secs(t1_str, t2_str):

    # check if either string is None and return 0
    if t1_str == None or t2_str == None:
        return 0

    # calculate the time
    time_f = "%Y-%m-%d %H:%M:%S"
    t1 = dt.datetime.strptime(t1_str,time_f)
    t2 = dt.datetime.strptime(t2_str,time_f)
    total = t2 - t1
    return total.total_seconds()

# function to get the total time in seconds, but removing time spent on other
def get_total_time_in_seconds(user_info):
    time_f = "%Y-%m-%d %H:%M:%S"
    total_time_in_secs = 0

    # if an end time (never finished)
    if user_info['completed'] == True:

        # loop through time list getting the time in seconds for each
        for time_tuple in user_info['time_list']:
            start = time_tuple[0]
            end = time_tuple[1]
            total_time_in_secs = total_time_in_secs + get_time_in_secs(start,end)

        # remove dead time while solving
        total_time_in_secs = total_time_in_secs - user_info['dead_time']

        # return total time and dead time
        return (total_time_in_secs, user_info['dead_time'])

    # did not succesfully complete the problem
    else:
        return (0, 0)

# function to write the time to first correct solution for each user
def time_correct_Parsons(inFileName, outFileName, divid, poll_id, ignore):

    # set the field size to max
    csv.field_size_limit(sys.maxsize)

    # open the output file for writing
    dir = os.path.dirname(__file__)

    # open the input and output files as csv files
    with open(os.path.join(dir, inFileName)) as csv_file:

        # get csv_reader
        csv_reader = csv.reader(csv_file)

        # create an empty user dictionary
        user_dict = dict()

        # loop through the data
        for cols in csv_reader:

            # get the user and problem id
            user = cols[1]
            currid = cols[4]
            time = cols[0]
            event = cols[2]
            act = cols[3]
            gap_time = 5 * 60 # 5 mintues

            # if header continue
            if user == "sid":
                continue

            # check if user in dictionary
            elif user in user_dict:
                user_info = user_dict.get(user)

                # if aswering poll question after finishing save answer and date and time
                if currid == poll_id and user_info["completed"] == True:
                    value = act.split(':')[0]
                    user_info['poll_value'] = value
                    user_info['poll_time'] = time
                    user_info['num_answers'] = user_info['num_answers'] + 1

                # If user has not solved this problem
                if user_info["completed"] == False:

                    # if working on the same problem and the amount of time between
                    # the last and this event is greater than the gap time then
                    # add to the amount of dead time
                    if currid == divid and user_info['last_divid'] == divid:
                        time_between = get_time_in_secs(user_info['last_time'], time)
                        if time_between >= gap_time:
                            user_info['dead_time'] = user_info['dead_time']+ time_between

                    # if working on a different problem then set start and end on time list
                    if currid != divid and user_info['last_divid'] == divid:
                        time_between = get_time_in_secs(user_info['last_time'], time)
                        if time_between < gap_time:
                            user_info['time_list'].append((user_info['start_time'], time))
                        else:
                            if user_info['start_time'] != user_info['last_time']:
                                user_info['time_list'].append((user_info['start_time'], user_info['last_time']))

                    # if back to same problem after another then reset the start time
                    if currid == divid and user_info['last_divid'] != divid:
                        user_info['start_time'] = time

                    # if same problem, and correct, record end time, add one to tests
                    if currid == divid and event == 'parsons' and act.startswith("correct"):
                        user_info['end_time'] = time
                        user_info['completed'] = True
                        user_info['num_tests'] = user_info['num_tests'] + 1
                        user_info['time_list'].append((user_info['start_time'], time))

                    # else if same problem and not correct add one to tests
                    elif currid == divid and event == "parsons" and act.startswith("incorrect"):
                        user_info['num_tests'] = user_info['num_tests'] + 1

                    # else if same problem and combined blocks increment that count
                    elif currid == divid and event == "parsonsMove" and act.startswith("combined"):
                        user_info['combined'] = user_info['combined'] + 1

                    # else if same problem and removed blocks increment that count
                    elif currid == divid and event == "parsonsMove" and act.startswith("removedDistractor"):
                        user_info['removed'] = user_info['removed'] + 1

                    # else if same problem and removed blocks increment that count
                    elif currid == divid and event == "parsonsMove" and act.startswith("removedIndentation"):
                        user_info['indented'] = user_info['indented'] + 1

                    # else if same problem and removed blocks increment that count
                    elif currid == divid and event == "parsonsMove" and act.startswith("reset"):
                        user_info['reset'] = user_info['reset'] + 1

                # else if solving again check if time is earlier than original start time and if is reset
                elif divid == currid and event == "parsons":
                    if is_before(time, user_info["start_time"]):
                        user_dict[user] = {'completed': False, 'start_time': time, 'end_time': None, 'time_list': [], 'last_divid': divid, 'other_divids': [],'last_time': None, 'dead_time' : 0, 'num_tests': 0, 'combined': 0, 'indented': 0, 'removed': 0, 'reset': 0, 'poll_value': None, 'poll_time': None, 'num_answers' : 0}

            # not in dictionary and parsonsMove so record the start time
            elif user not in user_dict and divid == currid and event == "parsonsMove" and act.startswith("start"):
                user_dict[user] = {'completed': False, 'start_time': time, 'end_time': None, 'time_list': [], 'last_divid': divid, 'other_divids': [], 'last_time': None, 'dead_time' : 0, 'num_tests': 0, 'combined': 0, 'indented': 0, 'removed': 0, 'reset': 0, 'poll_value': None, 'poll_time': None, 'num_answers': 0}


            # get user_info
            user_info = user_dict.get(user)

            if user_info != None:

                # reset last time to this one
                user_info['last_time'] = time

                # reset the last_divid
                user_info['last_divid'] = currid

                # if not completed add to list of other divids if not already there
                if user_info["completed"] == False and currid != divid and currid not in user_info['other_divids']:
                    user_info['other_divids'].append(currid)


    # print the number of people that attemptd the problem
    print(f"The number of people that attempted the problem is {len(user_dict.keys())}")
    not_solved_users = []

    # open the output file
    with open(os.path.join(dir, outFileName), "w") as outFile:
        csv_writer = csv.writer(outFile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csv_writer.writerow(['user_id', 'total_time', 'dead_time', 'num_tests', 'time_lists', 'start_time', 'end_time', 'combined', 'indented', 'removed', 'reset', 'other_divids', 'poll_value', 'poll_time', 'num_answers'])

        # loop through user dictionary
        count_not_correct = 0
        for user in user_dict.keys():
            user_info = user_dict[user]
            (total_time, other_time) = get_total_time_in_seconds(user_info)
            if total_time == 0 and user not in ignore:
                count_not_correct += 1
                not_solved_users.append(user)
            else:
                if user not in ignore:
                    csv_writer.writerow([user, total_time, user_info['dead_time'], user_info['num_tests'], user_info['time_list'], user_info['start_time'], user_info['end_time'], user_info['combined'], user_info['indented'], user_info['removed'], user_info['reset'], user_info['other_divids'], user_info["poll_value"], user_info["poll_time"], user_info['num_answers']])
        print(f"Number never solved {divid} is {count_not_correct}")
        print(not_solved_users)

users_to_ignore = [create a list of people to ignore if needed]
time_correct_Parsons([replace with input file name], [replace with output file name], [replace with problem], [replace with cognitive load poll name], users_to_ignore)