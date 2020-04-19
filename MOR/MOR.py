import pandas as pd
import numpy as np
import xlrd
import copy
import pulp as p

days = ["Monday", 'Tuesday', 'Wednesday', 'Thursday']
# Teams = ['TeamA', 'TeamB', 'TeamC', 'TeamD', 'TeamE', 'TeamF', 'TeamG']
# Teams = [Teams,TeamB_Emp, TeamC_Emp, TeamD_Emp, TeamE_Emp, TeamF_Emp, TeamG_Emp]
TeamA_Emp = [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1012, 1013, 1014, 1015]
TeamB_Emp = [2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017]
TeamC_Emp = [3001, 3002, 3003, 3004, 3005, 3006, 3007, 3008, 3009, 3010, 3011, 3012, 3013, 3014, 3015, 3016, 3017]
TeamD_Emp = [4001, 4002, 4003, 4004, 4005, 4006, 4007, 4008, 4009, 4010, 4011]
TeamE_Emp = [5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009, 5010, 5011]
TeamF_Emp = [6001, 6002, 6003, 6004, 6005, 6006, 6007, 6008, 6009, 6010, 6011, 6012, 6013, 6014, 6015]
TeamG_Emp = [7001, 7002, 7003, 7004, 7005, 7006, 7007, 7008, 7009, 7010, 7011, 7012, 7013, 7014]

Teams = [TeamA_Emp, TeamB_Emp, TeamC_Emp, TeamD_Emp, TeamE_Emp, TeamF_Emp, TeamG_Emp]

TeamA_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
TeamB_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
TeamC_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
TeamD_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9]
TeamE_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9]
TeamF_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
TeamG_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]

Stations = [TeamA_Stations, TeamB_Stations, TeamC_Stations, TeamD_Stations, TeamE_Stations, TeamF_Stations,
            TeamG_Stations]

SkillsA = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamA_Skills', index_col=0)
# print(SkillsA)

SkillsB = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamB_Skills', index_col=0)
# print(SkillsB)
SkillsC = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamC_Skills', index_col=0)
# a = pd.DataFrame(SkillsC)
# print(SkillsC)
SkillsD = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamD_Skills', index_col=0)
# print(SkillsD)
SkillsE = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamE_Skills', index_col=0)
# print(SkillsE)
SkillsF = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamF_Skills', index_col=0)
# print(SkillsF)
SkillsG = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                        r'Research\Course_Project\Default_Project\all_teams_skills.xlsx', 'TeamG_Skills', index_col=0)
# print(SkillsG)


Skills_set = [SkillsA, SkillsB, SkillsC, SkillsD, SkillsE, SkillsF, SkillsG]
# print(SkillsA['A1'])

PreferenceA = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamA', index_col=0)
PreferenceB = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamB', index_col=0)
PreferenceC = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamC', index_col=0)
PreferenceD = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamD', index_col=0)
PreferenceE = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamE', index_col=0)
PreferenceF = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamF', index_col=0)
PreferenceG = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                            r'Research\Course_Project\Default_Project\all_teams_prefer.xlsx', 'TeamG', index_col=0)

Preference = [PreferenceA, PreferenceB, PreferenceC, PreferenceD, PreferenceE, PreferenceF, PreferenceG]
# print(PreferenceA)


H_Attendance_A = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamA', index_col=0)
H_Attendance_B = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamB', index_col=0)
H_Attendance_C = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamC', index_col=0)
H_Attendance_D = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamD', index_col=0)
H_Attendance_E = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamE', index_col=0)
H_Attendance_F = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamF', index_col=0)
H_Attendance_G = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                               r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'TeamG', index_col=0)

# sprint(H_Attendance_A)
H_Attendance = [H_Attendance_A, H_Attendance_B, H_Attendance_C, H_Attendance_D, H_Attendance_E, H_Attendance_F,
                H_Attendance_G]
FullCost = 320
GoHome_Cost = 160
NoCost = 0

Current_AttendanceA = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamA',
                                    index_col=0)
Current_AttendanceB = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamB',
                                    index_col=0)
Current_AttendanceC = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamC',
                                    index_col=0)
Current_AttendanceD = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamD',
                                    index_col=0)
Current_AttendanceE = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamE',
                                    index_col=0)
Current_AttendanceF = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamF',
                                    index_col=0)
Current_AttendanceG = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                    r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New_TeamG',
                                    index_col=0)

Current_Attendance_Sheets = [Current_AttendanceA, Current_AttendanceB, Current_AttendanceC, Current_AttendanceD,
                             Current_AttendanceE, Current_AttendanceF, Current_AttendanceG]
# print(Current_Attendance)
Work_Sheet_A = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetA', index_col=0)
Work_Sheet_B = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetB', index_col=0)
Work_Sheet_C = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetC', index_col=0)
Work_Sheet_D = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetD', index_col=0)
Work_Sheet_E = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetE', index_col=0)
Work_Sheet_F = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetF', index_col=0)
Work_Sheet_G = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                             r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheetG', index_col=0)

Work_Sheets = [Work_Sheet_A, Work_Sheet_B, Work_Sheet_C, Work_Sheet_D, Work_Sheet_E, Work_Sheet_F, Work_Sheet_G]

Employee_Assigned_A = {1001: 0, 1002: 0, 1003: 0, 1004: 0, 1005: 0, 1006: 0, 1007: 0, 1008: 0, 1009: 0, 1010: 0,
                       1011: 0, 1012: 0, 1013: 0, 1014: 0, 1015: 0}
Station_Assigned_A = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 13: 0}
Attendance_Sheet_A = {1001: 0, 1002: 0, 1003: 0, 1004: 0, 1005: 0, 1006: 0, 1007: 0, 1008: 0, 1009: 0, 1010: 0, 1011: 0,
                      1012: 0, 1013: 0, 1014: 0, 1015: 0}

Employee_Assigned_B = {2001: 0, 2002: 0, 2003: 0, 2004: 0, 2005: 0, 2006: 0, 2007: 0, 2008: 0, 2009: 0, 2010: 0,
                       2011: 0, 2012: 0, 2013: 0, 2014: 0, 2015: 0, 2016: 0, 2017: 0}
Station_Assigned_B = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 13: 0}
Attendance_Sheet_B = {2001: 0, 2002: 0, 2003: 0, 2004: 0, 2005: 0, 2006: 0, 2007: 0, 2008: 0, 2009: 0, 2010: 0, 2011: 0,
                      2012: 0, 2013: 0, 2014: 0, 2015: 0, 2016: 0, 2017: 0}

Employee_Assigned_C = {3001: 0, 3002: 0, 3003: 0, 3004: 0, 3005: 0, 3006: 0, 3007: 0, 3008: 0, 3009: 0, 3010: 0,
                       3011: 0, 3012: 0, 3013: 0, 3014: 0, 3015: 0, 3016: 0, 3017: 0}
Station_Assigned_C = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0}
Attendance_Sheet_C = {3001: 0, 3002: 0, 3003: 0, 3004: 0, 3005: 0, 3006: 0, 3007: 0, 3008: 0, 3009: 0, 3010: 0,
                      3011: 0, 3012: 0, 3013: 0, 3014: 0, 3015: 0, 3016: 0, 3017: 0}

Employee_Assigned_D = {4001: 0, 4002: 0, 4003: 0, 4004: 0, 4005: 0, 4006: 0, 4007: 0, 4008: 0, 4009: 0, 4010: 0,
                       4011: 0}
Station_Assigned_D = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0}
Attendance_Sheet_D = {4001: 0, 4002: 0, 4003: 0, 4004: 0, 4005: 0, 4006: 0, 4007: 0, 4008: 0, 4009: 0, 4010: 0, 4011: 0}

Employee_Assigned_E = {5001: 0, 5002: 0, 5003: 0, 5004: 0, 5005: 0, 5006: 0, 5007: 0, 5008: 0, 5009: 0, 5010: 0,
                       5011: 0}
Station_Assigned_E = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0}
Attendance_Sheet_E = {5001: 0, 5002: 0, 5003: 0, 5004: 0, 5005: 0, 5006: 0, 5007: 0, 5008: 0, 5009: 0, 5010: 0, 5011: 0}

Employee_Assigned_F = {6001: 0, 6002: 0, 6003: 0, 6004: 0, 6005: 0, 6006: 0, 6007: 0, 6008: 0, 6009: 0, 6010: 0,
                       6011: 0, 6012: 0, 6013: 0, 6014: 0, 6015: 0}
Station_Assigned_F = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0}
Attendance_Sheet_F = {6001: 0, 6002: 0, 6003: 0, 6004: 0, 6005: 0, 6006: 0, 6007: 0, 6008: 0, 6009: 0, 6010: 0, 6011: 0,
                      6012: 0, 6013: 0, 6014: 0, 6015: 0}

Employee_Assigned_G = {7001: 0, 7002: 0, 7003: 0, 7004: 0, 7005: 0, 7006: 0, 7007: 0, 7008: 0, 7009: 0, 7010: 0,
                       7011: 0, 7012: 0, 7013: 0, 7014: 0}
Station_Assigned_G = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0}
Attendance_Sheet_G = {7001: 0, 7002: 0, 7003: 0, 7004: 0, 7005: 0, 7006: 0, 7007: 0, 7008: 0, 7009: 0, 7010: 0, 7011: 0,
                      7012: 0, 7013: 0, 7014: 0}

Employee_Assigned_list = [Employee_Assigned_A, Employee_Assigned_B, Employee_Assigned_C, Employee_Assigned_D,
                          Employee_Assigned_E, Employee_Assigned_F, Employee_Assigned_G]
Station_Assigned_list = [Station_Assigned_A, Station_Assigned_B, Station_Assigned_C, Station_Assigned_D,
                         Station_Assigned_E, Station_Assigned_F, Station_Assigned_G]
Attendance_Sheet_List = [Attendance_Sheet_A, Attendance_Sheet_B, Attendance_Sheet_C, Attendance_Sheet_D,
                         Attendance_Sheet_E, Attendance_Sheet_F, Attendance_Sheet_G]


def assign_stations2():
    Overall_Cost = 0
    overall_no_of_employees_working = 0
    overall_no_of_stations_operating = 0
    overall_unassigned_employees = []
    overall_employees_sent_home = []
    overall_employees_cross_trained = []
    no_of_employees_cross_trained = 0
    preference_count = 0
    employees_assigned_according_to_preference = []
    day_count = 0
    employees_under_cross_training = []
    for day in days:
        day_count += 1
        print()
        print('Schedule of Day: {}'.format(day))
        print()
        print('Day Count {}'.format(day_count))
        for i in range(0, 7):
            unassigned_employees = []
            Total_Cost = 0
            no_of_employees_working = 0
            no_of_stations_operating = 0
            Team = Teams[i]
            Skills = Skills_set[i]
            Station_Assigned = Station_Assigned_list[i]
            Employee_Assigned = Employee_Assigned_list[i]
            Attendance_Sheet = Attendance_Sheet_List[i]
            Preference_Sheet = Preference[i]
            Team_Station = Stations[i]
            Work_Sheet = Work_Sheets[i]
            Current_Attendance = Current_Attendance_Sheets[i]
            unassigned_employees = copy.deepcopy(Employee_Assigned_list[i])
            for employee in Team:
                if Employee_Assigned[employee] <= day_count != 0 and employee in employees_under_cross_training:
                    employees_under_cross_training.remove(employee)

            for Emp in Team:
                sum_row = (Skills.sum(axis=1))
                if sum_row[Emp] == 1:
                    for ws in Team_Station:
                        if Skills[ws][Emp] == 1 and Station_Assigned[ws] < day_count and Employee_Assigned[Emp] < day_count and Emp not in employees_under_cross_training and (Emp != max(Team)) and (Emp != (max(Team) - 1)):
                            # print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] = day_count
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            unassigned_employees.pop(Emp)
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            continue
                else:
                    for ws in Team_Station:
                        if Skills[ws][Emp] == 1 and Employee_Assigned[Emp] < day_count and Station_Assigned[ws] < day_count and (Emp != max(Team)) and (Emp != (max(Team) - 1)) and Emp not in employees_under_cross_training and Emp in unassigned_employees:
                            # print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] = day_count
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            unassigned_employees.pop(Emp)
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            continue

            if int(no_of_employees_working) < len(Team) and (len(Team) - no_of_employees_working - 2) >= 2:
                list = []
                station = []
                for ws in Team_Station:
                    for Emp in Team:
                        if Skills[ws][Emp] == 0 and Employee_Assigned[Emp] < day_count and Station_Assigned[ws] < day_count and (Emp != max(Team)) and (Emp != (max(Team) - 1)):
                            list.append(Emp)
                            station.append(ws)
                        if len(list) == 2 and station[0] == station[1]:
                            print(
                                'Employees {}, {} have been assigned station {}{} as they are unskilled'.format(list[0], list[1], chr(i + 65), station[0]))
                            no_of_employees_working += 2
                            no_of_stations_operating += 1
                            Current_Attendance[day][list[0]] += 1
                            Current_Attendance[day][list[1]] += 1
                            Total_Cost += (2 * 320)
                            Employee_Assigned[list[0]] = day_count
                            Employee_Assigned[list[1]] = day_count
                            Station_Assigned[station[0]] += 1
                            unassigned_employees.pop(list[0])
                            unassigned_employees.pop(list[1])
                            Work_Sheet[station[0]][list[0]] = 1
                            Work_Sheet[station[1]][list[1]] = 1
                            station.pop(0)
                            station.pop(0)
                            list.pop(0)
                            list.pop(0)
                            break
                        continue

            if int(no_of_employees_working) < len(Team):
                for ws in Station_Assigned:
                    if Station_Assigned[ws] == 0 and Employee_Assigned[max(Team)] < day_count:
                        no_of_employees_working += 1
                        no_of_stations_operating += 1
                        Current_Attendance[day][max(Team)] += 1
                        Station_Assigned[ws] += 1
                        # Employee_Assigned[(max(Team))] += 1
                        # Employee_Assigned[(max(Team) - 1)] += 1
                        Attendance_Sheet[(max(Team) - 1)] += 1
                        Attendance_Sheet[max(Team)] += 1
                        Current_Attendance[day][max(Team) - 1] += 1
                        Total_Cost += FullCost
                        unassigned_employees.pop(max(Team))
                        unassigned_employees.pop((max(Team) - 1))
                        print()
                        print(
                            'Tag Relief {} has been assigned to workstation {}{}'.format((max(Team)), chr(i + 65), ws))
                        print('Employee {} is the production Leader'.format((max(Team) - 1)))
                        break
                    continue

            # if Employee_Assigned[max(Team)] < 1 and Employee_Assigned[(max(Team) - 1)] < 1:
            if Current_Attendance[day][max(Team)] < 1 and Current_Attendance[day][(max(Team) - 1)] < 1:
                print('Employee {} is Tag Relief'.format(max(Team)))
                print('Employee {} is a production Leader'.format((max(Team) - 1)))
                # Employee_Assigned[(max(Team) - 1)] = day_count
                # Employee_Assigned[(max(Team))] = day_count
                Current_Attendance[day][max(Team)] += 1
                Current_Attendance[day][(max(Team) - 1)] += 1

                unassigned_employees.pop((max(Team) - 1))
                unassigned_employees.pop(max(Team))

            # print('Remaining Employees')
            for e in unassigned_employees:
                if Employee_Assigned[e] >= day_count:
                    unassigned_employees.pop(e)
                    break

            if len(unassigned_employees) == 0:
                print()
                print('All assignments done, no cross training required')
                print('No Remaining Employees')
            else:
                mylist = []
                for emp in unassigned_employees.keys():
                    mylist.append(emp)
                print('Employee {} hasn\'t been assigned any workstation'.format(mylist))

                input_task = input('Enter 1 to cross train employees, else 2 to send home')
                if int(input_task) == 1:
                    Total_Cost += FullCost
                    emp = input('Enter the employee you wish to cross train')
                    emp = int(emp)
                    Attendance_Sheet[emp] = day_count
                    # Employee_Assigned[emp] += 1
                    # Employee_Assigned[emp] += 1
                    # Current_Attendance[emp] += 1
                    # Current_Attendance[emp] += 1
                    Current_Attendance[day][emp] += 1
                    no_of_employees_cross_trained += 1
                    overall_employees_cross_trained.append(emp)
                    unassigned_employees.pop(int(emp))

                    for e in unassigned_employees:
                        Attendance_Sheet[emp] += 0
                        Total_Cost += GoHome_Cost
                    list_of_station_options = []
                    for ws in Team_Station:
                        if Skills[ws][emp] == 0:
                            list_of_station_options.append(ws)
                            continue
                    if len(list_of_station_options) == 0:
                        print('Employee {} is trained in all stations'.format(emp))
                    else:
                        station = input(
                            'Enter station from the following to train upon {}'.format(list_of_station_options))
                        if int(station) in list_of_station_options:
                            Skills[int(station)][emp] = 1
                            Attendance_Sheet[emp] += 1
                            Attendance_Sheet[emp] += 1
                            Employee_Assigned[emp] += 1
                            Employee_Assigned[emp] += 1
                            employees_under_cross_training.append(emp)

                        if len(unassigned_employees) != 0:
                            # for emp in Team:
                            for employee in unassigned_employees:
                                if Employee_Assigned[employee] == 0:
                                    Total_Cost += GoHome_Cost
                                    Attendance_Sheet[employee] += 0
                                    Current_Attendance[day][employee] += 0
                                    print('Employee {} is being sent Home'.format(employee))
                                    overall_employees_sent_home.append(employee)
                                    continue

                elif int(input_task) == 2:
                    for emp in Team:
                        if Employee_Assigned[emp] == 0 and (emp != max(Team)) and (emp != (max(Team) - 1)):
                            Total_Cost += GoHome_Cost
                            Attendance_Sheet[emp] += 0
                            print('Employee {} is being sent home'.format(unassigned_employees))
                            overall_employees_sent_home.append(emp)

            # break

            print()
            print('No. of employees present at work: {}'.format(no_of_employees_working))
            print('No. of workstations running {}'.format(no_of_stations_operating))
            print()
            print('Total Cost is {}$'.format(Total_Cost))
            print()
            print('Assignment Matrix of Team{}'.format(chr(65 + i)))
            print()
            print(Work_Sheet)
            print()
            print('New Skills Sheet for Team{}'.format(chr((65 + i))))
            print()
            print(Skills)
            # if i == 0:
            Skills_set[i] = Skills
            print('Attendance Sheet for Team{}'.format(chr(65 + i)))
            print()
            print(Current_Attendance_Sheets[i])
            Overall_Cost += Total_Cost
            overall_no_of_employees_working += no_of_employees_working
            overall_no_of_stations_operating += no_of_stations_operating
            print()
            print('Preference Matrix Solving')
            print()
            a = (preference_maximize(Work_Sheet, Skills, Preference_Sheet))

            continue

    print()
    print('Total Cost {}'.format(Overall_Cost))
    print()
    if len(overall_unassigned_employees) != 0:
        print('Overall Unassigned Employees {}'.format(overall_unassigned_employees))
    print()
    print('Overall No. of working Employees: {}'.format(overall_no_of_employees_working))
    print()
    print('Overall No. of stations Operating : {}'.format(overall_no_of_stations_operating))
    print()
    print('Employees sent Home {}'.format(overall_employees_sent_home))
    print()
    print('No. of Employees Cross Trained {}'.format(no_of_employees_cross_trained))
    print()
    print('Employees Cross Trained are: {}'.format(overall_employees_cross_trained))
    print()
    print('No. of Employees assigned according to Manager\'s preference: {}'.format(preference_count))
    print()
    print(
        'Employees assigned according to manager\'s preference: {}'.format(employees_assigned_according_to_preference))
    print()


def assign_stations3():
    Overall_Cost = 0
    overall_no_of_employees_working = 0
    overall_no_of_stations_operating = 0
    overall_unassigned_employees = []
    overall_employees_sent_home = []
    overall_employees_cross_trained = []
    no_of_employees_cross_trained = 0
    preference_count = 0
    employees_assigned_according_to_preference = []
    day_count = 0
    employees_under_cross_training = []
    for day in days:
        day_count += 1
        print()
        print('Schedule of Day: {}'.format(day))
        print()
        print('Day Count {}'.format(day_count))
        for i in range(0, 7):
            unassigned_employees = []
            Total_Cost = 0
            no_of_employees_working = 0
            no_of_stations_operating = 0
            Team = Teams[i]
            Skills = Skills_set[i]
            Skills2 = Skills_set[i]
            Station_Assigned = Station_Assigned_list[i]
            Employee_Assigned = Employee_Assigned_list[i]
            Attendance_Sheet = Attendance_Sheet_List[i]
            Preference_Sheet = Preference[i]
            Team_Station = Stations[i]
            Work_Sheet = Work_Sheets[i]
            Work_Sheet2 = Work_Sheets[i]
            Current_Attendance = Current_Attendance_Sheets[i]
            unassigned_employees = copy.deepcopy(Employee_Assigned_list[i])
            for employee in Team:
                if Employee_Assigned[employee] <= day_count != 0 and employee in employees_under_cross_training:
                    employees_under_cross_training.remove(employee)

            for ws in Team_Station:
                col_sum = (Skills.sum(axis=0))
                if col_sum[ws] == 1:
                    for Emp in Team:
                        if Skills[ws][Emp] == 1 and Station_Assigned[ws] < day_count and Employee_Assigned[Emp] < day_count and Emp not in employees_under_cross_training and (Emp != max(Team)) and (Emp != (max(Team) - 1)):
                            # print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] = day_count
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            unassigned_employees.pop(Emp)
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            continue
                else:
                    for Emp in Team:
                        if Skills[ws][Emp] == 1 and Employee_Assigned[Emp] < day_count and Station_Assigned[ws] < day_count and (Emp != max(Team)) and (Emp != (max(Team) - 1)) and Emp not in employees_under_cross_training and Emp in unassigned_employees:
                            # print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] = day_count
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            unassigned_employees.pop(Emp)
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation {}{}'.format(Emp, chr(i + 65), ws))
                            continue

            if int(no_of_employees_working) < len(Team) and (len(Team) - no_of_employees_working - 2) >= 2:
                list = []
                station = []
                for ws in Team_Station:
                    for Emp in Team:
                        if Skills[ws][Emp] == 0 and Employee_Assigned[Emp] < day_count and Station_Assigned[ws] < day_count and (Emp != max(Team)) and (Emp != (max(Team) - 1)):
                            list.append(Emp)
                            station.append(ws)
                        if len(list) == 2 and station[0] == station[1]:
                            print(
                                'Employees {}, {} have been assigned station {}{} as they are unskilled'.format(list[0], list[1], chr(i + 65), station[0]))
                            no_of_employees_working += 2
                            no_of_stations_operating += 1
                            Current_Attendance[day][list[0]] += 1
                            Current_Attendance[day][list[1]] += 1
                            Total_Cost += (2 * 320)
                            Employee_Assigned[list[0]] = day_count
                            Employee_Assigned[list[1]] = day_count
                            Station_Assigned[station[0]] += 1
                            unassigned_employees.pop(list[0])
                            unassigned_employees.pop(list[1])
                            Work_Sheet[station[0]][list[0]] = 1
                            Work_Sheet[station[1]][list[1]] = 1
                            station.pop(0)
                            station.pop(0)
                            list.pop(0)
                            list.pop(0)
                            break
                        continue

            if int(no_of_employees_working) < len(Team):
                for ws in Station_Assigned:
                    if Station_Assigned[ws] == 0 and Employee_Assigned[max(Team)] < day_count:
                        no_of_employees_working += 1
                        no_of_stations_operating += 1
                        Current_Attendance[day][max(Team)] += 1
                        Station_Assigned[ws] += 1
                        # Employee_Assigned[(max(Team))] += 1
                        # Employee_Assigned[(max(Team) - 1)] += 1
                        Attendance_Sheet[(max(Team) - 1)] += 1
                        Attendance_Sheet[max(Team)] += 1
                        Current_Attendance[day][max(Team) - 1] += 1
                        Total_Cost += FullCost
                        unassigned_employees.pop(max(Team))
                        unassigned_employees.pop((max(Team) - 1))
                        print()
                        print(
                            'Tag Relief {} has been assigned to workstation {}{}'.format((max(Team)), chr(i + 65), ws))
                        print('Employee {} is the production Leader'.format((max(Team) - 1)))
                        break
                    continue

            # if Employee_Assigned[max(Team)] < 1 and Employee_Assigned[(max(Team) - 1)] < 1:
            if Current_Attendance[day][max(Team)] < 1 and Current_Attendance[day][(max(Team) - 1)] < 1:
                print('Employee {} is Tag Relief'.format(max(Team)))
                print('Employee {} is a production Leader'.format((max(Team) - 1)))
                # Employee_Assigned[(max(Team) - 1)] = day_count
                # Employee_Assigned[(max(Team))] = day_count
                Current_Attendance[day][max(Team)] += 1
                Current_Attendance[day][(max(Team) - 1)] += 1

                unassigned_employees.pop((max(Team) - 1))
                unassigned_employees.pop(max(Team))

            # print('Remaining Employees')
            for e in unassigned_employees:
                if Employee_Assigned[e] >= day_count:
                    unassigned_employees.pop(e)
                    break

            if len(unassigned_employees) == 0:
                print()
                print('All assignments done, no cross training required')
                print('No Remaining Employees')
            else:
                mylist = []
                for emp in unassigned_employees.keys():
                    mylist.append(emp)
                print('Employee {} hasn\'t been assigned any workstation'.format(mylist))

                input_task = input('Enter 1 to cross train employees, else 2 to send home')
                if int(input_task) == 1:
                    Total_Cost += FullCost
                    emp = input('Enter the employee you wish to cross train')
                    emp = int(emp)
                    Attendance_Sheet[emp] = day_count
                    # Employee_Assigned[emp] += 1
                    # Employee_Assigned[emp] += 1
                    # Current_Attendance[emp] += 1
                    # Current_Attendance[emp] += 1
                    Current_Attendance[day][emp] += 1
                    no_of_employees_cross_trained += 1
                    overall_employees_cross_trained.append(emp)
                    unassigned_employees.pop(int(emp))

                    for e in unassigned_employees:
                        Attendance_Sheet[emp] += 0
                        Total_Cost += GoHome_Cost
                    list_of_station_options = []
                    for ws in Team_Station:
                        if Skills[ws][emp] == 0:
                            list_of_station_options.append(ws)
                            continue
                    if len(list_of_station_options) == 0:
                        print('Employee {} is trained in all stations'.format(emp))
                    else:
                        station = input(
                            'Enter station from the following to train upon {}'.format(list_of_station_options))
                        if int(station) in list_of_station_options:
                            Skills[int(station)][emp] = 1
                            Attendance_Sheet[emp] += 1
                            Attendance_Sheet[emp] += 1
                            Employee_Assigned[emp] += 1
                            Employee_Assigned[emp] += 1
                            employees_under_cross_training.append(emp)

                        if len(unassigned_employees) != 0:
                            # for emp in Team:
                            for employee in unassigned_employees:
                                if Employee_Assigned[employee] == 0:
                                    Total_Cost += GoHome_Cost
                                    Attendance_Sheet[employee] += 0
                                    Current_Attendance[day][employee] += 0
                                    print('Employee {} is being sent Home'.format(employee))
                                    overall_employees_sent_home.append(employee)
                                    continue

                elif int(input_task) == 2:
                    for emp in Team:
                        if Employee_Assigned[emp] == 0 and (emp != max(Team)) and (emp != (max(Team) - 1)):
                            Total_Cost += GoHome_Cost
                            Attendance_Sheet[emp] += 0
                            print('Employee {} is being sent home'.format(unassigned_employees))
                            overall_employees_sent_home.append(emp)

            # break

            print()
            print('No. of employees present at work: {}'.format(no_of_employees_working))
            print('No. of workstations running {}'.format(no_of_stations_operating))
            print()
            print('Total Cost is {}$'.format(Total_Cost))
            print()
            print('Assignment Matrix of Team{}'.format(chr(65 + i)))
            print()
            print(Work_Sheet)
            print()
            print('New Skills Sheet for Team{}'.format(chr((65 + i))))
            print()
            print(Skills)
            # if i == 0:
            Skills_set[i] = Skills
            print('Attendance Sheet for Team{}'.format(chr(65 + i)))
            print()
            print(Current_Attendance_Sheets[i])
            Overall_Cost += Total_Cost
            overall_no_of_employees_working += no_of_employees_working
            overall_no_of_stations_operating += no_of_stations_operating
            print()
            print('Preference Matrix Solving')
            print()
            a = (preference_maximize(Work_Sheet2, Skills2, Preference_Sheet))

            continue

    print()
    print('Total Cost {}'.format(Overall_Cost))
    print()
    if len(overall_unassigned_employees) != 0:
        print('Overall Unassigned Employees {}'.format(overall_unassigned_employees))
    print()
    print('Overall No. of working Employees: {}'.format(overall_no_of_employees_working))
    print()
    print('Overall No. of stations Operating : {}'.format(overall_no_of_stations_operating))
    print()
    print('Employees sent Home {}'.format(overall_employees_sent_home))
    print()
    print('No. of Employees Cross Trained {}'.format(no_of_employees_cross_trained))
    print()
    print('Employees Cross Trained are: {}'.format(overall_employees_cross_trained))
    print()
    print('No. of Employees assigned according to Manager\'s preference: {}'.format(preference_count))
    print()
    print(
        'Employees assigned according to manager\'s preference: {}'.format(employees_assigned_according_to_preference))
    print()


def preference_maximize(Worksheet, skills, preference, unit_cost=FullCost):
    print('Worksheet before optimization and Assignments are done')
    # print(Worksheet)
    people_present = Worksheet.index.to_list()

    preference_based_assignments = Worksheet.sum().sum()

    skills_present = skills[skills.index.isin(people_present)]

    skills_present = skills_present.applymap(lambda x: 0 if x == 0 else x)

    workstations = skills.columns.to_list()

    Employees = skills.index.to_list()

    problem = p.LpProblem("Preference", p.LpMaximize)
    Assignments = p.LpVariable.dicts('P', ((ws, emp) for ws in workstations for emp in Employees),
                                     cat='Binary')

    problem += p.lpSum([Assignments[(ws, Emp)] * preference.loc[Emp][ws] for ws in workstations for Emp in Employees])

    for emp in Employees:
        problem += p.lpSum([Assignments[(ws, emp)] for ws in workstations]) <= 1

    for ws in workstations:
        problem += p.lpSum([Assignments[(ws, emp)] * skills_present.loc[emp][ws] for emp in Employees]) <= 1

    for ws in workstations:
        for emp in Employees:
            if emp == max(Employees) or emp == (max(Employees) - 1):
                problem += p.lpSum((Assignments[(ws, emp)] * skills_present.loc[emp][ws])) == 0

    # for emp in Employees:
    #    if (emp != max(Employees)) and emp != (max(Employees) - 1):
    #       prob += p.lpSum([Assignments[(ws, emp)] for ws in workstations])

    # for emp in Employees: prob += p.lpSum(Worksheet[ws][emp] * [Assignments[(ws, emp)] for ws in workstations]) <= 1

    problem += p.lpSum(
        [Assignments[(ws, Emp)] for Emp in Employees for ws in workstations]) >= preference_based_assignments
    problem.solve()

    solve_status = p.LpStatus[problem.status]
    # print(solve_status)

    if solve_status == 'Optimal':
        result = pd.DataFrame(list(Assignments.items())).applymap(lambda x: x).T
        preferred_assignments = (len(problem.objective))
        # print('Objective Value is: {}'.format(preferred_assignments))
        excess_people = result.sum(axis=1).loc[lambda x: x == 0].index.to_list()
        assinged_people_cost = sum(result) * unit_cost

        for ws in workstations:
            for emp in Employees:
                Worksheet[ws][emp] = Assignments[(ws, emp)].varValue

        print(Worksheet)
        return True, result, excess_people, assinged_people_cost, preferred_assignments
    else:
        return False, None, None, None


assign_stations3()
# a = (preference_maximize(Work_Sheet_G, SkillsG, PreferenceG))

# First model assigns employees based in employees havinf lesser number of skills
# Second model assigns employees based on stations having lesser number of employees first
# Third model develop such that it assigns based on both stations and employees having lesser number of options
# First arrange employees based on their historic attendance percentages
# then check for stations having lowest number of possible employees who can work there based
#  then check for employees (if multiple) then select according to employee having lesser number of options