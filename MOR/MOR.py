from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
import numpy as np
import xlrd

# import Pulp
# import ortools

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

# TeamA_Stations = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13']
# TeamB_Stations = ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13']
# TeamC_Stations = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12']
# TeamD_Stations = ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9']
# TeamE_Stations = ['E1', 'E2', 'E3', 'E4', 'E5', 'E6', 'E7', 'E8', 'E9']
# TeamF_Stations = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11']
# TeamG_Stations = ['G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11']

Production_Leader = ['A12', 'B12', 'C11', 'D8', 'E8', 'F10', 'G10']
Tag_Leader = ['A13', 'B13', 'C12', 'D9', 'E9', 'F11', 'G11']

Stations_exc_PL_TR_A = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11']
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
                       1011: 0,
                       1012: 0, 1013: 0, 1014: 0, 1015: 0}
Station_Assigned_A = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 13: 0}
Attendance_Sheet_A = {1001: 0, 1002: 0, 1003: 0, 1004: 0, 1005: 0, 1006: 0, 1007: 0, 1008: 0, 1009: 0, 1010: 0, 1011: 0,
                      1012: 0, 1013: 0, 1014: 0, 1015: 0}

Employee_Assigned_B = {2001: 0, 2002: 0, 2003: 0, 2004: 0, 2005: 0, 2006: 0, 2007: 0, 2008: 0, 2009: 0, 2010: 0,
                       2011: 0,
                       2012: 0, 2013: 0, 2014: 0, 2015: 0, 2016: 0, 2017: 0}
Station_Assigned_B = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 13: 0}
Attendance_Sheet_B = {2001: 0, 2002: 0, 2003: 0, 2004: 0, 2005: 0, 2006: 0, 2007: 0, 2008: 0, 2009: 0, 2010: 0, 2011: 0,
                      2012: 0, 2013: 0, 2014: 0, 2015: 0, 2016: 0, 2017: 0}

Employee_Assigned_C = {3001: 0, 3002: 0, 3003: 0, 3004: 0, 3005: 0, 3006: 0, 3007: 0, 3008: 0, 3009: 0, 3010: 0,
                       3011: 0, 3012: 0, 3013: 0, 3014: 0, 3015: 0, 3016: 0, 3017: 0}
Station_Assigned_C = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0}
Attendance_Sheet_C = {3001: 0, 3002: 0, 3003: 0, 3004: 0, 3005: 0, 3006: 0, 3007: 0, 3008: 0, 3009: 0, 3010: 0,
                      3011: 0, 3012: 0, 3013: 0, 3014: 0, 3015: 0, 3016: 0, 3017: 0}

Employee_Assigned_D = {4001: 0, 4002: 0, 4003: 0, 4004: 0, 4005: 0, 4006: 0, 4007: 0, 4008: 0, 4009: 0, 4010: 0,
                       4011: 0, 4012: 0}
Station_Assigned_D = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0}
Attendance_Sheet_D = {4001: 0, 4002: 0, 4003: 0, 4004: 0, 4005: 0, 4006: 0, 4007: 0, 4008: 0, 4009: 0, 4010: 0, 4011: 0,
                      4012: 0}

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

'''
def assign_stations():
    unassigned_employees = []
    Total_Cost = 0
    no_of_employees_working = 0
    no_of_stations_operating = 0
    for Emp in Teams:
        sum_row = (SkillsA.sum(axis=1))
        # print(sum_row)
        # print((Skills_A[0]))
        if sum_row[Emp] == 1:
            # print('Employee having 1 skill only')
            for ws in TeamA_Stations:
                if SkillsA[ws][Emp] == 1 and Station_Assigned_A[ws] < 1 and Employee_Assigned_A[Emp] < 1:
                    Work_Sheet[ws][Emp] = 1
                    Employee_Assigned_A[Emp] += 1
                    no_of_employees_working += 1
                    no_of_stations_operating += 1
                    Station_Assigned_A[ws] += 1
                    Attendance_Sheet_A[Emp] += 1
                    Total_Cost += FullCost
                    print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                    continue
        else:
            for ws in TeamA_Stations:
                if SkillsA[ws][Emp] == 1 and Employee_Assigned_A[Emp] < 1 and Station_Assigned_A[ws] < 1 and (
                        Emp != max(Teams) and
                        Emp != (max(Teams) - 1)):
                    Work_Sheet[ws][Emp] = 1
                    Employee_Assigned_A[Emp] += 1
                    Station_Assigned_A[ws] += 1
                    Attendance_Sheet_A[Emp] += 1
                    no_of_employees_working += 1
                    no_of_stations_operating += 1
                    Total_Cost += FullCost
                    print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                    continue

    # print(no_of_employees_working)
    # print(no_of_stations_operating)
    # print(len(Teams))

    if int(no_of_employees_working) < len(Teams) and (len(Teams) - no_of_employees_working - 2) == 2:
        list = []
        station = []
        for ws in TeamA_Stations:
            for Emp in Teams:
                if SkillsA[ws][Emp] == 0 and Employee_Assigned_A[Emp] == 0 and Station_Assigned_A[
                    ws] == 0 and Emp != max(Teams) and Emp != (max(Teams) - 1):
                    list.append(Emp)
                    station.append(ws)
                if len(list) == 2 and station[0] == station[1]:
                    print(
                        'Employees {}, {} have been assigned station A{} as they are unskilled'.format(list[0], list[1],
                                                                                                       station[0]))
                    no_of_employees_working += 2
                    no_of_stations_operating += 1
                    Total_Cost += (2 * 320)
                    Employee_Assigned_A[list[0]] = 1
                    Employee_Assigned_A[list[1]] = 1
                    Station_Assigned_A[station[0]] = 1
                    Work_Sheet[station[0]][list[0]] = 1
                    Work_Sheet[station[1]][list[1]] = 1
                    list.pop(0)
                    list.pop(0)
                    break
                continue

    if int(no_of_employees_working) < len(Teams):
        for ws in Station_Assigned_A:
            # print(Station_Assigned_A[ws])
            # print(max(Teams))
            # print('A{}'.format(Employee_Assigned_A[max((Teams))]))
            if Station_Assigned_A[ws] == 0 and Employee_Assigned_A[max(Teams)] < 1:
                no_of_employees_working += 1
                no_of_stations_operating += 1
                Station_Assigned_A[ws] += 1
                Employee_Assigned_A[(max(Teams))] += 1
                Total_Cost += FullCost
                print()
                print('Tag Relief {} has been assigned to workstation A{}'.format((max(Teams)), ws))
                break
            continue

    # print('Remaining Employees')

    for Emp in Teams:
        # print(Emp)
        if int(Employee_Assigned_A[Emp]) != 0:
            continue
        elif int(Employee_Assigned_A[Emp]) == 0 and Emp != (max(Teams) - 1):
            unassigned_employees.append(Emp)

        # print(unassigned_employees)
    if len(unassigned_employees) == 0:
        print()
        print('All assignments done, no cross training required')
        print('No Remaining Employees')
    else:  # len(unassigned_employees) == 0:
        print('Employee {} hasn\'t been assigned any workstation'.format(unassigned_employees))

        input_task = input('Enter 1 to cross train employees, else 2 to send home')
        if int(input_task) == 1:
            Total_Cost += FullCost
            emp = input('Enter the employee you wish to cross train')
            emp = int(emp)
            unassigned_employees.remove(int(emp))
            list_of_station_options = []
            for ws in TeamA_Stations:
                if SkillsA[ws][emp] == 0:
                    list_of_station_options.append(ws)
                    continue
            if len(list_of_station_options) == 0:
                print('Employee {} is trained in all stations'.format(emp))
            else:
                station = input('Enter station from the following to train upon {}'.format(list_of_station_options))
                if int(station) in list_of_station_options:
                    SkillsA[int(station)][emp] = 1
                    Attendance_Sheet_A[emp] += 2

                if len(unassigned_employees) != 0:
                    for emp in Teams:
                        if Employee_Assigned_A[emp] == 0:
                            Total_Cost += GoHome_Cost
                            Attendance_Sheet_A[emp] += 1
                            for employee in unassigned_employees:
                                print('Employee {} is being sent home'.format(employee))
                        # break

        elif int(input_task) == 2:
            for emp in Teams:
                if Employee_Assigned_A[emp] == 0:
                    Total_Cost += GoHome_Cost
                    Attendance_Sheet_A[emp] += 1
                    print('Employee {} is being sent home'.format(unassigned_employees))
        # break

    print()
    print('No. of employees present at work: {}'.format(no_of_employees_working))
    print('No. of workstations running {}'.format(no_of_stations_operating))
    print()
    print('Total Cost is {}$'.format(Total_Cost))

    # print(Work_Sheet)
    # print('Assigned Employees are: ')
    # for i in Teams:
    # for j in TeamA_Stations:
    # if Work_Sheet[i][j] == 1:
    #   print('Employee {} has been assigned to workstation A{}'.format(Teams[i], TeamA_Stations[j]))


assign_stations()
'''


def assign_stations1():
    Overall_Cost = 0
    overall_no_of_employees_working = 0
    overall_no_of_stations_operating = 0
    overall_unassigned_employees = []
    overall_employees_sent_home = []
    overall_employees_cross_trained = []
    no_of_employees_cross_trained = 0
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
        Team_Station = Stations[i]
        Work_Sheet = Work_Sheets[i]
        Current_Attendance = Current_Attendance_Sheets[i]
        for day in days:
            for Emp in Team:
                sum_row = (Skills.sum(axis=1))
                if sum_row[Emp] == 1:
                    # print('Employee having 1 skill only')
                    for ws in Team_Station:
                        if Skills[ws][Emp] == 1 and Station_Assigned[ws] < 1 and Employee_Assigned[Emp] < 1:
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] += 1
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                            continue
                else:
                    for ws in Team_Station:
                        # print('{}++{}--{}***{}'.format(Station_Assigned[ws], ws, Employee_Assigned[Emp], Emp))
                        if Skills[ws][Emp] == 1 and Employee_Assigned[Emp] < 1 and Station_Assigned[ws] < 1 and (
                                Emp != max(Team)) and (Emp != (max(Team) - 1)):
                            Work_Sheet[ws][Emp] = 1
                            Employee_Assigned[Emp] += 1
                            Station_Assigned[ws] += 1
                            Attendance_Sheet[Emp] += 1
                            no_of_employees_working += 1
                            no_of_stations_operating += 1
                            Current_Attendance[day][Emp] += 1
                            Total_Cost += FullCost
                            print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                            continue

            # print(no_of_employees_working)
            # print(no_of_stations_operating)
            # print(len(Teams))

            if int(no_of_employees_working) < len(Team) and (len(Team) - no_of_employees_working - 2) == 2:
                list = []
                station = []
                for ws in Team_Station:
                    for Emp in Team:
                        if Skills[ws][Emp] == 0 and Employee_Assigned[Emp] == 0 and Station_Assigned[ws] == 0 and Emp != max(Team) and Emp != (max(Team) - 1):
                            list.append(Emp)
                            station.append(ws)
                        if len(list) == 2 and station[0] == station[1]:
                            print(
                                'Employees {}, {} have been assigned station A{} as they are unskilled'.format(list[0],
                                                                                                               list[1],
                                                                                                               station[
                                                                                                                   0]))
                            no_of_employees_working += 2
                            no_of_stations_operating += 1
                            Current_Attendance[day][list[0]] += 1
                            Current_Attendance[day][list[1]] += 1
                            Total_Cost += (2 * 320)
                            Employee_Assigned[list[0]] = 1
                            Employee_Assigned[list[1]] = 1
                            Station_Assigned[station[0]] = 1
                            Work_Sheet[station[0]][list[0]] = 1
                            Work_Sheet[station[1]][list[1]] = 1
                            list.pop(0)
                            list.pop(0)
                            for emp in list:
                                overall_employees_sent_home.append(list)
                            break
                        continue

            if int(no_of_employees_working) < len(Team):
                for ws in Station_Assigned:
                    # print(Station_Assigned_A[ws])
                    # print(max(Teams))
                    # print('A{}'.format(Employee_Assigned_A[max((Teams))]))
                    if Station_Assigned[ws] == 0 and Employee_Assigned[max(Team)] < 1:
                        no_of_employees_working += 1
                        no_of_stations_operating += 1
                        Current_Attendance[day][max(Team)] += 1
                        Station_Assigned[ws] += 1
                        Employee_Assigned[(max(Team))] += 1
                        Total_Cost += FullCost
                        print()
                        print('Tag Relief {} has been assigned to workstation A{}'.format((max(Team)), ws))
                        break
                    continue

            # print('Remaining Employees')

            for Emp in Team:
                # print(Emp)
                if int(Employee_Assigned[Emp]) != 0:
                    continue
                elif int(Employee_Assigned[Emp]) == 0 and Emp != (max(Team) - 1):
                    unassigned_employees.append(Emp)

            # print(unassigned_employees)
            if len(unassigned_employees) == 0:
                print()
                print('All assignments done, no cross training required')
                print('No Remaining Employees')
            else:  # len(unassigned_employees) == 0:
                print('Employee {} hasn\'t been assigned any workstation'.format(unassigned_employees))

                input_task = input('Enter 1 to cross train employees, else 2 to send home')
                if int(input_task) == 1:
                    Total_Cost += FullCost
                    emp = input('Enter the employee you wish to cross train')
                    emp = int(emp)
                    no_of_employees_cross_trained += 1
                    overall_employees_cross_trained.append(emp)
                    unassigned_employees.remove(int(emp))
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
                            Attendance_Sheet[emp] += 2

                        if len(unassigned_employees) != 0:
                            # for emp in Team:
                            for employee in unassigned_employees:
                                if Employee_Assigned[employee] == 0:
                                    Total_Cost += GoHome_Cost
                                    Attendance_Sheet[employee] += 1
                                    Current_Attendance[day][employee] += 1
                                    print('Employee {} is being sent Home'.format(employee))
                                    overall_employees_sent_home.append(employee)
                                # break

                elif int(input_task) == 2:
                    for emp in Team:
                        if Employee_Assigned[emp] == 0:
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
            print('Assignment Matrix of Team', i)
            print()
            print(Work_Sheet)
            print()
            Overall_Cost += Total_Cost
            overall_no_of_employees_working += no_of_employees_working
            overall_no_of_stations_operating += no_of_stations_operating

        # print(Work_Sheet)
        # print('Assigned Employees are: ')
        # for i in Teams:
        # for j in TeamA_Stations:
        # if Work_Sheet[i][j] == 1:
        #   print('Employee {} has been assigned to workstation A{}'.format(Teams[i], TeamA_Stations[j]))

    print("Assignment Matrix for Team A")
    # print()
    # print(Work_Sheet)
    print()
    # print('Employee_Assigned in Team {}'.format(i))
    print()
    # print()
    # print('Station_Assigned in Team {}'.format(i))
    # print(Station_Assigned)
    # print()
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


assign_stations1()
