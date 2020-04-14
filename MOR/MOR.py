from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
import numpy as np
import xlrd

# import Pulp
# import ortools

days = ["Monday", 'Tuesday', 'Wednesday', 'Thursday']
# Teams = ['TeamA', 'TeamB', 'TeamC', 'TeamD', 'TeamE', 'TeamF', 'TeamG']
# Teams = [TeamA_Emp,TeamB_Emp, TeamC_Emp, TeamD_Emp, TeamE_Emp, TeamF_Emp, TeamG_Emp]
TeamA_Emp = [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1012, 1013, 1014, 1015]
TeamB_Emp = [2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017]
TeamC_Emp = [3001, 3002, 3003, 3004, 3005, 3006, 3007, 3008, 3009, 3010, 3011, 3012, 3013, 3014, 3015, 3016, 3017]
TeamD_Emp = [4001, 4002, 4003, 4004, 4005, 4006, 4007, 4008, 4009, 4010, 4011]
TeamE_Emp = [5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009, 5010, 5011]
TeamF_Emp = [6001, 6002, 6003, 6004, 6005, 6006, 6007, 6008, 6009, 6010, 6011, 6012, 6013, 6014, 6015]
TeamG_Emp = [7001, 7002, 7003, 7004, 7005, 7006, 7007, 7008, 7009, 7010, 7011, 7012, 7013, 7014]

Teams = [TeamA_Emp, TeamB_Emp, TeamC_Emp, TeamD_Emp, TeamE_Emp, TeamF_Emp, TeamG_Emp]

# TeamA_Stations = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13']
TeamA_Stations = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
TeamB_Stations = ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13']
TeamC_Stations = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12']
TeamD_Stations = ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9']
TeamE_Stations = ['E1', 'E2', 'E3', 'E4', 'E5', 'E6', 'E7', 'E8', 'E9']
TeamF_Stations = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11']
TeamG_Stations = ['G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11']

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

FullCost = 320
GoHome_Cost = 160
NoCost = 0

Current_Attendance = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                                   r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'New', index_col=0)
# print(Current_Attendance)
Work_Sheet = pd.read_excel(r'C:\Users\ashra\Downloads\Models of Operational '
                           r'Research\Course_Project\Default_Project\all_team_att.xlsx', 'WorkSheet', index_col=0)

Employee_Assigned = {1001: 0, 1002: 0, 1003: 0, 1004: 0, 1005: 0, 1006: 0, 1007: 0, 1008: 0, 1009: 0, 1010: 0, 1011: 0,
                     1012: 0, 1013: 0, 1014: 0, 1015: 0}

Station_Assigned = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 13: 0}

Attendance_Sheet = {1001: 0, 1002: 0, 1003: 0, 1004: 0, 1005: 0, 1006: 0, 1007: 0, 1008: 0, 1009: 0, 1010: 0, 1011: 0,
                    1012: 0, 1013: 0, 1014: 0, 1015: 0}


def assign_stations():
    unassigned_employees = []
    Total_Cost = 0
    no_of_employees_working = 0
    no_of_stations_operating = 0
    for Emp in TeamA_Emp:
        sum_row = (SkillsA.sum(axis=1))
        # print(sum_row)
        # print((Skills_A[0]))
        if sum_row[Emp] == 1:
            # print('Employee having 1 skill only')
            for ws in TeamA_Stations:
                if SkillsA[ws][Emp] == 1:
                    Work_Sheet[ws][Emp] = 1
                    Employee_Assigned[Emp] += 1
                    no_of_employees_working += 1
                    no_of_stations_operating += 1
                    Station_Assigned[ws] += 1
                    Attendance_Sheet[Emp] += 1
                    Total_Cost += FullCost
                    print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                    continue
        else:
            for ws in TeamA_Stations:
                if SkillsA[ws][Emp] == 1 and Employee_Assigned[Emp] < 1 and Station_Assigned[ws] < 1:
                    Work_Sheet[ws][Emp] = 1
                    Employee_Assigned[Emp] += 1
                    Station_Assigned[ws] += 1
                    Attendance_Sheet[Emp] += 1
                    no_of_employees_working += 1
                    no_of_stations_operating += 1
                    Total_Cost += FullCost
                    print('Employee {} has been assigned to workstation A{}'.format(Emp, ws))
                    continue

    if no_of_employees_working < len(TeamA_Emp):
        for ws in Station_Assigned:
            if Station_Assigned[ws] == 0 and Employee_Assigned[(max(TeamA_Emp))] < 1:
                no_of_employees_working += 1
                no_of_stations_operating += 1
                Station_Assigned[ws] += 1
                Employee_Assigned[(max(TeamA_Emp))] += 1
                Total_Cost += FullCost
                print()
                print('Tag Relief {} has been assigned to workstation A{}'.format((max(TeamA_Emp)), ws))
                break
            continue

    # print('Remaining Employees')

    for Emp in TeamA_Emp:
        if Employee_Assigned[Emp] == 0:
            unassigned_employees.append(Emp)
        if len(unassigned_employees) != 0:
            print('Employee {} hasn\'t been assigned any workstation'.format(unassigned_employees))
        else:
            print()
            print('All assignments done, no cross training required')
            print('No Remaining Employees')
            break

        input_task = input('Enter 1 to cross train employees, else 2 to send home')
        if input_task == 1:
            Total_Cost += FullCost
            emp = input('Enter the employee you wish to cross train')
            unassigned_employees.remove(emp)
            list_of_station_options = []
            for ws in TeamA_Stations:
                if SkillsA[emp][ws] == 0:
                    list_of_station_options.append(ws)

            station = input('Enter station from the following to train upon {}'.format(list_of_station_options))
            if station in list_of_station_options:
                SkillsA[station][emp] = 1
                Attendance_Sheet[emp] += 2

            if len(unassigned_employees) != 0:
                for emp in TeamA_Emp:
                    if Employee_Assigned[emp] == 0:
                        Total_Cost += GoHome_Cost
                        Attendance_Sheet[emp] += 1
                        print('Employee {} is being sent home'.format(unassigned_employees))
            break

        elif input_task == 2:
            for emp in TeamA_Emp:
                if Employee_Assigned[emp] == 0:
                    Total_Cost += GoHome_Cost
                    Attendance_Sheet[emp] += 1
                    print('Employee {} is being sent home'.format(unassigned_employees))
            break

    print()
    print('No. of employees present at work: {}'.format(no_of_employees_working))
    print('No. of workstations running {}'.format(no_of_stations_operating))
    print()
    print('Total Cost is {}$'.format(Total_Cost))

    # print(Work_Sheet)
    # print('Assigned Employees are: ')
    # for i in TeamA_Emp:
    # for j in TeamA_Stations:
    # if Work_Sheet[i][j] == 1:
    #   print('Employee {} has been assigned to workstation A{}'.format(TeamA_Emp[i], TeamA_Stations[j]))


assign_stations()

print("Assignment Matrix for Team A")
print()
print(Work_Sheet)
