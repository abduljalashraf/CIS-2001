from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
import numpy as np
import xlrd
import copy
# from pulp import LpVariable, LpMaximize, LpStatus, lpSum, LpProblem, LpSolver
import pulp as p
import openpyxl


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

# Production_Leader = ['A12', 'B12', 'C11', 'D8', 'E8', 'F10', 'G10']
# Tag_Leader = ['A13', 'B13', 'C12', 'D9', 'E9', 'F11', 'G11']

# Stations_exc_PL_TR_A = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11']
Stations = [TeamA_Stations, TeamB_Stations, TeamC_Stations, TeamD_Stations, TeamE_Stations, TeamF_Stations,
            TeamG_Stations]

SkillsA = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsA.csv',index_col=0)
# print(SkillsA)

SkillsB = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsB.csv', index_col=0)
# print(SkillsB)
SkillsC = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsC.csv', index_col=0)
# a = pd.DataFrame(SkillsC)
# print(SkillsC)
SkillsD = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsD.csv', index_col=0)
# print(SkillsD)
SkillsE = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsE.csv', index_col=0)
# print(SkillsE)
SkillsF = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsF.csv', index_col=0)
# print(SkillsF)
SkillsG = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsG.csv', index_col=0)
# print(SkillsG)
# Skills_set = [SkillsA, SkillsB, SkillsC, SkillsD, SkillsE, SkillsF, SkillsG]
# print(SkillsA['A1'])
# print(SkillsA)
PreferenceA = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferA.csv', index_col=0)
PreferenceB = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferB.csv', index_col=0)
PreferenceC = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferC.csv', index_col=0)
PreferenceD = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferD.csv', index_col=0)
PreferenceE = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferE.csv', index_col=0)
PreferenceF = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferF.csv', index_col=0)
PreferenceG = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Preference\all_teams_preferG.csv', index_col=0)

# Preference = [PreferenceA, PreferenceB, PreferenceC, PreferenceD, PreferenceE, PreferenceF, PreferenceG]
# print(PreferenceA)
# print(PreferenceA.loc[1001])

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
Work_Sheet_A = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetA.csv', index_col=0)
Work_Sheet_B = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetB.csv', index_col=0)
Work_Sheet_C = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetC.csv', index_col=0)
Work_Sheet_D = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetD.csv', index_col=0)
Work_Sheet_E = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetE.csv', index_col=0)
Work_Sheet_F = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetF.csv', index_col=0)
Work_Sheet_G = pd.read_csv(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetG.csv', index_col=0)

# Work_Sheets = [Work_Sheet_A, Work_Sheet_B, Work_Sheet_C, Work_Sheet_D, Work_Sheet_E, Work_Sheet_F, Work_Sheet_G]
# print(Work_Sheet_A)
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


def preference_maximize(Worksheet, skills, preference, unit_cost=FullCost):
    # print('Preference is :')
    # print(preference)
    print(Worksheet)
    people_present = Worksheet.index.to_list()
    # print(people_present)
    preference_based_assignments = Worksheet.sum().sum()
    # print(preference_based_assignments)
    skills_present = skills[skills.index.isin(people_present)]
    #print(skills_present.loc[1001][1])
    skills_present = skills_present.applymap(lambda x: 0.5 if x == 0 else x)
    #print(skills_present.loc[1001][2])
    workstations_list = skills.columns.to_list()
    # print(workstations_list)
    employees = skills.index.to_list()
    # print('--------------')
    # print(int(employees[2]))
    # print('------------')
    prob = p.LpProblem("preference", p.LpMaximize)
    assignment = p.LpVariable.dicts('Prefer', ((ws, emp) for ws in workstations_list for emp in employees), cat='Binary')
    # print(assignment[(1, 1001)].value())
    # print(assignment(1, 1001))
    # for key in assignment.keys():
    # for i in [1001.0, 1002.0, 1003.0, 1004.0, 1005.0, 1006.0, 1007.0, 1008.0, 1009.0, 1010.0, 1011.0, 1012.0, 1013.0, 1014.0, 1015.0]:
    #    assignment.pop('Unnamed: 14', i)
    # print(assignment.keys())
    prob += p.lpSum([assignment[(ws, Emp)] * preference.loc[Emp][ws] for ws in workstations_list for Emp in employees])
    # print('prob')
    # print(prob)
    for workstation in workstations_list:
        prob += p.lpSum([assignment[(workstation, employee)] * skills_present.loc[employee][workstation] for employee in employees]) <= 1

    for employee in employees:
        prob += p.lpSum([assignment[(workstation, employee)] for workstation in workstations_list]) <= 1

    # for employee in employees:
    #    if (employee != max(employee)) and employee != (max(employees) - 1):
     #       prob += p.lpSum([assignment[(workstation, employee)] for workstation in workstations_list])

    # for employee in employees:
    #    prob += p.lpSum(Worksheet[workstation][employee] * [assignment[(workstation, employee)] for workstation in workstations_list]) <= 1

    prob += p.lpSum([assignment[(ws, Emp)] for Emp in employees for ws in workstations_list]) >= preference_based_assignments
    # print('PROB')
    # print(prob)
    prob.solve()

    solve_status = p.LpStatus[prob.status]
    print(solve_status)
    # print(type(assignment))
    # print(pd.DataFrame(list(assignment.items())))
    # print(pd.DataFrame(assignment).applymap(lambda x: x))
    # for v in prob.variables():
        # print(v)

    if solve_status == 'Optimal':
        result = pd.DataFrame(list(assignment.items())).applymap(lambda x: x).T
        #print('Printing values \n')
        #print(result.values)
        preference_based_count = (value(prob.objective))
        print('Objective Value is: {}'.format(preference_based_count))
        # print(preference_based_count.is_expression_type)
        #print()
        #print(preference_based_count)
        excess_people = result.sum(axis=1).loc[lambda x: x == 0].index.to_list()
        #print('Excess People')
        # print(excess_people)
#        excess_people = result.loc[lambda x: x == 0].sum(axis=1).index.to_list()
        assinged_people_cost = sum(result) * unit_cost
#        preference_based_count
        #for v in prob.variables():
        #    print(v.name, "=", v.value)

        output = []
        for workstation in Worksheet:
            for employee in employees:
                var_output = {
                    'Employee': employee,
                    'Workstation': workstation,
                    'Assigned': assignment[(workstation, employee)].varValue
                }
                output.append(var_output)

        # print(output)
        # print(type(output))

        outputdf = pd.DataFrame.from_records(output)
        outputdf.to_excel(r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\all_teams_skills.xlsx', sheet_name="Sheet7")

        return True, result, excess_people, assinged_people_cost, preference_based_count
    else:
        return False, None, None, None


# a = ('b', 'a')
# print(type(a))
a = (preference_maximize(Work_Sheet_G, SkillsG, PreferenceG))
# print(a)


