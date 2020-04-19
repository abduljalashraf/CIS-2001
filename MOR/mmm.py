from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
import numpy as np
import xlrd
import copy
import pulp as p
import openpyxl

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

Tag_Relief = [1015, 2017, 3017, 4011, 5011, 6015, 7014]
Production_Leader = [1014, 2016, 3016, 4010, 5010, 6014, 7013]

SkillsA = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Skills\all_teams_skillsA.csv',
    index_col=0)
# print(SkillsA)

SkillsB = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsB.csv',
    index_col=0)
# print(SkillsB)
SkillsC = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsC.csv',
    index_col=0)
# a = pd.DataFrame(SkillsC)
# print(SkillsC)
SkillsD = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsD.csv',
    index_col=0)
# print(SkillsD)
SkillsE = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsE.csv',
    index_col=0)
# print(SkillsE)
SkillsF = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsF.csv',
    index_col=0)
# print(SkillsF)
SkillsG = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Skills\all_teams_skillsG.csv',
    index_col=0)
# print(SkillsG)
# Skills_set = [SkillsA, SkillsB, SkillsC, SkillsD, SkillsE, SkillsF, SkillsG]
# print(SkillsA['A1'])
# print(SkillsA)
PreferenceA = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferA.csv',
    index_col=0)
PreferenceB = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferB.csv',
    index_col=0)
PreferenceC = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferC.csv',
    index_col=0)
PreferenceD = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferD.csv',
    index_col=0)
PreferenceE = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferE.csv',
    index_col=0)
PreferenceF = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferF.csv',
    index_col=0)
PreferenceG = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Preference\all_teams_preferG.csv',
    index_col=0)

# Preference = [PreferenceA, PreferenceB, PreferenceC, PreferenceD, PreferenceE, PreferenceF, PreferenceG]

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
Work_Sheet_A = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format '
    r'files\Attendance\WorksheetA.csv',
    index_col=0)
Work_Sheet_B = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetB.csv',
    index_col=0)
Work_Sheet_C = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetC.csv',
    index_col=0)
Work_Sheet_D = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetD.csv',
    index_col=0)
Work_Sheet_E = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetE.csv',
    index_col=0)
Work_Sheet_F = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetF.csv',
    index_col=0)
Work_Sheet_G = pd.read_csv(
    r'C:\Users\ashra\Downloads\Models of Operational Research\Course_Project\Default_Project\CSV format files\Attendance\WorksheetG.csv',
    index_col=0)


# Work_Sheets = [Work_Sheet_A, Work_Sheet_B, Work_Sheet_C, Work_Sheet_D, Work_Sheet_E, Work_Sheet_F, Work_Sheet_G]


def preference_maximize(Worksheet, skills, preference, unit_cost=FullCost):
    print('Worksheet before optimization and Assignments are done')
    print(Worksheet)
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

    # for emp in Employees:
    #    if (emp != max(Employees)) and emp != (max(Employees) - 1):
    #       prob += p.lpSum([Assignments[(ws, emp)] for ws in workstations])

    # for emp in Employees: prob += p.lpSum(Worksheet[ws][emp] * [Assignments[(ws, emp)] for ws in workstations]) <= 1

    problem += p.lpSum(
        [Assignments[(ws, Emp)] for Emp in Employees for ws in workstations]) >= preference_based_assignments
    problem.solve()

    solve_status = p.LpStatus[problem.status]
    print(solve_status)

    if solve_status == 'Optimal':
        result = pd.DataFrame(list(Assignments.items())).applymap(lambda x: x).T
        preferred_assignments = (len(problem.objective))
        print('Objective Value is: {}'.format(preferred_assignments))
        excess_people = result.sum(axis=1).loc[lambda x: x == 0].index.to_list()
        assinged_people_cost = sum(result) * unit_cost

        for ws in workstations:
            for emp in Employees:
                Worksheet[ws][emp] = Assignments[(ws, emp)].varValue

        print(Worksheet)
        return True, result, excess_people, assinged_people_cost, preferred_assignments
    else:
        return False, None, None, None


a = (preference_maximize(Work_Sheet_G, SkillsG, PreferenceG))

