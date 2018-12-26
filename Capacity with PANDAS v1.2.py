# In this version, labor cost changes in function of the day AND the job.
# The cost is used to prioritize people according to their qualification

#################################################### READ ME!!!! ###################################################
#TO DO
# I. Not all employees are supposed to work for eveyr job.
# Make sure the cumulative function gets the sum of working hours of only the related employees
# That would render the model more efficient

# To make this faster dataframes have a cumsum function. For converting to dataframe converting tuples to multiindex might be the key.
#Look at the following link:
# https://pandas.pydata.org/pandas-docs/stable/generated/pandas.MultiIndex.from_tuples.html

#FIXED? >>> Keep testing
#time range you use in the 3 blocks should be coherent:
    #where you declare emphour variables
    #HERE
    #where you declare precedence variable
####################################################################################################################

import pulp
import pandas as pd
import numpy as np
import datetime
import csv
import xlrd
import math

# read jobs from the xlsx file
xl_wb = xlrd.open_workbook("C:/Users/gurkaali/Documents/Info/Ben/Case Study/InputFile.xlsx")
xl_sht = xl_wb.sheet_by_index(0)

jobs_df = pd.DataFrame()
for c in range(1, xl_sht.ncols):
    if len(xl_sht.cell(2, c).value) == 0: #this cell onwards no input exists
        break
    jobs_df = jobs_df.append([[xl_sht.cell(2, c).value, int(xl_sht.cell(3, c).value), int(xl_sht.cell(4, c).value), int(xl_sht.cell(5, c).value)]], ignore_index=True)
jobs_df.columns = ['jobs', 'earliest_start', 'latest_end','job_hours']
jobs_df.set_index('jobs', inplace = True)

jobs = list(jobs_df.index.values)

# Number of max workdays to plan is given by the suer input
schedule_end = int(xl_sht.cell(20, 1).value) + 1
days = list(range(1, schedule_end))

employees = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(11, c).value)) == 0: #this cell onwards no input exists
        break
    employees.append(xl_sht.cell(11, c).value)
#Employees by tasks in order of their profeciency - except Assembly where everyone is equally qualified
employees_design = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(13, c).value)) == 0: #this cell onwards no input exists
        break
    employees_design.append(xl_sht.cell(13, c).value)

employees_turning = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(14, c).value)) == 0: #this cell onwards no input exists
        break
    employees_turning.append(xl_sht.cell(14, c).value)

employees_milling = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(15, c).value)) == 0: #this cell onwards no input exists
        break
    employees_milling.append(xl_sht.cell(15, c).value)

employees_assembly = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(16, c).value)) == 0: #this cell onwards no input exists
        break
    employees_assembly.append(xl_sht.cell(16, c).value)

employees_measurement = []
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(17, c).value)) == 0: #this cell onwards no input exists
        break
    employees_measurement.append(xl_sht.cell(17, c).value)

# ADD OVERTIME COSTS DICTIONARY !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
costs = {}
for j in jobs:
    costs[j] = {}
    #Design job has equally qualified employees
    if j[-1] == 'D':
        for e in employees_design:
            costs[j][e] = {}
            for d in days:
                costs[j][e][d] = {}
                if e in employees_design:
                    costs[j][e][d] = 10 + d * 0.0001
                else:
                    costs[j][e][d] = 88888 + d * 0.0001

    #the 1st person in the employees_turning list has a 1st day cost of 10 incrementing till 10.05 at day 500
    #the 2nd person in the employees_turning list has a 1st day cost of 11 incrementing till 11.05 at day 500
    #the last day cost of better qualified person must be below the 1st day cost of the less qualified person
    if j[-1] == 'T':
        for e in employees_turning:
            costs[j][e] = {}
            for d in days:
                costs[j][e][d] = {}
                if e in employees_turning:
                    costs[j][e][d] = (10 + employees_turning.index(e)) + d * 0.0001
                else:
                    costs[j][e][d] = 88888 + d * 0.0001
    if j[-1] == 'M':
        for e in employees_milling:
            costs[j][e] = {}
            for d in days:
                costs[j][e][d] = {}
                if e in employees_milling:
                    costs[j][e][d] = (10 + employees_milling.index(e)) + d * 0.0001
                else:
                    costs[j][e][d] = 88888 + d * 0.0001
    if j[-1] == 'A':
        for e in employees_assembly:
            costs[j][e] = {}
            for d in days:
                costs[j][e][d] = {}
                if e in employees_assembly:
                    costs[j][e][d] = (10 + employees_assembly.index(e)) + d * 0.0001
                else:
                    costs[j][e][d] = 88888 + d * 0.0001
    if j[-1] == 'E':
        for e in employees_measurement:
            costs[j][e] = {}
            for d in days:
                costs[j][e][d] = {}
                if e in employees_measurement:
                    costs[j][e][d] = (10 + employees_measurement.index(e)) + d * 0.0001
                else:
                    costs[j][e][d] = 88888 + d * 0.0001
#
# # Populate the overtime cost table based on the regular hour costs
# costs_overtime = {}
#
# for j in jobs:
#     costs_overtime[j] = {}
#     if j[-1] == 'D':
#         for e in employees:
#             costs_overtime[j][e] = {}
#             for d in days:
#                 costs_overtime[j][e][d] = {}
#                 if e in employees_design:
#                     costs_overtime[j][e][d] = 100 + d * 0.0001
#                 else:
#                     costs_overtime[j][e][d] = 99999 + d * 0.0001
#     #the 1st person in the employees_turning list has a 1st day cost of 10 incrementing till 10.05 at day 500
#     #the 2nd person in the employees_turning list has a 1st day cost of 11 incrementing till 11.05 at day 500
#     #the last day cost of better qualified person must be below the 1st day cost of the less qualified person
#     if j[-1] == 'T':
#         for e in employees:
#             costs_overtime[j][e] = {}
#             for d in days:
#                 costs_overtime[j][e][d] = {}
#                 if e in employees_turning:
#                     costs_overtime[j][e][d] = (100 + employees_turning.index(e)) + d * 0.0001
#                 else:
#                     costs_overtime[j][e][d] = 99999 + d * 0.0001
    # if j[-1] == 'M':
    #     for e in employees:
    #         costs_overtime[j][e] = {}
    #         for d in days:
    #             costs_overtime[j][e][d] = {}
    #             if e in employees_milling:
    #                 costs_overtime[j][e][d] = (100 + employees_milling.index(e)) + d * 0.0001
    #             else:
    #                 costs_overtime[j][e][d] = 99999 + d * 0.0001
    # if j[-1] == 'A':
    #     for e in employees:
    #         costs_overtime[j][e] = {}
    #         for d in days:
    #             costs_overtime[j][e][d] = {}
    #             if e in employees_assembly:
    #                 costs_overtime[j][e][d] = (100 + employees_assembly.index(e)) + d * 0.0001
    #             else:
    #                 costs_overtime[j][e][d] = 99999 + d * 0.0001
    # if j[-1] == 'E':
    #     for e in employees:
    #         costs_overtime[j][e] = {}
    #         for d in days:
    #             costs_overtime[j][e][d] = {}
    #             if e in employees_measurement:
    #                 costs_overtime[j][e][d] = (100 + employees_measurement.index(e)) + d * 0.0001
    #             else:
    #                 costs_overtime[j][e][d] = 99999 + d * 0.0001

costs_reformed = {(outerKey, innerKey, innermostKey): values for outerKey, innerDict in costs.items() for innerKey, innestDict in innerDict.items() for innermostKey, values in innestDict.items()}
df_input = pd.Series(costs_reformed).reset_index()
df_input.columns = ['Jobs', 'Employees', 'Days', 'Costs']
df_input.set_index(['Jobs', 'Employees', 'Days'], inplace=True)
df_input = df_input.reorder_levels(['Employees', 'Jobs', 'Days']) #this is to reorder levels and match with the previous version. The loops below follow the previous order

model = pulp.LpProblem("My LP Problem", pulp.LpMinimize)

# variables & parameters for precedence constraints
# J2 can start only after the end of J1 & J3 can start only after the end of J1:
job_sequences_df = pd.DataFrame()
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(7, c).value)) == 0: #this cell onwards no input exists
        break
    job_sequences_df = job_sequences_df.append([[xl_sht.cell(7, c).value, xl_sht.cell(9, c).value]], ignore_index=True)
job_sequences_df.columns = ['precedent', 'antecedent']

# Optimize constraints: Due to precedence, there cannot be work on every day for every job
df = pd.merge(jobs_df,job_sequences_df[['precedent','antecedent']],left_index = True, right_on='precedent', how='left')
df.rename(columns={'precedent':'job'}, inplace = True)
df2 = pd.merge(df,job_sequences_df[['precedent','antecedent']],left_on = 'job', right_on='antecedent', how='left')
df2 = df2[['job_hours',  'earliest_start',  'latest_end', 'job', 'antecedent_x', 'precedent']]
df2 = df2.rename(columns={'antecedent_x':'antecedents', 'precedent': 'precedents'})
# bring the latest end time of antecedents
df3 = pd.merge(df2,jobs_df[['latest_end']],left_on = 'antecedents', right_index = True, how='left')
# bring the earliest start time of precedents
df4 = pd.merge(df3,jobs_df[['earliest_start']],left_on = 'precedents', right_index = True, how='left')
df4.rename(columns={'earliest_start_x':'earliest_start', 'latest_end_x':'latest_end', 'latest_end_y':'antecedent_latest_end', 'earliest_start_y': 'precedent_earliest_start'}, inplace = True)
df5 = df4.groupby(['job','earliest_start', 'latest_end']).agg({'precedent_earliest_start': np.max, 'antecedent_latest_end': np.min})
#get rid of the index on job's earliest and latest times which prevents row level max and min calculation
df5.reset_index(inplace = True)
#then reset the index as job
df5.set_index('job', inplace = True)
# Note that this optimizes the time period for 2 consecutive jobs. In case of 3 or more jobs following each other the individual time periods of 2nd previous or 2nd later (or further) jobs are ignored
df5['practical_earliest_start'] = df5[['earliest_start','precedent_earliest_start']].max(axis=1).astype(int)
df5['practical_latest_end'] = df5[['latest_end','antecedent_latest_end']].min(axis=1).astype(int)

#clean the df dataframe: drop records with impractical days of work that are identified in df5:
compteur = 0
for j in df5.index:
    a = int(df5.loc[j]['practical_earliest_start'])
    b = int(df5.loc[j]['practical_latest_end'])
    practical_days_list = list(range(a, b + 1))
    if compteur == 0:
        df6 = df_input[(df_input.index.get_level_values('Days').isin(practical_days_list)) & (df_input.index.get_level_values('Jobs').isin([j]))]
    else:
        df6 = pd.concat([df6, df_input[(df_input.index.get_level_values('Days').isin(practical_days_list)) & (df_input.index.get_level_values('Jobs').isin([j]))]])
    compteur = compteur + 1

emp_hours = pulp.LpVariable.dicts("employee hours",
                                   ((employees, jobs, days) for employees, jobs, days in df6.index),
                                   lowBound=0,
                                   cat='Continuous')

z = pulp.LpVariable.dicts("Bin_z for precedence constraints",
                                  ((j, d) for j in job_sequences_df.antecedent.unique()
                                            for d in df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()),
                                  cat='Binary')

x = pulp.LpVariable.dicts("Bin_x for precedence constraints",
                                  ((j, d) for j in job_sequences_df.precedent.unique()
                                            for d in df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()),
                                  cat='Binary')

#get an array of unique jobs with precedence-antecedence relation
jobs_sequenced_unique = np.unique(np.append(job_sequences_df.antecedent, job_sequences_df.precedent))

# # OBJECTIVE FUNCTION
# ADD OVERTIME COSTS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
model += pulp.lpSum([emp_hours[employees, jobs, days] * df6.loc[(employees, jobs, days), 'Costs'] for employees, jobs, days
                    in df6.index])

# CONSTRAINTS
# Sum of hours spent by an employee on all jobs has to be less than or equal to 8hrs a day (or 45hrs a week)
# ADD OVERTIME COSTS VERSION AS SEPARATE BLOCK !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
for d in df6.index.levels[2]: # days
    #first get the employees who are qualified for Job = j, then get the unique list of days for that job
    for i in df6.loc[(df6.index.get_level_values('Days') == d)].index.get_level_values('Employees').unique(): # employees qualified for job j
        #get a list of jobs applicable only in Day d:
        #print('i, j , d: ', str(i), str(j), str(d))
        denyo = df6.loc[(df6.index.get_level_values('Days') == d) & (df6.index.get_level_values('Employees') == i)].index.get_level_values('Jobs').unique()
        model += pulp.lpSum([emp_hours[(i, j, d)] for j in denyo]) <= 45
print("Employees' work-hour constraints defined at ", str(datetime.datetime.now().time()))

# Total work done on a job is equal to its hour requirement
# ADD OVERTIME COSTS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
for j in jobs_df.index: #jobs
    dayyi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
    emmi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique()
    model += pulp.lpSum([emp_hours[i, j, d] for i in emmi #only the employees qualified for the job in iteration
                         for d in dayyi]) >= jobs_df.job_hours[j]
print("Jobs' work-hour constraints defined at ", str(datetime.datetime.now().time()))

Q_cum = {}
# To be able to use the dataframe cumulative sum function in cumulative hour calculation, convert the variable dictionary to a dataframe first:
emp_hours_df = pd.DataFrame(list(emp_hours.items()), columns = ('key','variable'),index=pd.MultiIndex.from_tuples(emp_hours.keys(), names=('employee', 'job', 'day')))
#info: https://thispointer.com/python-pandas-how-to-create-dataframe-from-dictionary/
# ADD OVERTIME COSTS VERSION OF THE emp_hours_df!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

# If a job is in precedence relationship with multiple jobs, no need to create Q_cum's each time.
# Create the cumulative hours only for the period it can be done i.e. between earliest start and latest finish dates
# Use the jobs_sequenced_unique where all jobs with precedence relationship are listed only once:
for j in jobs_sequenced_unique: #j is a list of 2 jobs
# INCLUDE OVERTIME COSTS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Q_cum[j] = {}
    dief_this_job = emp_hours_df.loc[emp_hours_df.index.get_level_values('job') == j]
    dief = dief_this_job.groupby(['job', 'day'])['variable'].agg('sum')
    dief.sort_index(level=['day'], inplace = True) # for the cum sum to work correctly, it must be ordered by day
    print("Q_cum for  job ", j, " defined at ", str(datetime.datetime.now().time()))

    j_earliest = int(df5.loc[j]['practical_earliest_start'])
    j_latest = int(df5.loc[j]['practical_latest_end'])
    dief_cumulus = dief.loc[(dief.index.get_level_values('day') >= j_earliest) & (dief.index.get_level_values('day') <= j_latest)].cumsum() #you had defined a variable for each day but some are reduncant e.g. those before earliest start

    dayyi_amca = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
    for d in dayyi_amca:
        #print('job: ', j, 'day: ', d, 'j_earliest:', j_earliest)
        Q_cum[j, d] = dief_cumulus[d - j_earliest]

for seq in job_sequences_df.index:
    job_pre = job_sequences_df.loc[seq]['precedent']
    job_ante = job_sequences_df.loc[seq]['antecedent']

    a1 = int(df5.loc[job_pre]['practical_earliest_start'])
    a2 = int(df5.loc[job_ante]['practical_earliest_start'])

    b1 = int(df5.loc[job_pre]['practical_latest_end'])
    b2 = int(df5.loc[job_ante]['practical_latest_end'])
    start = max(a1, a2)
    end = min(b1, b2)

    for d in range(start, end+1):
        model += pulp.lpSum([jobs_df.job_hours[job_pre] - Q_cum[job_pre, d] - 0.0001*(1 - x[job_pre, d])]) >= 0
        model += pulp.lpSum([jobs_df.job_hours[job_pre] - Q_cum[job_pre, d] - jobs_df.job_hours[job_pre] * (1 - x[job_pre, d])]) <= 0
        model += pulp.lpSum([Q_cum[job_ante, d] - 0.0001 * z[job_ante, d]]) >= 0
        model += pulp.lpSum([Q_cum[job_ante, d] - jobs_df.job_hours[job_ante] * z[job_ante, d]]) <= 0
        model += pulp.lpSum([x[job_pre, d]] - z[job_ante, d]) >= 0
print("Precedence constraints defined and started solver at %s..." % str(datetime.datetime.now().time()))

model.solve()
print("Problem solved with status %s at %s" % (pulp.LpStatus[model.status], str(datetime.datetime.now().time())))

total_cost = pulp.value(model.objective)
filedirectory = "C:/Users/gurkaali/Documents/Info/Ben/Case Study/Outputs/Plan Capacity Output_"
filename = str(datetime.datetime.now().date()) + '_' + str(datetime.datetime.now().time())[0:2] + '_' + str(datetime.datetime.now().time())[3:5] + '_' + str(datetime.datetime.now().time())[6:8] + ".csv"
w = open(filedirectory + filename, "w")
w.write("The total cost is: ," + str(total_cost) + "\n")
w.write("employee" + "," + "job" + "," + "day_week" + "," + "hours_assigned\n")
for j in df6.index.levels[1]: #jobs
    emmi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique()
    for i in emmi: #employees
        dayyi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
        for d in dayyi: #days
            #print('job, employee, day: ', j, i, d)
            w.write(i + "," + j + "," + str(d) + "," + str(emp_hours[(i,j,d)].varValue) + "\n")
            #print(i, j, d, emp_hours[(i,j,d)].varValue)
w.close()

print("The total cost is: ", total_cost)