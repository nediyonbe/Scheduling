#################################################### READ ME!!!! ###################################################
# nediyonbe - 28/05/19
# Cost is rather a fictious parameter used to prioritize tasks e.g. cost of working on day1 is less than day2 so that 
# the task is done continuously. The cost is also used to prioritize people according to their qualification.
# Labor cost changes in function of the day AND the job.
# The cumulative costs used for precedence constraint gets the sum of working hours of only the related employees: 
# Q_cum dictionary has
# the cost only for the relevant days and employees. This renders the model more efficient
# time range you use should be coherent in different parts of the code such as:
# - where you declare emphour variables
# - where you declare precedence variable
####################################################################################################################
import pulp
import pandas as pd
import numpy as np
import datetime
import csv
import xlrd

# read jobs from the xlsx file
xl_wb = xlrd.open_workbook("C:/Users/gurkaali/Documents/Info/Ben/Case Study/InputFile.xlsx")
xl_sht = xl_wb.sheet_by_index(0)

jobs_df = pd.DataFrame()
for c in range(1, xl_sht.ncols):
    if len(xl_sht.cell(2, c).value) == 0: #this cell onwards no input exists
        break
    jobs_df = jobs_df.append([[xl_sht.cell(2, c).value, int(xl_sht.cell(3, c).value), int(xl_sht.cell(4, c).value), int(xl_sht.cell(5, c).value), xl_sht.cell(6, c).value]], ignore_index=True)
jobs_df.columns = ['jobs', 'earliest_start', 'latest_end', 'job_hours', 'single_employee']
jobs_df.set_index('jobs', inplace = True)

jobs = list(jobs_df.index.values)

# Number of max workdays to plan is given by the user input
schedule_end = int(xl_sht.cell(20, 1).value) + 1
days = list(range(1, schedule_end))

def read_from_xlsx(listy, start_column, data_row):
    for c in range(start_column, xl_sht.ncols):
        if len(str(xl_sht.cell(data_row, c).value)) == 0: #this cell onwards no input exists
            break
        listy.append(xl_sht.cell(data_row, c).value)

employees = []
read_from_xlsx(employees, 1, 11)

employees_design = []
read_from_xlsx(employees_design, 1, 13)

employees_turning = []
read_from_xlsx(employees_turning, 1, 14)

employees_milling = []
read_from_xlsx(employees_milling, 1, 15)

employees_assembly = []
read_from_xlsx(employees_assembly, 1, 16)

employees_measurement = []
read_from_xlsx(employees_measurement, 1, 17)

#Employees by tasks in order of their profeciency - except Design where everyone is equally qualified
# varying_proficiency_across_team: 
# 1 - Assign employees starting with the most proficient one as they have different levels of proficiency
# 0: otherwise
# base_for_qualified: symbolic 1st day cost for qualified employees
# base_for_unqualifieds: symbolic 1st day cost for unqualified employees pushing the model to assign to qualified people
# increment_by_day: symbolic increment in cost by each time period pushing the model to assign first to earliest period possible
# eg. the 1st person in the employees_turning list has a 1st day cost of 10 incrementing till 10.05 at day 500
#     the 2nd person in the employees_turning list has a 1st day cost of 11 incrementing till 11.05 at day 500
#     the last day cost of better qualified person must be below the 1st day cost of the less qualified person
def add_qualification_costs(listy, varying_proficiency_across_team, base_for_qualified, base_for_unqualifieds, increment_by_day):
    for e in listy:
        costs[j][e] = {}
        for d in days:
            costs[j][e][d] = {}
            if e in listy:
                costs[j][e][d] = (base_for_qualified + varying_proficiency_across_team * listy.index(e)) + d * increment_by_day
            else:
                costs[j][e][d] = base_for_unqualifieds + d * increment_by_day

costs = {}
for j in jobs:
    costs[j] = {}
    if j[-1] == 'D':
        add_qualification_costs(employees_design, 0, 10, 88888, 0.0001) # 0 for Design job having equally qualified employees
    if j[-1] == 'T':
        add_qualification_costs(employees_turning, 1, 10, 88888, 0.0001)
    if j[-1] == 'M':
        add_qualification_costs(employees_milling, 1, 10, 88888, 0.0001)
    if j[-1] == 'A':
        add_qualification_costs(employees_assembly, 1, 10, 88888, 0.0001)
    if j[-1] == 'E':
        add_qualification_costs(employees_measurement, 1, 10, 88888, 0.0001)

# overtime costs will be calculated based on these further below
costs_reformed = {(outerKey, innerKey, innermostKey): values for outerKey, innerDict in costs.items() for innerKey, innestDict in innerDict.items() for innermostKey, values in innestDict.items()}
df_input = pd.Series(costs_reformed).reset_index()
df_input.columns = ['Jobs', 'Employees', 'Days', 'Costs']
df_input.set_index(['Jobs', 'Employees', 'Days'], inplace=True)
df_input = df_input.reorder_levels(['Employees', 'Jobs', 'Days']) #this is to reorder levels and match with the previous version. The loops below follow the previous order

model = pulp.LpProblem("My LP Problem", pulp.LpMinimize)

# variables & parameters for precedence constraints
job_sequences_df = pd.DataFrame()
for c in range(1, xl_sht.ncols):
    if len(str(xl_sht.cell(7, c).value)) == 0: #this cell onwards no input exists
        break
    job_sequences_df = job_sequences_df.append([[xl_sht.cell(7, c).value, xl_sht.cell(9, c).value]], ignore_index=True)
job_sequences_df.columns = ['precedent', 'antecedent']

input_jobs_df = jobs_df.reset_index()[['jobs','earliest_start','latest_end']] 
def find_practical_latest_end(x): #x is a df with distinct jobs on every row along with their one level precedent and antecedent
    not_nan_count = 100
    compteur = 1
    #bring one more level precedents
    while not_nan_count > 0:
        if compteur == 1:
            col_left_on = 'jobs'
        else:
            col_left_on = 'antecedent'+str(compteur-1) # you will use this column in join to bring the antecedent of antecedent
        z = pd.merge(x, job_sequences_df[['precedent','antecedent']], left_on = col_left_on, right_on='precedent', how='left')
        filter_col = [col for col in z if col.startswith('antecedent') or col.startswith('earliest_start') or col.startswith('latest_end') or col == 'jobs']
        z = z[filter_col]
        if compteur == 1:
            col_left_on = 'antecedent'
        else:
            col_left_on = 'antecedent'#+str(compteur-1) # you will use this column in join to bring the earlieast and latest times of antecedent
        z2 = pd.merge(z, jobs_df[['earliest_start', 'latest_end']], left_on = col_left_on, right_index=True, how='left')
        z2.rename(columns={'antecedent':'antecedent'+str(compteur), 
                            'earliest_start_x':'earliest_start'+str(compteur-1), 
                            'latest_end_x':'latest_end'+str(compteur-1), 
                            'earliest_start_y':'earliest_start'+str(compteur), 
                            'latest_end_y':'latest_end'+str(compteur) }, 
                            inplace=True)
        x = z2
        # calculate the number of not-null precedents. You will continue until you hit 0
        not_nan_count = z2['antecedent'+str(compteur)].count()
        compteur += 1
    #Get the earliest among antecedent jobs' latest end times
    filter_antecedent_col = [col for col in z if col.startswith('latest_end')] #columns to check the min across
    x['practical_latest_end'] = x[filter_antecedent_col].min(axis=1).astype(int) #across the row, get the min
    x = x[['jobs', 'practical_latest_end']].drop_duplicates() #a job with different branches of antecedent can have different min's for each branch
    x = x.groupby('jobs').min(axis=0).astype(int) #get the min for every job across col > Necessary for jobs having different antecedent branches
    return x

def find_practical_earliest_start(x): #x is a df with distinct jobs on every row along with their one level precedetn and antecedent
    not_nan_count = 100
    compteur = 1
    #bring one more level precedents
    while not_nan_count > 0:
        if compteur == 1:
            col_left_on = 'jobs'
        else:
            col_left_on = 'precedent'+str(compteur-1) # you will use this column in join to bring the antecedent of antecedent
        z = pd.merge(x, job_sequences_df[['precedent','antecedent']], left_on = col_left_on, right_on='antecedent', how='left')
        filter_col = [col for col in z if col.startswith('precedent') or col.startswith('earliest_start') or col.startswith('latest_end') or col == 'jobs']
        z = z[filter_col]
        if compteur == 1:
            col_left_on = 'precedent'
        else:
            col_left_on = 'precedent'#+str(compteur-1) # you will use this column in join to bring the earlieast and latest times of antecedent
        z2 = pd.merge(z, jobs_df[['earliest_start', 'latest_end']], left_on = col_left_on, right_index=True, how='left')
        z2.rename(columns={'precedent':'precedent'+str(compteur), 
                            'earliest_start_x':'earliest_start'+str(compteur-1), 
                            'latest_end_x':'latest_end'+str(compteur-1), 
                            'earliest_start_y':'earliest_start'+str(compteur), 
                            'latest_end_y':'latest_end'+str(compteur) }, 
                            inplace=True)
        x = z2
        # calculate the number of not-null precedents. You will continue until you hit 0
        not_nan_count = z2['precedent'+str(compteur)].count()
        compteur += 1
    #Get the latest among precedent jobs' earliest start times
    filter_precedent_col = [col for col in z if col.startswith('earliest_start')]
    x['practical_earliest_start'] = x[filter_precedent_col].max(axis=1).astype(int) #across the row, get the max
    x = x[['jobs', 'practical_earliest_start']].drop_duplicates() #a job with different branches of antecedent can have different max's for each branch
    x = x.groupby('jobs').min(axis=0).astype(int) #get the max for every job across col > Necessary for jobs having different antecedent branches
    return x

df_left = find_practical_latest_end(input_jobs_df)
df_right = find_practical_earliest_start(input_jobs_df)
df5 = pd.merge(df_left, df_right, left_index=True, right_index=True)

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

# Create the overtime costs based on regular costs given above: The lowest overtime cost i.e overtime cost of the most skilled person 
# day 1, must be higher than the highest regular cost, i.e. the cost of the least skilled employee the very last day of the planning period
df6['Costs_overtime'] = np.floor(df6['Costs']) * 9000 + np.modf(df6['Costs'])[0]

# DECLARE VARIABLES
emp_hours = pulp.LpVariable.dicts("employee hours",
                                   ((employees, jobs, days) for employees, jobs, days in df6.index),
                                   lowBound=0,
                                   cat='Continuous')

emp_hours_overtime = pulp.LpVariable.dicts("employee hours overtime",
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

s = pulp.LpVariable.dicts("Bin_w for single employee constraints",
                                  ((j, i) for j in jobs_df[jobs_df['single_employee'] == 'YES'].index.values
                                            for i in employees),
                                  cat='Binary')

#get an array of unique jobs with precedence-antecedence relation
jobs_sequenced_unique = np.unique(np.append(job_sequences_df.antecedent, job_sequences_df.precedent))

# # OBJECTIVE FUNCTION
# model += pulp.lpSum([emp_hours[employees, jobs, days] * df6.loc[(employees, jobs, days), 'Costs'] for employees, jobs, days in df6.index])
model += pulp.lpSum([emp_hours[employees, jobs, days] * df6.loc[(employees, jobs, days), 'Costs'] + emp_hours_overtime[employees, jobs, days] * df6.loc[(employees, jobs, days), 'Costs_overtime'] for employees, jobs, days in df6.index])

# CONSTRAINTS
# CONSTRAINTS - Sum of hours spent by an employee on all jobs has to be less than or equal to 8hrs a day (or 45hrs a week)
for d in df6.index.levels[2]: # days
    #first get the employees who are qualified for Job = j, then get the unique list of days for that job
    for i in df6.loc[(df6.index.get_level_values('Days') == d)].index.get_level_values('Employees').unique(): # employees qualified for job j
        #get a list of jobs applicable only in Day d:
        denyo = df6.loc[(df6.index.get_level_values('Days') == d) & (df6.index.get_level_values('Employees') == i)].index.get_level_values('Jobs').unique()
        model += pulp.lpSum([emp_hours[(i, j, d)] for j in denyo]) <= 45 #45 is the weekly working time
print("Employees' work-hour constraints defined at ", str(datetime.datetime.now().time()))

# CONSTRAINTS - Sum of OVERTIME hours spent by an employee on all jobs has to be less than or equal to Xhrs a day (or 5*Xhrs a week)
for d in df6.index.levels[2]: # days
    #first get the employees who are qualified for Job = j, then get the unique list of days for that job
    for i in df6.loc[(df6.index.get_level_values('Days') == d)].index.get_level_values('Employees').unique(): # employees qualified for job j
        #get a list of jobs applicable only in Day d:
        denyo = df6.loc[(df6.index.get_level_values('Days') == d) & (df6.index.get_level_values('Employees') == i)].index.get_level_values('Jobs').unique()
        model += pulp.lpSum([emp_hours_overtime[(i, j, d)] for j in denyo]) <= 20 #20 is the weekly overtime possible


# CONSTRAINTS - Total work done on a job is equal to its hour requirement
for j in jobs_df.index: #jobs
    dayyi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
    emmi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique()
    model += pulp.lpSum([emp_hours[i, j, d] + emp_hours_overtime[i, j, d] for i in emmi #only the employees qualified for the job in iteration
                        for d in dayyi]) >= jobs_df.job_hours[j]
print("Jobs' work-hour constraints defined at ", str(datetime.datetime.now().time()))

# CONSTRAINTS - Precedence
Q_cum = {}
Q_cum_overtime = {} # EKLEDIM
# To be able to use the dataframe cumulative sum function in cumulative hour calculation, convert the variable dictionary to a dataframe first:
emp_hours_df = pd.DataFrame(list(emp_hours.items()), columns = ('key','variable'),index=pd.MultiIndex.from_tuples(emp_hours.keys(), names=('employee', 'job', 'day')))
#info: https://thispointer.com/python-pandas-how-to-create-dataframe-from-dictionary/
# cumulative sum of OVERTIME working hours:
emp_hours_df_overtime = pd.DataFrame(list(emp_hours_overtime.items()), columns = ('key','variable'),index=pd.MultiIndex.from_tuples(emp_hours_overtime.keys(), names=('employee', 'job', 'day')))

# If a job is in precedence relationship with multiple jobs, no need to create Q_cum's each time.
# Create the cumulative hours only for the period it can be done i.e. between earliest start and latest finish dates
# Use the jobs_sequenced_unique where all jobs with precedence relationship are listed only once:
for j in jobs_sequenced_unique: #j is a list of 2 jobs
    Q_cum[j] = {}
    dief_this_job = emp_hours_df.loc[emp_hours_df.index.get_level_values('job') == j]
    dief = dief_this_job.groupby(['job', 'day'])['variable'].agg('sum')
    dief.sort_index(level=['day'], inplace = True) # for the cum sum to work correctly, it must be ordered by day
    print("Q_cum for  job ", j, " defined at ", str(datetime.datetime.now().time()))

    j_earliest = int(df5.loc[j]['practical_earliest_start'])
    j_latest = int(df5.loc[j]['practical_latest_end'])
    dief_cumulus = dief.loc[(dief.index.get_level_values('day') >= j_earliest) & 
        (dief.index.get_level_values('day') <= j_latest)].cumsum() 
        #you had defined a variable for each day but some are reduncant e.g. those before earliest start

    dayyi_amca = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
    for d in dayyi_amca:
        Q_cum[j, d] = dief_cumulus[d - j_earliest]
    # Overtime version below:
    Q_cum_overtime[j] = {}
    dief_this_job_overtime = emp_hours_df_overtime.loc[emp_hours_df_overtime.index.get_level_values('job') == j]
    dief_overtime = dief_this_job_overtime.groupby(['job', 'day'])['variable'].agg('sum')
    dief_overtime.sort_index(level=['day'], inplace = True) # for the cum sum to work correctly, it must be ordered by day
    print("Q_cum_overtime for  job ", j, " defined at ", str(datetime.datetime.now().time()))

    j_earliest = int(df5.loc[j]['practical_earliest_start'])
    j_latest = int(df5.loc[j]['practical_latest_end'])
    dief_cumulus_overtime = dief_overtime.loc[(dief_overtime.index.get_level_values('day') >= j_earliest) & 
        (dief_overtime.index.get_level_values('day') <= j_latest)].cumsum() 
        #you had defined a variable for each day but some are reduncant e.g. those before earliest start

    for d in dayyi_amca:
        Q_cum_overtime[j, d] = dief_cumulus_overtime[d - j_earliest]

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
        model += pulp.lpSum([jobs_df.job_hours[job_pre] - Q_cum[job_pre, d] - Q_cum_overtime[job_pre, d] - 0.0001*(1 - x[job_pre, d])]) >= 0
        model += pulp.lpSum([jobs_df.job_hours[job_pre] - Q_cum[job_pre, d] - Q_cum_overtime[job_pre, d] - jobs_df.job_hours[job_pre] * (1 - x[job_pre, d])]) <= 0
        model += pulp.lpSum([Q_cum[job_ante, d] + Q_cum_overtime[job_ante, d] - 0.0001 * z[job_ante, d]]) >= 0
        model += pulp.lpSum([Q_cum[job_ante, d] + Q_cum_overtime[job_ante, d] - jobs_df.job_hours[job_ante] * z[job_ante, d]]) <= 0
        model += pulp.lpSum([x[job_pre, d]] - z[job_ante, d]) >= 0
print("Precedence constraints defined at %s..." % str(datetime.datetime.now().time()))

# CONSTRAINTS - Single Employee: Some tasks e.g. Design require the assignment of the same employee from start to end
# Important! Q_cum below has been calculated for precedence constraintsa which is also used for single employee constraint
# All jobs requiring single employee are also precedents of other jobs so it is OK.
# Otherwise, one should also calculate Q_cum for these jobs but only at their last day (or time period)
# Similar to the Q_cum calculation above, do the similar thing withy employee as one iof the keys:
Q_cum_petit = {}
Q_cum_petit_overtime = {} 

for j in jobs_df[jobs_df['single_employee'] == 'YES'].index.values:
    j_earliest = int(df5.loc[j]['practical_earliest_start'])
    j_latest = int(df5.loc[j]['practical_latest_end'])
    
    Q_cum_petit[(j, i)] = {}
    Q_cum_petit_overtime[(j, i)] = {} 

    dief_this_job_petit = emp_hours_df.loc[(emp_hours_df.index.get_level_values('job') == j)]
    dief_this_job_petit_overtime = emp_hours_df_overtime.loc[(emp_hours_df.index.get_level_values('job') == j)] 

    dief_petit = dief_this_job_petit.groupby(['job', 'employee', 'day'])['variable'].agg('sum')
    dief_petit.sort_index(level=['employee', 'day'], inplace = True) # for the cum sum to work correctly, it must be ordered by day 

    dief_petit_overtime = dief_this_job_petit_overtime.groupby(['job', 'employee', 'day'])['variable'].agg('sum') 
    dief_petit_overtime.sort_index(level=['employee', 'day'], inplace = True) # for the cum sum to work correctly, it must be ordered by day 

    print("Q_cum_petit for  job / employee ", j, '/ ', i, " defined at ", str(datetime.datetime.now().time()))
    for i in df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique():
        dief_petit_care = dief_petit.loc[dief_petit.index.get_level_values('employee') == i]
        #you had defined a variable for each day but some are reduncant e.g. those before earliest start
        dief_cumulus_petit = dief_petit_care.cumsum()
        dief_cumulus_petit = dief_cumulus_petit.loc[(dief_cumulus_petit.index.get_level_values('day') == j_latest)]
        Q_cum_petit[j, i] = dief_cumulus_petit.item()
        # overtime version below:
        dief_petit_care_overtime = dief_petit_overtime.loc[dief_petit_overtime.index.get_level_values('employee') == i]
        #you had defined a variable for each day but some are reduncant e.g. those before earliest start
        dief_cumulus_petit_overtime = dief_petit_care_overtime.cumsum()
        dief_cumulus_petit_overtime = dief_cumulus_petit_overtime.loc[(dief_cumulus_petit_overtime.index.get_level_values('day') == j_latest)]
        Q_cum_petit_overtime[j, i] = dief_cumulus_petit_overtime.item()

for j in jobs_df[jobs_df['single_employee'] == 'YES'].index.values: #jobs with a single designated employee constraint
    last_day_j = jobs_df.loc[j]['latest_end'] # this must be the practical latest end but it does not matter as design tasks also have precedence contraints
    qualified_empi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique()
    model += pulp.lpSum([s[(j, i)] for i in qualified_empi]) == 1
    for i in qualified_empi:
        model += pulp.lpSum([Q_cum_petit[(j, i)]  + Q_cum_petit_overtime[(j, i)] - 10000000 * s[j, i]]) <= 0 #only the employees qualified for the job in iteration
print("Single employee constraint defined for related tasks and started solver at %s..." % str(datetime.datetime.now().time()))

model.solve(pulp.CPLEX_CMD())
print("Problem solved with status %s at %s" % (pulp.LpStatus[model.status], str(datetime.datetime.now().time())))
#path = 'C:\\Program Files\\IBM\\ILOG\\CPLEX_Studio129\\cplex\\bin\\x64_win64\\cplex.exe'
total_cost = pulp.value(model.objective)
filedirectory = "C:/Users/gurkaali/Documents/Info/Ben/Case Study/Outputs/Plan Capacity Output_"
filename = str(datetime.datetime.now().date()) + '_' + str(datetime.datetime.now().time())[0:2] + '_' + str(datetime.datetime.now().time())[3:5] + '_' + str(datetime.datetime.now().time())[6:8] + ".csv"
w = open(filedirectory + filename, "w")
w.write("The total cost is: ," + str(total_cost) + "\n")
w.write("employee" + "," + "job" + "," + "day_week" + "," + "hours_assigned" + "," + "overtime_hours_assigned\n")
for j in df6.index.levels[1]: #jobs
    emmi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Employees').unique()
    for i in emmi: #employees
        dayyi = df6.loc[(df6.index.get_level_values('Jobs') == j)].index.get_level_values('Days').unique()
        for d in dayyi: #days
            w.write(i + "," + j + "," + str(d) + "," + str(emp_hours[(i,j,d)].varValue) + "," + str(emp_hours_overtime[(i,j,d)].varValue) + "\n")
w.close()

print("The total cost is: ", total_cost)