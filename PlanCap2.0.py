#!/usr/bin/env python
# -*- coding: latin-1 -*-
import sys
sys.path.append(r"C:\Python27\Lib\site-packages\PuLP-1.5.4-py2.7.egg\pulp")
import ctypes #to show messageboxes. Comes with Python. Works ONLY under windows
##import sys #required to cancel script when an input error is detected
from xlrd import * #to read excel file
#import Tkinter, tkFileDialog, tkMessageBox #for user interface
from tkinter import filedialog, Tk
import os # for file directories etc

root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
book = open_workbook(file_path)
book_directory = os.path.dirname(file_path) #retrieve the directory of excel file. The output file will be exported here later on

sheet0 = book.sheet_by_index(0)

nbr_jobs = int(sheet0.cell(0,20).value) #Nbr of jobs entered is retrieved from the cell where it is calculated via formula
print("Number of jobs: %d " %nbr_jobs)

#List of Jobs - INPUT
Jobs = range(1,nbr_jobs+1) # JobNum_Dic below keeps info on which Job No corresponds to which Job Name

JobNum_Dic = {} # Create the dictionary where job numbers are the key and names are assigned values.
NumJob_Dic = {} # Create the dictionary where names are the key and job numbers are assigned values.

# Set the period list
if sheet0.cell(5,23).ctype != XL_CELL_EMPTY:
    Num_Periods = int(sheet0.cell(5,23).value)+1
    Periods = range(1,Num_Periods)
else:
    msgbox = ctypes.windll.user32.MessageBoxA
    ret = msgbox(None, "You have not set the number of periods." + "\n" + "Aborting now", "Ooops!", 0)
    print(ret)
    sys.exit()
    
#Dictionary of Job Workhour Requirements - INPUT
JobTime = {}

#Job Deadlines and earliest start dates - INPUT
T = {}

#Dictionary for the qualification matrix - INPUT
Qualif = {}

# Dictionary of Jobs where the max number of employees assigned has to be limited
EmpLimitedJobs_Dic = {}

# Dictionary of Job Categories required for the Equipment constraint where key is the job category and value is the job
CategoryJob_Dic = {}
# Dictioanary of Equipment Numbers where key is the job cat. and value is the number of equipment
CatNbrEquip_Dic = {}

# Feed input objects
for i in range(1, nbr_jobs+1):
    #First check whether all job-related info is entered
    if sheet0.cell(i,1).ctype == XL_CELL_EMPTY or sheet0.cell(i,2).ctype == XL_CELL_EMPTY or sheet0.cell(i,3).ctype == XL_CELL_EMPTY:
        msgbox = ctypes.windll.user32.MessageBoxA
        ret = msgbox(None, "There is missing information on job given in Excel sheet row: %d" %(i+1) + "\n" + "Aborting now", "Ooops!", 0)
        print(ret)
        sys.exit()
    elif sheet0.cell(i,4).ctype == XL_CELL_EMPTY:
        msgbox = ctypes.windll.user32.MessageBoxA
        ret = msgbox(None, "At least one employye must be assigned to the job given in Excel sheet row: %d" %(i+1) + "\n" + "Aborting now", "Ooops!", 0)
        print(ret)
        sys.exit()  
    else:
        JobNum_Dic[i] = sheet0.cell(i,0).value # register the matching job name for job no i
        NumJob_Dic[sheet0.cell(i,0).value] = i # register the matching job no for job name
        JobTime[i] = sheet0.cell(i,3).value # register manhour requirement for job no i
        St_Date = int(sheet0.cell(i,1).value) #earliest start period
        End_Date = int(sheet0.cell(i,2).value) #latest finish period
        T[i] = (St_Date,End_Date) #register earlist start and latest finish periods
        Qualif[i] = sheet0.row_values(i,4,12) #includes blank strings when less than 8 employees are assigned to job i
        for x in Qualif[i]:
            while True:
                try:
                    Qualif[i].remove("")
                except ValueError:
                    break
        #Qualif[i] = filter(len, Qualif[i]) # remove blank strings from the Qualification Matrix
        if sheet0.cell(i,12).ctype != XL_CELL_EMPTY: #Enter the max number of employees for the job if any
            EmpLimitedJobs_Dic[i] = int(sheet0.cell(i,12).value)
        if sheet0.cell(i,13).ctype != XL_CELL_EMPTY: #Enter the job category if any
            Job_Category = str(sheet0.cell(i,13).value)
            if Job_Category not in CategoryJob_Dic.keys():
                CategoryJob_Dic[Job_Category]=set([])
            CategoryJob_Dic[Job_Category].add(i)

#Nbr of equipments available for each job category
if sheet0.cell(14,23).ctype != XL_CELL_EMPTY:
    NbrJobCat = int(sheet0.cell(14,23).value)
    for h in range(17, NbrJobCat+18-1):
        CatNbrEquip_Dic[str(sheet0.cell(h,23).value)] = int(sheet0.cell(h,24).value)

# Job Precedencies - INPUT
Pre = set([])
Nbr_Pre_Constraints = int(sheet0.cell(12,23).value) #number of precedence constraints are count on excel in cell V13
for y in range(1,Nbr_Pre_Constraints+1):
    first_job = sheet0.cell(y,16).value
    first_job_no = NumJob_Dic[first_job]
    second_job = sheet0.cell(y,17).value
    second_job_no = NumJob_Dic[second_job]
    Pre.add((first_job_no,second_job_no)) #add the precedent-preceeding jobs tuple to the set

# set the number of daily periods at the beginning of the planning period
if sheet0.cell(6,23).ctype == XL_CELL_EMPTY:
    Nbr_DailyPeriods = 0
else:
    Nbr_DailyPeriods = int(sheet0.cell(6,23).value)

# set the number of 4KW periods at the end of the planning period
if sheet0.cell(7,23).ctype == XL_CELL_EMPTY:
    Nbr_4KWPeriods = 0
else:
    Nbr_4KWPeriods = int(sheet0.cell(7,23).value)

# set regular capacity
if sheet0.cell(2,23).ctype == XL_CELL_EMPTY:
    msgbox = ctypes.windll.user32.MessageBoxA
    ret = msgbox(None, "You must enter regular capacity of working hours" + "\n" + "Enter 0 if none" + "\n" + "Aborting now", "Ooops!", 0)
    print(ret)
    sys.exit()
else:
    WH_Daily = int(sheet0.cell(2,23).value)

# set overtime capacity
if sheet0.cell(3,23).ctype == XL_CELL_EMPTY:
    msgbox = ctypes.windll.user32.MessageBoxA
    ret = msgbox(None, "You must enter overtime capacity of working hours" + "\n" + "Enter 0 if none" + "\n" + "Aborting now", "Ooops!", 0)
    print(ret)
    sys.exit()
else:
    OT_Daily = int(sheet0.cell(3,23).value)

# ------------------------------------------------------------ END OF EXCEL INPUT ----------------------------------------------------------------------
sys.path.append("C:\Python27\Lib\site-packages\PuLP-1.5.4-py2.7.egg\pulp") #W/O this no module named pulp error popos up :S
from pulp import *
from datetime import datetime #necessary to print time

print(str(datetime.now()) + " - model has been initiated")

# ---------------------------------------------- QUALIF2: REVERSED QUALIFICATION MATRIX ------------------------------------
Qualif2 = {} #As opposed to Qualif, Qualif2 has employee j as key and the list of jobs assigned to j as value
for i in Qualif: #for every job i in Qualif keys
    for j in Qualif[i]: # for every employee j assigned to job i
        if j in Qualif2: # if employee j is already among Qualif2 keys, assign job j to employee i in Qualif2
            Qualif2[j].append(i) 
        else: # if employee j is not among Qualif2 keys, add to Qualif2 then assign job j to employee i in Qualif2
            Qualif2[j]=[]
            Qualif2[j].append(i)

# ------------------------ SET ARTIFICIAL COSTS TO PRIORITIZE EARLIER COMPLETION AND EMPLOYEES OF PREFERENCE IN A QUALIFICATION MATRIX -----------
#The first employee assigned to a job in period 1 has a cost 10 units/hr; the cost increases 0.01 units/hr for every period
# AND 2 units/hr for the employee at the lower priority level
QualifCost = {}
for i in Qualif:
    QualifCost[i]={}
    for j in Qualif[i]:
        QualifCost[i][j]={}
        for t in Periods:
            QualifCost[i][j][t]={}
            FirstCost = 10 + 2*(Qualif[i].index(j))
            QualifCost[i][j][t]=FirstCost+(t-1)*0.01

#- OVERTIME -  
#The first employee assigned to a job in period 1 has a cost 40 units/hr; the cost increases 0.01 units/hr for every period
#AND 2 units/hr for the employee at the lower priority level
QualifCostOT = {}
for i in Qualif:
    QualifCostOT[i]={}
    for j in Qualif[i]:
        QualifCostOT[i][j]={}
        for t in Periods:
            QualifCostOT[i][j][t]={}
            FirstCostOT = 40 + 2*(Qualif[i].index(j))
            QualifCostOT[i][j][t]=FirstCostOT+(t-1)*0.01
# ------------------------------------------------------------- SET CAPACITY ----------------------------------------------------           
#Employee capacity for the next 7days + following 54 KWs - INPUT / TO BE MADE FLEXIBLE FOR PLANT SHUT DOWNS, SICK LEAVES ETC.
K = {}
for j in Qualif2:
    K[j]={}
    for t in Periods:
        if t < Nbr_DailyPeriods : #if t corresponds to one of the initial DAILY periods, the capacity will be daily as well
            K[j][t] = WH_Daily
        elif t > Num_Periods - Nbr_4KWPeriods: #if t corresponds to one of the last MONTHLY(4 KW) periods, the capacity will be monthly as well
            K[j][t] = WH_Daily * 20
        else: # rest of the periods are taken weekly (5 working days)
            K[j][t] = WH_Daily * 5

#- OVERTIME - INPUT - / TO BE MADE FLEXIBLE FOR PLANT SHUT DOWNS, SICK LEAVES ETC.
KOT = {}
for j in Qualif2:
    KOT[j]={}
    for t in Periods:
        if t < Nbr_DailyPeriods : #if t corresponds to one of the initial DAILY periods, the capacity will be daily as well
            KOT[j][t] = OT_Daily
        elif t > Num_Periods - Nbr_4KWPeriods: #if t corresponds to one of the last MONTHLY(4 KW) periods, the capacity will be monthly as well
            KOT[j][t] = OT_Daily * 20
        else: # rest of the periods are taken weekly (5 working days)
            KOT[j][t] = OT_Daily * 5

print(str(datetime.now()) + " - parameters defined")

# --------------------------------------------------------- DECLARE PROBLEM OBJECT -----------------------------------------
# Create the "prob" variable to contain the problem data
prob = LpProblem("The Capacity Planning Problem", LpMinimize)

# -------------------------------------------- DECLARE WORKING HOUR AND OVERTIME VARIABLES ---------------------------------
# A dictionary called "WorkHour_vars" is created to contain the referenced Variables
WorkHour_vars = LpVariable.dicts("WH",(Jobs,Qualif2,Periods),lowBound =0)

#Refine the variables dictionary as not every employee is assigned to every job
#SORUN BURADAAAAAAAAAAAAAAAAAAAAAA
#TEST
for j in Qualif[i]:
    print(Qualif[i])
WorkHour_vars2 = {}
for i in WorkHour_vars:
	WorkHour_vars2[i] = {}
for i in WorkHour_vars:
	for j in WorkHour_vars[i]:
		if j in Qualif[i]:
			WorkHour_vars2[i][j]=WorkHour_vars[i][j]

WorkHour_vars=WorkHour_vars2    # Replace the original variables dic with the new one.

##- OVERTIME
WorkHourOT_vars = LpVariable.dicts("OT",(Jobs,Qualif2,Periods),lowBound =0)

#Refine the variables dictionary as not every employee is assigned to every job
WorkHourOT_vars2 = {}
for i in WorkHourOT_vars:
	WorkHourOT_vars2[i] = {}
for i in WorkHourOT_vars:
	for j in WorkHourOT_vars[i]:
		if j in Qualif[i]:
			WorkHourOT_vars2[i][j]=WorkHourOT_vars[i][j]

WorkHourOT_vars=WorkHourOT_vars2    # Replace the original variables dic with the new one.

print(str(datetime.now()) + " - variables defined")

# ------------------------------------- DECLARE BINARY VARIABLES ON JOB START/JOB END (USED FOR PRECEDENCE CONSTRAINTS) ---------------
#create a set of jobs with a precedence constraint
JobCouples=set([]) 
for couple in Pre:
	for i in couple:
		if i not in JobCouples:
			JobCouples.add(i)

# Binary variables: Xit = 1 if Job i has ended in period t / Zit = 1 if Job i has started in period t
IfStarted = LpVariable.dicts("z",(JobCouples,Periods),0,1,LpInteger)
IfEnded = LpVariable.dicts("x",(JobCouples,Periods),0,1,LpInteger)

# ---------------------------------------- SUM OF REGULAR AND OVERTIME HOURS MUST BE EQUAL TO THE JOB"S MANHOUR REQUIREMENT --------------------
# WHs required to complete each job
for i in Jobs:
    prob += lpSum([WorkHour_vars[i][j][t] for j in Qualif[i] for t in Periods]) + lpSum([WorkHourOT_vars[i][j][t] for j in Qualif[i] for t in Periods]) == JobTime[i] #, "Job WH requirement constraint"

# ------------------------------ HOURS OF EMPLOYEE J ON JOB I IN PERIOD T CANNOT BE GREATER THAN HIS/HER CAPACITY IN PERIOD T ------------------------
# WHs of an employee at time t cannot be greater than his/her capacity in that period
#TEST
for k in K:
    print(k)
for t in Periods:
    for j in Qualif2:
        prob += lpSum([WorkHour_vars[i][j][t] for i in Qualif2[j]]) <= K[j][t] # "Employee K constraint"

##- OVERTIME - 
# WHs of an employee at time t cannot be greater than his/her capacity in that period
for t in Periods:
    for j in Qualif2:
        prob += lpSum([WorkHourOT_vars[i][j][t] for i in Qualif2[j]]) <= KOT[j][t] #, "Employee K constraint"    

# -------------------------------------------- DEADLINE / EARLEST START TIME RELATED CAPACITY CONSTRAINT ----------------------------------
# WH of employee j at time t on job i has to be zero provided that time t is beyond job i"s deadline
for i in Jobs:
    for t in Periods:
        for j in Qualif[i]:
            if t > T[i][1] or t < T[i][0]: #T holds the deadline information. T[i] gives a tuple where T[i][0]=Earliest Start Date, T[i][1]=Deadline
                K2 = 0
            else:
                K2 = K[j][t]
            prob += lpSum([WorkHour_vars[i][j][t]]) <= K2 #, "Deadline related K constraint"

##- OVERTIME - 
# WH of employee j at time t on job i has to be zero provided that time t is beyond job i"s deadline
for i in Jobs:
    for t in Periods:
        for j in Qualif[i]:
            if t > T[i][1] or t < T[i][0]:
                KOT2 = 0
            else:
                KOT2 = KOT[j][t]
            prob += lpSum([WorkHourOT_vars[i][j][t]]) <= KOT2 #, "Deadline related K constraint"

print(str(datetime.now()) + " - constraints other than precedence defined")

## ----------------------------------- EQUIPMENT CONSTRAINT ------------------------------------------------------------
# In order to set equipment constraints, define a binary variable Mit: 1 when work is done on job i in period t
# Create a set of jobs with equipment constraint
JobsWEqConstraint = set([])
for cat in CategoryJob_Dic:
    JobsWEqConstraint.update(CategoryJob_Dic[cat]) #union of sets - adding gives error

IfWorked = LpVariable.dicts("m",(JobsWEqConstraint,Periods),0,1,LpInteger)
for cat in CategoryJob_Dic: #for every job category in CategoryJob_Dic
    for t in Periods:
        prob += lpSum(IfWorked[i][t] for i in CategoryJob_Dic[cat]) <= CatNbrEquip_Dic[cat] # Number of jobs within the same job category in a period must be less than or equal to the number of equipments. So for 4KW periods this may not reflect reality!
for t in Periods:
    for i in JobsWEqConstraint: #for every job with an Equipment constraint
        prob += lpSum([(WorkHour_vars[i][j][t] + WorkHourOT_vars[i][j][t] - IfWorked[i][t]*JobTime[i]) for j in Qualif[i]]) <= 0

# --------------------------------- NUMBER OF EMPLOYEES WORKING TOGETHER ON A JOB CONSTRAINT / IF 1; ONLY ONE EMPLOYEE CAN BE ASSIGNED --------------------------
# NOT VERY MEANINGFUL FOR LONGER THAN A DAY PERIODS - CAN BE MODIFIED WHERE TOTAL NBR OF HOURS OF ALL EMPLOYEES ON JOB i IS EQUAL TO THEIR TOTAL CAP AT PERIOD t
IfEmpWorked = LpVariable.dicts("n",(EmpLimitedJobs_Dic.keys(),Qualif2,Periods),0,1,LpInteger)

IfEmpWorked2 = {}
for i in IfEmpWorked: #for every job i in EmpLimitedJobs_Dic (jobs with number of employee constraint)
    IfEmpWorked2[i] = {}
    for j in IfEmpWorked[i]:
        if j in Qualif[i]:
            IfEmpWorked2[i][j] = IfEmpWorked[i][j]

IfEmpWorked = IfEmpWorked2 # Replace

for t in Periods:
    for i in EmpLimitedJobs_Dic: #for every job with a nbr of employee constraint
        if (EmpLimitedJobs_Dic[i] == 1) and (len(Qualif[i]) == 1): #Nonsense to add constraint for a job that can be done only by one employee
            continue
        else:
            prob += lpSum([IfEmpWorked[i][j][t] for j in Qualif[i]]) == EmpLimitedJobs_Dic[i] #RHS of the equation gives the number of employee value
for t in Periods:
    for i in EmpLimitedJobs_Dic:
        for j in Qualif[i]: #for every employee j assigned to job i
            prob += lpSum([(WorkHour_vars[i][j][t] + WorkHourOT_vars[i][j][t] - IfEmpWorked[i][j][t]*JobTime[i])]) <= 0
#!!!!!!!!!!!!!!!!! WH and OT values to be made equal for jobs done by multiple employees simultaneously.
#!!!!!!!!!!!!!!!!! For now only the number of employees on job i at time t can be constrained. No problem when a single employee has to work.
#!!!!!!!!!!!!!!!!! However for multiple employees there is no guarantee that their WH and OT hours will be equal which implies they work together.
##        if x > 1: #x<len(Qualif[i]) olunca uyari koy!!!!!!!!!!!!!!!!!!!
##                #If number of employees given as constraint is greater than 1, working hours (regular and overtime)
##                #of employees should be the same as well BOTH FOR OVERTIME AND REGULAR WORKING HOURS
##                #For that matter the list employees assigned to job i are put in a list two times in the same order
##                # and every consecutive pair is used in the constraint below:
##                listy = Qualif[i] + Qualif[i]
##                for c in range(0,len(Qualif[i])):
##                    emp1=listy[c]
##                    emp2=listy[c+x-1]
### ERROR: 
####                    prob += lpSum([(WorkHourOT_vars[i][emp1][t] - WorkHourOT_vars[i][emp2][t])]) == 0
##                    prob += lpSum([(WorkHour_vars[i][emp1][t] - WorkHour_vars[i][emp2][t])]) == 0

# -------------------------------------------------- PRECEDENCE CONSTRAINTS -------------------------------------------------------------------- 
# Job A cannot start before Job B ends !!!!!!!!!!!!!!!!!!!!!!!!!! Dictionary"de sorun var!!!!!!!!!!!!!!!!!!!!!!
Q = {} #Dictionary Q hols the information on cumulative workhours completed for every time period for every job in at least one precedence relationship
for job_couple in Pre: #for every job couple with precedence constraint. Pre is a set, Job couple is a tuple
    for i in job_couple:
        if i in Q.keys(): #A job can appear in several precedence tuples. If the cumulative workhours have been already calculated, you can skip to the next job
            continue
        Q[i]={}
        for t in Periods:
            summy = 0
            for k in range(1,t+1): #for every time period before AND at the current time t period. Range(1,t+1) = [1,2,3,...,t] End point is not part of the generated list
                for j in Qualif[i]: # for every employee assigned to Job i
                    summy = summy + WorkHourOT_vars[i][j][k] + WorkHour_vars[i][j][k]
                Q[i][t] = summy #Cumulative sum of WH and OT spent by employees assigned to job i until time t

# Precedence constraints "TO BE MADE FLEXIBLE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Pre_PosteriorJobs = set([])
for i in Pre:
    lavoro = i[1]
    if lavoro not in Pre_PosteriorJobs: #The posteirod job cannot start at time 1. A job can be posterior to multiple jobs. To avoid entering the same constraints several times, check whether you did before
        prob += lpSum(IfStarted[lavoro][1]) == 0
    Pre_PosteriorJobs.add(lavoro)
##prob += lpSum(IfStarted["J3"][1]) == 0
##prob += lpSum(IfStarted["J4"][1]) == 0

#All jobs with precedence constraint and previously entered in Pre are saved in ProcessedJobs for the Output file
ProcessedJobs = set([])
for couple in Pre:
    for i in couple:
        if i in ProcessedJobs:
            continue
        else:
            ProcessedJobs.add(i)

for job_couple in Pre: #for every job precedence couple in Pre
    for t in Periods:
        for i in job_couple:
                prob += lpSum(JobTime[i]-Q[i][t]-JobTime[i]*(1-IfEnded[i][t])) <= 0 #J1-1; J2-1; J1-2; J2-2; J1-3; J2-3; J1-4; J2-4;  
                prob += lpSum(JobTime[i]-Q[i][t]-0.5*(1-IfEnded[i][t])) >= 0 #
                prob += lpSum(Q[i][t]-0.5*IfStarted[i][t]) >= 0 #
                prob += lpSum(Q[i][t]-JobTime[i]*IfStarted[i][t]) <= 0 #               
        if t!=Periods[-1]: #for period t+1 no constraint should be defined. Periods[-1] is equal to the last item in periods
            #Job precedence constraint is entered once both jobs have been processed - Indentation is hence critical!
            Prejob = job_couple[0] #the first job to be completed in the tuple
            Postjob = job_couple[1] # the proceeding job in the tuple
            prob += lpSum(IfEnded[Prejob][t]-IfStarted[Postjob][t+1]) >= 0

print(str(datetime.now()) + " - precedence constraints defined")

#--------------------------------------------- OBJECTIVE FUNCTION -------------------------------------------------
# To calculate the opportunity cost of not making an employee work, create a dictionary where unused capacity of each employee is held for every period:
LostCapacity = {}
for j in Qualif2: #for every employee
    LostCapacity[j]={}
    for t in Periods: #for every period
        lostsummy=0
        for i in Qualif2[j]: #for every job i assigned to employee j
            lostsummy += WorkHour_vars[i][j][t] #lostsummy is the USED capacity of employee in period t
        LostCapacity[j][t] = K[j][t]-lostsummy #K[j][t]: UNUSED capacity of employee j at time t
#Obj function has three parts:
        #1. Cost of regular working hours
        #2. Cost of overtime
        #3. Opportunity cost of not using labor capacity: 150*(1/t)*LostCapacity[j][t] where 150 is a constant multiplied by 1/t to...
        # make the unused capacity of any employee assigned to a job more costly initially. This would force model to get job done earlier by...
        # allocating the work across assigned employees from the start wherever possible
prob += lpSum(QualifCost[i][j][t]*WorkHour_vars[i][j][t] for i in Jobs for j in Qualif[i] for t in Periods) + lpSum(QualifCostOT[i][j][t]*WorkHourOT_vars[i][j][t] for i in Jobs for j in Qualif[i] for t in Periods) + lpSum(150*(1/t)*LostCapacity[j][t] for j in Qualif2 for t in Periods), "Capacity Allocation Across Employees Through Time"
print(str(datetime.now()) + " - objective function created")
#------------------------------------------------- WRITE MODEL INTO AN LP FILE -----------------------------------------------
# The problem data is written to an .lp file
prob.writeLP("CapacityModel.lp")

print(str(datetime.now()) + " - lp file has been created")

#------------------------------------------------------------ SOLVE -----------------------------------------------------------------------
# The problem is solved using PuLP"s choice of Solver
prob.solve()
##prob.solve(COIN_CMD())
##prob.solve(GLPK())

# The status of the solution is printed to the screen. In case the solution is ineasible, raise an error and exit python
print("Status:", LpStatus[prob.status])
if LpStatus[prob.status] != "Optimal":
    tkMessageBox.showwarning("Infeasible Solution", "No optimal solution has been found!")
    sys.exit() 

print(str(datetime.now()) + " - Model Solved")

##-------------------------------------------------- CREATE AN OUTPUT TXT FILE -------------------------------------------------------
JobStart = {}
JobEnd={}
JobDone={}
OT={}
WH={}
EmployeeCap={}
for i in Qualif: # for every job
    OT[i]={}
    WH[i]={}
    for j in Qualif[i]: # for every employee j assigned to job i
        OT[i][j]={}
        WH[i][j]={}

for i in ProcessedJobs: #for every job WITH A PRECEDENCE CONSTRAINT 
    JobStart[i]={}
    JobEnd[i]={}

for i in JobsWEqConstraint: #for every job with an equipment constraint
    JobDone[i]={}

#iterate over variables
for v in prob.variables():
    varry = v.name.split("_") # job, period and employee names are separated via underscore. Varry is a list object
    varry[-1] = int(varry[-1]) # last element of the list is for time period which is converted here from string to integer
    varry[1] = int(varry[1]) # second element is for job no which is converted here from string to integer
    if varry[0]=="x" or varry[0]=="z" or varry[0]=="m":
        jobby = varry[1]
        periody = varry[2]
    if varry[0]=="OT" or varry[0]=="WH":
        jobby = varry[1]
        empy = varry[2]
        periody = varry[3]
    if varry[0] == "x": #x is the bin variable to indicate job has been completed
        JobEnd[jobby][periody]=v.varValue

    if varry[0] == "z": #z is the bin variable to indicate job has started
        JobStart[jobby][periody]=v.varValue

    if varry[0] == "m": #m is the bin variable to indicate an employee has worked on job i with an equipment constraint at a given period.
        JobDone[jobby][periody]=v.varValue

    if varry[0] == "OT": #OH is the var to indicate overtime hours of employee j on job i in period t
        OT[jobby][empy][periody]=v.varValue
        
    if varry[0] == "WH": #OH is the var to indicate overtime hours of employee j on job i in period t
        WH[jobby][empy][periody]=v.varValue
        
# ------------------------------- CREATE EMPLOYEE BASED DICTIONARY FOR GRAPHS -------------------------------------------
maxy = 0 #to retrieve the maximum scale to be used in plots
for j in Qualif2: #for every employee
    EmployeeCap[j]=[0]*len(Periods) #Create a list of 0"s. 0"s will be replaced in periods where employee has worked
    for t in Periods:
        empsummy = 0
        for i in Qualif2[j]: #for every job assigned to employee j
            empsummy += WH[i][j][t]+OT[i][j][t]
        EmployeeCap[j][t-1] = empsummy/K[j][t]*100 #EmployeeCap[j] is a list where index starts with 0 by definition. Hours of employee j at time 1 are assigned to the item with index 0 in the list, i.e. the first item
    maxscale = max(EmployeeCap[j])
    if maxscale > maxy:
        maxy = maxscale

#Write variables to a txt file
f = open(book_directory + "/Output.txt", "w")
f.write("Output created at: " + str(datetime.now()) + "\n" + "Periods may be of varying lengths! Initial %d: daily; last %d: 4KW; the remainder: KW" %(Nbr_DailyPeriods, Nbr_4KWPeriods) + "\n")
f.write("\t" + "\t" + "\t") #3 tabs required to align with the rest of the table
# Write period numbers
for t in sorted(Periods):
    f.write(str(t) + "\t")
f.write("\n") #Go to new line
for i in sorted(Jobs):  #WH:
    for j in WH[i]:
        f.write(JobNum_Dic[i]) #f.write(JobNum_Dic[i].encode('UTF-8'))
        f.write('\t') #Write the name of the job
        f.write("WH_J" + str(i) + "\t" + j + "\t")
        for t in sorted(WH[i][j]): #Retrieves values following the period order
            f.write(str(WH[i][j][t]) +"\t")
        f.write("\n")

for i in sorted(Jobs): #OT:
    for j in OT[i]:
        f.write(JobNum_Dic[i]) #.encode("UTF-8") + "\t") #Write the name of the job
        f.write("OT_J" + str(i) + "\t" + j + "\t")
        for t in sorted(OT[i][j]): #Retrieves values following the period order
            f.write(str(OT[i][j][t]) +"\t")
        f.write("\n")

for i in JobStart:
    f.write(JobNum_Dic[i]) #.encode("UTF-8") + "\t") #Write the name of the job
    f.write("JobStart_J" + str(i) + "\t" + "\t") #2 tabs to align with OH and WH variables
    for t in sorted(JobStart[i]):
        f.write(str(JobStart[i][t]) +"\t")
    f.write("\n")

for i in JobEnd: #2 tabs to align with OH and WH variables
    f.write(JobNum_Dic[i]) #.encode("UTF-8") + "\t") #Write the name of the job
    f.write("JobEnd_J" + str(i) + "\t" + "\t")
    for t in sorted(JobEnd[i]):
        f.write(str(JobEnd[i][t]) +"\t")
    f.write("\n")

for i in JobDone: #2 tabs to align with OH and WH variables
    f.write(JobNum_Dic[i]) #.encode("UTF-8") + "\t") #Write the name of the job
    f.write("JobDone_J" + str(i) + "\t" + "\t")
    for t in sorted(JobDone[i]):
        f.write(str(JobDone[i][t]) +"\t")
    f.write("\n")

f.close()
print(str(datetime.now()) + " - Output exported to file")

import pylab as p
import numpy as np
import math

#To squeeze charts, omit periods where no action is planned, if any
for t in reversed(sorted(Periods)):
    sumi = 0
    for i in WH:
        for j in WH[i]:
            sumi += WH[i][j][t]
    if sumi != 0:
        last_period = t
        break
    else:
        continue

fig, axs = p.subplots(int(math.ceil(len(Qualif2.keys())/2.0)),2, figsize=(len(Periods)*0.75,9)) #number of subplotrows is the rounded up number of employees/2.
# In order for math.ceil to work properly number must be divided by 2.0 a float otherwise may rounddown instead. Then the float result must be converted to int for mathpotlib
fig.subplots_adjust(hspace = 2, wspace=0.05) #put space between subplots
axs = axs.ravel()
fig.suptitle("Capacity Utilization", fontsize=18)
ind = np.arange(1,last_period+1) #prepare the x-axis
for i in range(len(Qualif2.keys())):
    emppp=list(Qualif2.keys())[i] #Python 3.X does not let dic keys to be indexed. Convert keys to list instead
    y = np.array(EmployeeCap[emppp][:last_period]) #Hrs of employee data are retrieved from EmployeeCap. The last periods where no action exists are omitted
    axs[i].bar(ind, y, facecolor="#E40045",align="center")
    axs[i].set_ylabel("%", fontsize=8)
    axs[i].set_xlabel("Period", fontsize=8)
    axs[i].set_title(emppp,fontstyle="italic")
    axs[i].set_xticks(ind)
    axs[i].tick_params(axis="both",which="major",labelsize=8)
    axs[i].axhline(y=100, color="#E40045",linestyle="dotted")
fig.tight_layout()
p.savefig(book_directory + "/Employee Capacity.pdf", orientation="landscape", bbox_inches=None)

y_Gantt = [] # for y axis
c_Gantt = []
x_start = []
x_end = []
print("starting to create Gantt Chart")
for i in Jobs: # for every job
    x_end_temp  = []
    WH_temp = []
    t = 1
    #collect start times for job i
    for t in sorted(Periods[:last_period]): #iterate over periods in sorted order
        merda = 0
        cazzo = 0
        for j in WH[i]:
            hrWH = WH[i][j][t]
            merda = merda + hrWH
        for j in OT[i]:
            hrOT = OT[i][j][t]
            cazzo = cazzo + hrOT
        WH_temp.append(cazzo+merda)
    for g in sorted(range(0, last_period)):
        if g == 0 and WH_temp[g] > 0:
            x_start.append(g) #as start period, t-1 is given
            c_Gantt.append(i)
            y_Gantt.append(i)
        elif g > 0 and WH_temp[g] > 0 and WH_temp[g-1] == 0:
            x_start.append(g)
            c_Gantt.append(i)
            y_Gantt.append(i)
    for g in reversed(sorted(range(0, last_period))):
        if g == last_period-1 and WH_temp[g] > 0:
            x_end_temp.append(g+1) #as end period, t is given
        elif g < last_period-1 and WH_temp[g] > 0 and WH_temp[g+1] == 0:
            x_end_temp.append(g+1)
    x_end_temp.reverse()
    x_end = x_end + x_end_temp
fig2 = p.figure()
p.rcParams["font.sans-serif"] = ["Georgia"] #Change font all over chart
p.grid(True)
color_dic_temp = {0:"#FFCD00", 1:"#E40045", 2:"#4B96D2", 3:"#CAD0D8", 4:"#6E7E8D", 5:"#303C49", 6:"#91B900", 7:"#AFAFAF", 8:"#303C49"}
color_dic = {}
from random import uniform
for i in Jobs:
    color_dic[i] = color_dic_temp[i%9]

color_mapper = np.vectorize(lambda x: color_dic.get(x))        
p.hlines(y_Gantt, x_start, x_end, colors=color_mapper(c_Gantt), linewidth=9.0)
p.title("Gantt Chart")
p.xlabel("Periods (Period Lengths: Initial %d: daily; last %d: 4KW; the remainder: KW)" %(Nbr_DailyPeriods, Nbr_4KWPeriods), fontsize=6)
p.ylabel("Job #", fontsize=6)
p.ylim([0,len(Jobs)+1])
p.xlim([1,last_period+1])
##p.annotate("DENEME", xy=(0.05,1/32.0), xycoords="axes fraction")
##txt = """
##    Lorem ipsum dolor sit amet, consectetur adipisicing elit,
##    sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
##    Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris
##    nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in
##    reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla
##    pariatur. Excepteur sint occaecat cupidatat non proident, sunt in
##    culpa qui officia deserunt mollit anim id est laborum."""
##fig2.text(.1,-1,txt)
p.xticks(np.arange(0, last_period+1, 5.0),fontsize=6)
p.yticks(np.arange(1, len(Jobs)+1, 1.0),fontsize=6)
p.savefig(book_directory + "/Gantt Chart.pdf", orientation="landscape", bbox_inches=None)

print(str(datetime.now()) + " - Charts created")
   


