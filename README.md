# Intro
This is a task scheduling program that prioritizes different design / manufacturing tasks with the objective of keeping the costs minimal while respecting the constraints below.

A prior version of this has been used in a bus assembly plant. I have updated that for publishing as an Operations Research Case Study.

While the real life example that triggered the creation of this program takes place in a manufacturing environment, it can be applied to almost any projet management / scheduling job

Between the 2 python files, PlanCap2.0.py is an earlier version developed by traditional python objects such as dicts. Capacity with PANDAS.py makes use of pandas for storing variables which makes the model execution faster. Capacity with PANDAS.py is the actual model in use.

## Objective Function
Complete tasks with minimum cost while respecting the constraints. Note that the cost used in the program is a pseudo-cost i.e. it is used to prioritize tasks in a way that regular work is preferred over overtime, working in earlier days over later, assignment of work to more qualified employees etc.

## Constraints
1. Tasks may have a precedence prerequisite i.e. to be able to start task B, task A should have ended
2. Each task must be finished before a given deadline i.e. latest finish time
3. Tasks may have an earliest start time due to procurement or organization related issues
4. Each task can only be accomplished by a certain group of employees.
5. Among the employees qualified for a certain task, there is an order of preference.
6. There is a legal limit on the regular working hours which applies for each employee
7. There is a legal limit on the overtime working hours which applies for each employee
8. Certain tasks need to be done by a single employee from start to end e.g. some design tasks that cannot be divided into modules.

## Parameters
1. Earliest Start Time for each task
2. Latest End Date (deadline) foes each task
3. Qualification Matrix of tasks vs employees
4. Order of tasks
5. Cost of one hour of regular work
6. Cost of one hour of overtime work
7. Whether a task is supposed to be done completely by the same employee from start to end

The outcome of the model is the number of hours each employee works on each task for a given planning period

## Installation
Use Python along with the following modules: pandas / pulp / numpy / datetime / xlrd / csv

Among the modules above pulp is the one enabling linear optimization. It is an open source alternative to software such as cplex and gams. To get the best level of efficiency, it should be used with pandas. However, the built in solver of pulp (https://projects.coin-or.org/Cbc) would not be enough after a certain number of variables. This program makes use of cplex solver.

xlrd module is to read the input parameters from spreadsheet which makes running different models easier. Should you need the file, feel free to contact me. To give you an idea, it looks like the following:
![](2019-05-30-16-53-03.png)

As the default solver has limited performance, I strongly suggest an alternative. I use cplex solver here which significantly improves the performance. However think about eliminating redundant variables first before investing in a proprietary solution.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Licence
[MIT](https://choosealicense.com/licenses/mit/)



