# Intro
This is task scheduling program that prioritizes different manufacturing tasks with the objective of keeping the costs minimal while respecting the constraints below.

A prior version of this has been used in a bus assembly plant which I am now updating for publishing as an Operations Research Case Study.

While the real life example that triggered the creation of this program takes place in a manufacturing environment, it can be applied to almost any projet management / scheduling job

Between the 2 python files, PlanCap2.0.py is an earlier version developed by traditional python objects such as dicts. Capacity with PANDAS.py makes use of pandas for storing variables which makes the model execution faster. This last file is the actual model being developed.

## Constraints
1. Tasks may have a precedence prerequisite i.e. to be able to start task B, task A should have ended
2. Each task must be finished before a given deadline i.e. latest finish time
3. Tasks may have an earliest start time due to procurement or organization related issues
4. Each task can only be accomplished by a certain group of employees.
5. Among the employees qualified for a certain task, there is an order of preference.
6. There is a legal limit on the regular working hours which applies for each employee
7. There is a legal limit on the overtime working hours which applies for each employee

## Parameters
1. Earliest Start Time for each task
2. Latest End Date (deadline) foes each task
3. Qualification Matrix of tasks vs employees
4. Order of tasks
5. Cost of one hour of regular work
6. Cost of one hour of overtime work

The outcome of the model is the number of hours each employee works on each task for a given planning period

## Installation
Use Python along with the following modules: pandas / pulp / numpy / datetime / xlrd

Among the modules above pulp is crucial for solving the linear optimization. It is an open source alternative to  it software such as cplex and gams. To get the best level of efficiency, it should be used with pandas.

xlrd module is to read the input parameters from spreadsheet which makes running different models easier. SHould you need it, feel free to contact me.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Licence
[MIT](https://choosealicense.com/licenses/mit/)



