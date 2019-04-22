# Intro
This is an ongoing work for task scheduling:
This program aims at prioritizing different manufacturing tasks with the objective of keeping the costs minimal while respecting the following constraints:
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
Use Python along with the following modules: pandas / pulp / numpy / datetime / csv / xlrd

Among the modules above pulp is crucial for solving the linear optimization. It is an open source alternative to  it software such as cplex and gams. To get the best level of efficiency, it should be used with pandas.

xlrd module is to read the input parameters from spreadsheet which makes running different models easier. SHould you need it, feel free to contact me.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Licence
[MIT](https://choosealicense.com/licenses/mit/)



