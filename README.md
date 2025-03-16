# Secret Santa Assignment Generator



## Table of Contents
1. [Overview](#overview)
2. [Installation](#installation)
3. [Running the program](#running-the-program)
4. [Code Explanation](#code-explanation)

## Overview
This script generates Secret Santa pairings for employees while ensuring that no employee is paired with themselves or their previous year's Secret Santa recipient. The assignments are saved as a CSV file with a timestamped filename.


## Installation

### Prerequisites
- Python 
  
You can install the dependencies with :
```bash
pip install -r requirements.txt
```

## Running the program

To run the program, navigate to the root directory of the project and execute the following command:
```bash
python main.py
```
 
## Code Explanation 
- Modularized the code into dedicated helper functions that streamline the process of extracting and organizing the necessary data:
  - The *get_employees* function extracts the employee names and emails from the input Excel file and stores them in a dictionary format where Employee_EmailID is the key and Employee_Name is the value.
 - The *get_prev_years_pairs* function extracts the previous year's pairs from the provided Excel file and returns them as a dictionary where Employee_EmailID is the key and Secret_Child_EmailID is the value.
- The *assign_pairs* function, which contains the core logic of the Secret Santa assignment:
  - The function iterates through the list of employees, ensuring that a valid secret child is assigned to each employee while meeting the following conditions:
    - The employee cannot be assigned to themselves.
    - The employee cannot be assigned the same secret child from the previous year (if applicable).
- A list of remaining employees is maintained and updated dynamically, ensuring that once a secret child is assigned, they cannot be selected again.
- If a valid pair is not found for any employee, an exception is raised, preventing invalid assignments and maintaining the integrity of the process.

