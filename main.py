import csv
import pandas as pd
from datetime import datetime

def write_to_csv(data, filename):
    try:
        with open(filename, mode='w', newline='') as file:
            writer =  csv.writer(file)        
            writer.writerows(data)
    except Exception as e:
        print(f'Error: {e}')   
    except:
        print('')         

        
def get_employees(file_path):
    try:
        df = pd.read_excel(file_path)
        employees=dict(zip(df['Employee_EmailID'], df['Employee_Name']))
        return employees
    except Exception as e:
        print(f'Error: {e}')   
    except:
        print('Error fetching employees')
    

def get_prev_years_pairs(file_path):
    try:
        df = pd.read_excel(file_path)
        previous_assignments = dict(zip(df['Employee_EmailID'], df['Secret_Child_EmailID']))
        return previous_assignments
    except Exception as e:
        print(f'Error: {e}')   
    except:
        print('Error fetching prev year pairs')

    

def assign_pairs(employees,previous_year_pairs):
    employees_list=list(employees.keys())
    remaining_employees=list(employees.keys())
    pairs=[['Employee_Name', 'Employee_EmailID', 'Secret_Child_Name','Secret_Child_EmailID']]

    for employee in employees_list:
        valid_pairs=[e for e in remaining_employees if e != employee and e!=previous_year_pairs.get(employee)]
        if not valid_pairs:
            raise Exception("No valid pair found")
        valid_pair =  valid_pairs[0]
        remaining_employees.remove(valid_pair)
        
        pairs.append([employees[employee],employee,employees[valid_pair],valid_pair])
        
    return pairs       
    


class SecretSanta:

    def __init__(self,employees_path,previous_year_pairs_path):
        self.employees_path=employees_path
        self.previous_year_pairs_path=previous_year_pairs_path

    def run(self):
        employees = get_employees(self.employees_path)
        previous_year_pairs =get_prev_years_pairs(self.previous_year_pairs_path)
        data =assign_pairs(employees,previous_year_pairs)
        write_to_csv(data, f"Secret_Santa_Assignments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")


        


if __name__=="__main__":
    file=SecretSanta(employees_path='Employee-List.xlsx',previous_year_pairs_path='Secret-Santa-Game-Result-2023.xlsx') ## timestamped filename
    file.run()
