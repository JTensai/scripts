import pandas as pd
import argparse
import re

# Create an argument parser
parser = argparse.ArgumentParser(description='Process Excel file and save results as CSV.')

# Add the input file argument
parser.add_argument('input_file', help='Path to the input Excel file')

# Parse the command-line arguments
args = parser.parse_args()

# Load the Excel file
try:
    excel_file = pd.read_excel(args.input_file, sheet_name='Sheet1', header=None)
except FileNotFoundError:
    print(f"Error: The file '{args.input_file}' was not found.")
    exit(1)

# General rule of thumb, rows will change, columns hopefully shouldn't

class Employee:
    def __init__(self, id, name, net_pay, default_department, per_program_gross_pay, gross_pay, employee_taxes, deductions, employer_taxes_minus_futa):
        self.id = id
        self.name = name
        self.net_pay = net_pay
        self.default_department = default_department
        self.per_program_gross_pay = per_program_gross_pay
        self.gross_pay = gross_pay
        self.employee_taxes = employee_taxes
        self.deductions = deductions
        self.tax_rate = (gross_pay - net_pay - deductions) / gross_pay
        self.employer_taxes_minus_futa = employer_taxes_minus_futa
    
    def toString(self):
        # Define the string representation
        program_gross_string = "\n\t\t".join([f"{program}: {gross}" for program, gross in self.per_program_gross_pay.items()])
        return f"{self.name} ({self.id})\n\tDefault Department: {self.default_department}\n\tNet Pay: {self.net_pay}\n\tEmployee Taxes: {self.employee_taxes} ({round(self.tax_rate * 100, 2)}%)\n\tEmployer Taxes: {self.employer_taxes}\n\tEmployer FUTA: {self.employer_futa}\n\tDeductions: {self.deductions}\n\tGross Pay: {self.gross_pay}\n\tPrograms:\n\t\t{program_gross_string}"

def extract_number_from_string(input_string):
    # Use regular expression to extract the number
    match = re.search(r'\d+', input_string)

    if match:
        extracted_number = match.group()
        return extracted_number
    else:
        return None

def find_department(row_number, department_start_lines):
    # Sort the department_start_lines dictionary by row number
    sorted_departments = sorted(department_start_lines.items(), key=lambda x: x[1])
    
    # Iterate through the sorted departments to find the correct department
    current_department = None
    for department, start_line in sorted_departments:
        if row_number >= start_line:
            current_department = department
        else:
            break
    
    return current_department

def get_num(row, col):
    if not pd.isna(excel_file.iat[row, col]):
        return round(float(str(excel_file.iat[row, col]).replace(',', '')), 2)
    else:
        return None


# column indecies
program_col = 1             # Employee ID & program hours per employee
grand_total_col = 2         # column where "Grand Tot:" can be found
employee_total_col = 3      # column where "Employee Tot:" can be found
department_title_col = 5    # title of default program
employee_name_col = 6       # employee name
hours_worked_col = 8        # number of hours worked for a program
net_pay_col = 10            # direct deposit / net pay amount
gross_pay_col = 14          # gross pay totals for this pay period
taxes_col = 25              # taxes for this pay period
deductions_col = 34         # deductions for this pay period
futa_col = 38               # column where FUTA text can be found
employer_tax_col = 43       # ER Taxes total

departments_dict = {}
# BUILD DEFAULT DEPARTMENT LOOKUP DICTIONARY
for dept_name_row_index, dept_name_value in enumerate(excel_file.iloc[:, department_title_col]):
    # Check if the cell value contains anything, this signals entry into a new department
    if not pd.isna(dept_name_value):
        # print(f"Found Dept: '{str(dept_name_value)}' at Row index: {dept_name_row_index+1}, Column F")
        dept_name = dept_name_value
        if 'TumbleBunnies' in dept_name_value:
            dept_name = 'Tumblebunny'
        departments_dict[str(dept_name)] = dept_name_row_index

total_futa = 0
employees = []

for employee_name_row_index, employee_name_value in enumerate(excel_file.iloc[:, employee_name_col]):
    # Check if the cell value contains anything, this signals entry into a new department
    if not pd.isna(employee_name_value):
        # print(f"\nFound Employee: '{str(employee_name_value)}' at Row {employee_name_row_index+1}, Column G")
        # get employee name from spreadsheet
        employee_name = employee_name_value
        # get employee number from spreadsheet
        employee_number = extract_number_from_string(str(excel_file.iat[employee_name_row_index, program_col]))
        # get employee net pay / direct deposit amount
        employee_net_pay = get_num(employee_name_row_index + 1, net_pay_col)


        for index_j, cell_value_j in enumerate(excel_file.iloc[employee_name_row_index:, futa_col]):
            if "FUTA" in str(cell_value_j):
                employer_futa = get_num(employee_name_row_index+index_j, employer_tax_col)
                break


        # dictionary with keys being program names and values being gross pay for that program
        employee_per_program_gross_pay = {}
        name_row_offset = 0

        # employee_name_row_index + 2 so that we start on the row of the first program with pay/hours
        first_program_row = employee_name_row_index + 2
        for employee_program_gross_row_index, value in enumerate(excel_file.iloc[first_program_row:, gross_pay_col]):
            current_row = first_program_row + employee_program_gross_row_index

            # add program and gross pay to the dictionary
            employee_program_gross_pay = get_num(current_row, gross_pay_col)
            employee_program_hours = get_num(current_row, hours_worked_col)
            employee_program_name = None

            # check if this is a program row that has payment or hours for this pay period
            if employee_program_gross_pay is not None or employee_program_hours is not None:
                # get program name, it may not be in line with the gross pay for the program if the number of programs is too small for a given employee
                # if it is empty, add a local offset and try again.
                
                # loop a limited number of times looking for the matching program name
                for i in range(101):
                    # if is empty
                    test_cell = excel_file.iat[current_row + name_row_offset, program_col]
                    if pd.isna(test_cell): # if the test cell is empty, add to the offset and test again
                        name_row_offset = name_row_offset + 1
                    else:
                        # we found the program name!
                        employee_program_name = str(test_cell)
                        if "Polka Dots" in employee_program_name:
                            employee_program_name = "Polkadots"
                        break
                    
                    if i > 99:
                        print(f"ERROR: Unable to find program for listed gross pay for {employee_name}: ${employee_program_gross_pay}")
                        break

            if employee_program_hours is None:
                # If we have pay but no hours, it was a bonus of some kind.
                if employee_program_gross_pay is not None:
                    employee_per_program_gross_pay[employee_program_name] = employee_program_gross_pay
            else:
                # But if we have hours
                if employee_program_gross_pay is None:
                    # AND our pay is empty, that means we don't have a pay rate for the hours reported and our employee did not get money they earned.
                    print(f"\n!!! ERROR! {employee_name} has {employee_program_hours} hours reported for {employee_program_name} but earned $0. Check their pay rate for that category.")
                else:
                    # All is good, we have hours, we have pay. Things are great
                    employee_per_program_gross_pay[employee_program_name] = employee_program_gross_pay

            # +1 to the starting row because we know the totals are the row after the last program
            employee_total_cell = excel_file.iat[current_row + 1, employee_total_col]
            if "Employee Tot:" in str(employee_total_cell):
                # grab the gross_pay, taxes, and deductions
                employee_gross_pay = get_num(current_row + 1, gross_pay_col)
                if employee_gross_pay > 0:
                    employee_taxes = get_num(current_row + 1, taxes_col)
                    employee_deductions = get_num(current_row + 1, deductions_col)
                    employer_taxes_minus_futa = get_num(current_row + 1, employer_tax_col)
                            
                    # print(f"\tEmployer Taxes for {employee_name}: {employer_taxes_minus_futa} - {employer_futa} = {employer_taxes_minus_futa - employer_futa}")
                    # futa is billed to us separately so don't include it in the breakdown for the tax withdrawal
                    
                    total_futa += employer_futa
                    # print(f"total_futa: {total_futa}")
                    employer_taxes_minus_futa -= employer_futa

                    # finishing with this employee, grab the stated totals and move to next employee
                    employee = Employee(employee_number, employee_name, employee_net_pay, find_department(current_row, departments_dict), employee_per_program_gross_pay, employee_gross_pay, employee_taxes, employee_deductions, employer_taxes_minus_futa)
                    # print(f"{employee.toString()}")
                    employees.append(employee)
                else:
                    print(f"\nWarning: {employee_name} had no earnings this pay period.")
                break

stated_gross = None
stated_taxes = None
stated_deductions = None
stated_employer_taxes_minus_futa = None
stated_employer_futa = None
stated_net = None
stated_ca_ett = None

grand_total_data_frame = None

# Iterate through the cells in Column C
for index, cell_value in enumerate(excel_file.iloc[:, grand_total_col]):
    # Check if the cell value contains "Grand Tot:"
    if "Grand Tot:" in str(cell_value): 
        stated_gross = get_num(index, gross_pay_col)
        stated_taxes = get_num(index, taxes_col)
        stated_deductions = get_num(index, deductions_col)
        stated_net = get_num(index + 2, gross_pay_col + 3)

        for index_j, cell_value_j in enumerate(excel_file.iloc[index:, futa_col]):
            if "FUTA" in str(cell_value_j):
                stated_employer_futa = get_num(index+index_j, employer_tax_col)
            if "CA ETT" in str(cell_value_j):
                stated_ca_ett = get_num(index+index_j, employer_tax_col)
                break
            
        # we subtract out FUTA here because it 
        stated_employer_taxes_minus_futa = get_num(index, employer_tax_col) - stated_employer_futa
        break












tracked_programs = ['Admin', 'Dance', 'Events', 'Gymnastics', 'Hospitality', 'Polkadots', 'Swim', 'TAG', 'Team', 'Tumblebunny', 'Maintenance']

tax_by_program = {}
net_by_program = {}
gross_by_program = {}
employer_taxes_by_program = {}

for tracked_program in tracked_programs:
    tax_by_program[tracked_program] = 0
    net_by_program[tracked_program] = 0
    gross_by_program[tracked_program] = 0
    employer_taxes_by_program[tracked_program] = 0

# should store deductions per category as well, so we know how to refund the expenses
deducted_by_program = {}

# used to verify all money is accounted for and there are no rounding errors.
calculated_gross_total = 0
calculated_deductions = 0
for employee in employees:
    # print(employee.toString())
    calculated_gross_total += employee.gross_pay
    
    # but for now, just get the total deduction per employee
    calculated_deductions += employee.deductions
    employee.deductions_remaining = employee.deductions

    # Edge case here where there is more deductions than earnings. Flag it so we can handle it if it pops up
    if round(employee.deductions, 2) > round(employee.gross_pay - employee.employee_taxes, 2):
        print(f"ERROR: {employee.name} has more deductions ({employee.deductions}) than available net pay ({employee.gross_pay - employee.employee_taxes}). This is very weird and will need special handling.")


    # will contain the consolidated list of programs instead of ALL pay categories
    employee_mapped_program_gross = {}

    # There are a select few people that need to be exclude or split between programs
    # Nasa Nergui       660735 | Manage = Split evenly
    # Linsay Groom      91844 | Manage = Split evenly
    # Ashley Lewis      95380 | Manage = Split evenly
    # Scott Wilkie      685470 | Manage = Split evenly
    # Khulan Purevjav   693133 | Clean = Split half into each program
    special_case_employees = [660735, 693133, 91844, 95380, 685470]
    
    # Manage hours for special case employees will be split between main programs
    split_programs = ['Events', 'Gymnastics', 'Hospitality', 'Tumblebunny', 'Dance', 'Swim', 'TAG']
    split_amount = 0

    if employee.id == 660735: # NASANJARGAL NERGUI
        split_amount = (employee.per_program_gross_pay["Manage"] * (2 / 3)) / len(split_programs) # 2/3 of managing goes to all programs
        employee.per_program_gross_pay["TAG"] += employee.per_program_gross_pay["Manage"] / 3 # 1/3 of managing goes specifically to TAG (which also has some from the split)
        del employee.per_program_gross_pay["Manage"]

    elif employee.id == 693133: # KHULAN PUREVJAV
        # Move "Clean" hours to "Maintenance"
        employee.default_department = "Maintenance"

    elif employee.id == 91844: # LINDSAY A GROOM
        # Should have specific hours for TEAM and COMP, those are handled as normal
        split_amount = employee.per_program_gross_pay["Manage"] / len(split_programs)
        del employee.per_program_gross_pay["Manage"]

    elif employee.id == 95380: # ASHLEY M LEWIS
        # Should have specific hours for TEAM and COMP, those are handled as normal
        split_amount = employee.per_program_gross_pay["Manage"] / len(split_programs)
        del employee.per_program_gross_pay["Manage"]

    elif employee.id == 685470: # SCOTT A WILKIE
        # Move "Manage" hours to "Maintenance"
        employee.default_department = "Maintenance"
        
    #current problem is that program gross pay is reassigned correctly, but taxes are just shifted to the person's default program
    
    for split_program in split_programs:
        if split_program in employee_mapped_program_gross:
            employee_mapped_program_gross[split_program] += split_amount
        else:
            employee_mapped_program_gross[split_program] = split_amount

    # Loop over all employee programs that the employee had pay in
    for employee_program, employee_program_gross in employee.per_program_gross_pay.items():
        adjusted_program = employee_program

        if employee.id not in special_case_employees: # Everyone else
            if employee_program in tracked_programs:
                adjusted_program = employee_program
            else:
                # Many categories like "Mentor" need to be adjusted into their default program
                map_to_default = ["Manage", "Mentor", "Full Class", "Training", "Overtime", "Private Lessons", "Senior Coach", "Sick", "Split Shift Premium", "Trainer", "Bonus", "Gift Cards or $$"]
                if employee_program in map_to_default:
                    adjusted_program = employee.default_department
                elif ('Camps' in employee_program or 'Kids Night Out' in employee_program):
                    adjusted_program = 'Events'
                elif 'Clean' in employee_program:
                    adjusted_program = 'Maintenance'
                elif 'Team Coach Fee' in employee_program:
                    adjusted_program = 'Team'
                else:
                    print(f"Unhandled program! {employee.name}: {employee_program}")

        # GROSS pay
        if adjusted_program in employee_mapped_program_gross:
            employee_mapped_program_gross[adjusted_program] += employee_program_gross
        else:
            employee_mapped_program_gross[adjusted_program] = employee_program_gross

    # we sort this one so that the programs are set up with values smallest to largets. This makes handling deductions easier.
    employee_mapped_program_gross = dict(sorted(employee_mapped_program_gross.items(), key=lambda item: item[1]))

    deduction_amount_already_applied = 0
    programs_processed = 0
    
    # loop over adjusted program gross totals and calculate taxes and apply deductions
    for adjusted_program, employee_program_gross in employee_mapped_program_gross.items(): 
        if adjusted_program in gross_by_program:
            gross_by_program[adjusted_program] += employee_program_gross
        else:
            gross_by_program[adjusted_program] = employee_program_gross

        # Split tax for each program. Tax rate is per employee, not per program, so this is ok
        # print(f"{employee.name}: {adjusted_program}: {employee_program_gross} | {employee.tax_rate}")
        employee_program_tax = employee_program_gross * employee.tax_rate
        
        if adjusted_program in tax_by_program:
            tax_by_program[adjusted_program] += employee_program_tax 
        else:
            tax_by_program[adjusted_program] = employee_program_tax

        # In order to properly classify employer taxes into their correct program, we need to know how much of the income came from a certain program
        # Calculate employer taxes by program
        program_ratio = employee_program_gross / employee.gross_pay        
        employer_program_tax = employee.employer_taxes_minus_futa * program_ratio
        if adjusted_program in employer_taxes_by_program:
            employer_taxes_by_program[adjusted_program] += employer_program_tax 
        else:
            employer_taxes_by_program[adjusted_program] = employer_program_tax





        # find the amount to take from each program to evenly distribute deductions across programs
        # TODO: if I ever map certain deductions to certain programs, this will need to change
        deductions_split_among_programs = (employee.deductions - deduction_amount_already_applied) / (len(employee_mapped_program_gross) - programs_processed)

        # calculate net pay BEFORE applying deductions
        net_pay_before_deductions = employee_program_gross - employee_program_tax

        if deductions_split_among_programs <= net_pay_before_deductions:
            employee_program_net = net_pay_before_deductions - deductions_split_among_programs

            deduction_amount_already_applied += deductions_split_among_programs
            programs_processed += 1
        else:
            employee_program_net = 0
            # print(f"I KNEW THIS CODE WAS WORTH WRITING! Totally a case where split deduction amount was greater than what was earned in a category. {employee.name}: {adjusted_program}")
            deduction_amount_already_applied += net_pay_before_deductions
            programs_processed += 1



        if programs_processed == len(employee_mapped_program_gross) and deduction_amount_already_applied < employee.deductions:
            print(f"ERROR: Unable to apply all deductions because the last category ({adjusted_program}) didn't have enough money in it to cover the distribution.")
            print(f"\tBecause we sort the dict by program gross earnings, this case should now only occur when deductions are more than net earnings.")
            print(f"\tAmount in last program {adjusted_program}: {employee_program_gross} | still left to deduct: {deductions_split_among_programs}")



        if adjusted_program in net_by_program:
            net_by_program[adjusted_program] += employee_program_net
        else:
            net_by_program[adjusted_program] = employee_program_net



# print(f"\n\nGROSS: Every pay category with non-zero pay:\n")
# for program, gross in gross_by_program.items():
#     print(f"\t{program}: {round(gross, 2)}")


print(f"\n\nNote: There are 2 withdrawals for taxes, a large one and a small one.\nThe small one is the total FUTA (${round(total_futa,2)}) taxes (Federal unemployment tax).\nThe larger one consists of: (employee_taxes + employer_taxes - FUTA + CA_ETT)\nHowever, since we take out the FUTA charge per employee, we don't do it here.")

print(f"\nTAX: Every pay category with non-zero pay:\n")
print(f"\tAdmin: \t\t{stated_ca_ett} (CA ETT)")

calculated_employer_taxes_minus_futa = 0
calculated_employee_taxes = 0
tax_per_program_sum = 0
for program, tax in tax_by_program.items():
    combined_taxes = tax + employer_taxes_by_program[program]
    if combined_taxes > 0:
        calculated_employee_taxes += tax
        calculated_employer_taxes_minus_futa += employer_taxes_by_program[program]
        spaces = "\t" if len(program) >= 6 else "\t\t"
        print(f"\t{program}: {spaces}{round(tax, 2)}   \t+ {round(employer_taxes_by_program[program], 2)}   \t= {round(combined_taxes, 2)}")

        tax_per_program_sum += combined_taxes


print(f"\n\n\tCalculated Total Taxes: \t\t{round(tax_per_program_sum + stated_ca_ett, 2)}")
print(f"\tIt should equal (stated values): \t{stated_taxes + stated_employer_taxes_minus_futa + stated_ca_ett}\n")

print(f"\tCalculated Employee Taxes: \t\t{round(calculated_employee_taxes, 2)}\t(All employee taxes added together)")
print(f"\tStated Employee Taxes: \t\t\t{round(stated_taxes, 2)}\n")

print(f"\tCalculated Employer Taxes: \t\t{round(calculated_employer_taxes_minus_futa, 2)}\t(All employer taxes per employee minus FUTA added together)")
print(f"\tStated Employer Taxes: \t\t\t{round(stated_employer_taxes_minus_futa, 2)}\n")

print(f"\tCalculated Employer FUTA: \t\t{round(total_futa, 2)} \t(All employee FUTA values added together)")
print(f"\tStated Employer FUTA: \t\t\t{round(stated_employer_futa, 2)}\n")


print(f"\n\temployee_taxes + employer_taxes + CA_ETT = Amount Taken From Bank Account")
print(f"\t{calculated_employee_taxes} + {calculated_employer_taxes_minus_futa} + {stated_ca_ett} = {calculated_employee_taxes + calculated_employer_taxes_minus_futa + stated_ca_ett}")

# print(f"\nEMPLOYER TAX:\n")
# # this one we can calculate on the fly using already gathered data and percentages
# for program, tax in tax_by_program.items():
#     proportial_taxes = (tax / calculated_employee_taxes) * employer_taxes
#     print(f"\t{program}: {round(proportial_taxes, 2)}")
#     calculated_employer_taxes_minus_futa += proportial_taxes
#     employer_taxes_by_program[program] = proportial_taxes


total_direct_deposited = 0
print(f"\nNET: Every pay category with non-zero pay:\n")
for program, net in net_by_program.items():
    if net > 0:
        print(f"\t{program}: {round(net, 2)}")
        total_direct_deposited += net



print(f"\n\tCALCULATED NET: \t\t{round(total_direct_deposited, 2)}")
print(f"\tSTATED NET: \t\t\t{round(stated_net, 2)}")

print(f"\n\tCALCULATED DEDUCTIONS: \t\t{calculated_deductions}")
print(f"\tSTATED DEDUCTIONS: \t\t{round(stated_deductions, 2)}\n")

print(f"\tCOMBINED CALCULATED: \t\t{round(calculated_deductions + calculated_employee_taxes + total_direct_deposited, 2)}\n")

print(f"\tCALCULATED GROSS: \t\t{round(calculated_gross_total, 2)}")
print(f"\tSTATED GROSS: \t\t\t{round(stated_gross, 2)}\n")







# # calculated_hourly_gross = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# # calculated_tax_rate = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# # calculated_hourly_net = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# # calculated_hourly_taxed = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]

# per_employee_hourly_gross = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# per_employee_tax_rate = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# per_employee_hourly_net = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]
# per_employee_hourly_taxed = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1]

# data_dictionary = {
#   'Programs': ['Events', 'Gymnastics', 'Hospitality', 'Tumblebunny', 'Dance', 'Swim', 'TAG', 'Polkadots', 'Sick', 'Clean'],

#   # Looking at the provided "Department" totals, so not breaking things up per hour, just based on an employee's default department
#   # 'Dept Gross Pay': dept_gross,
#   # 'Dept Net Pay': dept_net,
#   # 'Dept Taxes Pay': dept_taxed,

#   # Using the hours submitted for each program on Page 16. Should be more accurate, especially for wages, but it normalizes taxes
#   'Calculated Hourly Gross Pay': calculated_hourly_gross,
#   'Calculated Effective Tax Rate': calculated_tax_rate,
#   'Calculated Hourly Net Pay': calculated_hourly_net,
#   'Calculated Hourly Taxed Pay': calculated_hourly_taxed,

#   # Calculate a tax rate per employee using their own hours and taxed amount
#   'Per Employee Hourly Gross Pay': per_employee_hourly_gross,
#   'Per Employee Effective Tax Rate': per_employee_tax_rate,
#   'Per Employee Hourly Net Pay': per_employee_hourly_net,
#   'Per Employee Hourly Taxed Amount': per_employee_hourly_taxed
# }

# result_df = pd.DataFrame(data_dictionary)



# empty_row = pd.DataFrame({" ":[" "]})

# combined_df = pd.DataFrame()
# if grand_total_data_frame is not None:
#     combined_df = pd.concat([combined_df, grand_total_data_frame], ignore_index=True)
#     combined_df = pd.concat([combined_df, empty_row], ignore_index=True)


# if departments_data_frame is not None:
#     combined_df = pd.concat([combined_df, departments_data_frame], ignore_index=True)
#     combined_df = pd.concat([combined_df, empty_row], ignore_index=True)

# if result_df is not None:
#     combined_df = pd.concat([combined_df, result_df], ignore_index=True)
#     combined_df = pd.concat([combined_df, empty_row], ignore_index=True)

# # Generate the output file name based on the input file name
# output_file = args.input_file.replace('.xls', '_output.csv')

# # Save the result DataFrame to CSV
# if combined_df is not None:
#     combined_df.to_csv(output_file, index=False)
# else:
#     result_df.to_csv(output_file, index=False)

# print(f"Results saved to {output_file}")
