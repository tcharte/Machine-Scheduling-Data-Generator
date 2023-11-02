# Imports
import os
import itertools
import numpy as np
import pandas as pd
# Also requires xlsxwriter to be installed

np.random.seed(0)

# Parameters
job_numbers = [30, 60, 90, 120, 150, 180]
machine_numbers = [2, 4, 6, 8]
# scheduling_weeks = [3, 4, 6] # up to
scheduling_weeks = [1, 2, 3, 4, 6] # up to
personnel_capacity = [1, 2, 3, 4, 5, 6, 7] # up to
TW_density = [0, 1] # Percentage of jobs with time windows
# TW_density = [0.5] # Percentage of jobs with time windows
machine_eligibility_constraint = [0, 1] # percent of jobs that can only be processed on a given machine
weekly_personnel_availability = 2250  # Weekly personnel availability value

# Functions
def generateJobProcessingTimes(min_processing_time, max_processing_time, n_machines, n_jobs, machine_eligibility):
    # job_proc_times = np.random.uniform(low=min_processing_time, high=max_processing_time, size=(n_machines, n_jobs)) # Continuous Uniform distribution
    job_proc_times = np.random.randint(low=min_processing_time, high=max_processing_time, size=(n_machines, n_jobs)) # Discrete Uniform distribution

    initial_setup_time = max_processing_time

    job_initial_setup_times = np.full((n_machines, n_jobs), initial_setup_time)

    machine_eligibilities = generateMachineEligibilities(n_machines, n_jobs, machine_eligibility)

    # set processing times to zero on non-eligible machines
    for i, machine_eligibility in enumerate(machine_eligibilities):
        for job in range(n_jobs):
            if job not in machine_eligibility:
                job_proc_times[i, job] = 0
                # job_initial_setup_times[i, job] = 0
    
    return job_proc_times, job_initial_setup_times, machine_eligibilities

def generateSequenceDependantSetupTimes(n_machines, n_jobs, setup_time_lb, setup_time_ub):
    # Sequence Dependant Setup Times

    # Based on coordinate method in paper to ensure triangle inequality is met
    # Equation from paper:
    # l + ((u-l)/100)[|x1a -x2a|+|y1a-y2a|]      (l+((u-l)/100)[|x1b-x2b|+|y1b-y2b|])

    plane_coordinates = np.full([n_machines, n_jobs, 2, 2], 0)
    seq_dep_setup_times = np.full([n_machines, n_jobs, n_jobs], 0)
    for i in range(n_machines):
        np.fill_diagonal(seq_dep_setup_times[i], 0, wrap=False)
        for j in range(n_jobs):
            # plane_coordinates[i,j,:] = np.random.uniform(low=0.0, high=50.0, size=[2,2]) # 2x2 matrix of coordinates based on paper
            plane_coordinates[i,j,:] = np.random.randint(low=0.0, high=50.0, size=[2,2]) # 2x2 matrix of coordinates based on paper

        for j in range(n_jobs):
            for k in range(n_jobs):
                if j!=k:
                    if j>k:
                        x1a = plane_coordinates[i, j, 0, 0]
                        x2a = plane_coordinates[i, k, 0, 0]
                        y1a = plane_coordinates[i, j, 0, 1]
                        y2a = plane_coordinates[i, k, 0, 1]
                    elif k>j:
                        x1a = plane_coordinates[i, j, 1, 0] # x1b
                        x2a = plane_coordinates[i, k, 1, 0] # x2b
                        y1a = plane_coordinates[i, j, 1, 1] # y1b
                        y2a = plane_coordinates[i, k, 1, 1] # y2b

                    seq_dep_setup_times[i, j, k] = setup_time_lb + ((setup_time_ub-setup_time_lb)/100)*(np.abs(x1a -x2a)+np.abs(y1a-y2a))
                    
    return seq_dep_setup_times

def generateTimeWindows(n_jobs, n_weeks, tw_density, tw_length, week_minutes):
    day_minutes = week_minutes/5

    # Generate random indexes for windows
    n_tw = int(tw_density * n_jobs) # number of jobs with time windows
    tw_indexes = np.random.choice(n_jobs, size=n_tw, replace=False) # indexes of jobs with time windows

    # Generate random lengths for time windows
    # tw_lengths = np.random.randint(low=1, high=tw_length, size=n_tw) # lengths of time windows (weeks)

    # Generate random time window lengths based on number of days rather than weeks
    tw_lengths = np.random.randint(low=0, high=5*tw_length, size=n_tw) # lengths of time windows (days)

    # Generate random start times for time windows
    # tw_starts = np.random.randint(low=0, high=n_weeks-tw_lengths, size=n_tw) # start times of time windows (weeks)

    # Generate start time based on days rather than weeks
    tw_starts = np.random.randint(low=0, high=5*n_weeks-tw_lengths, size=n_tw) # start times of time windows (days)
    tw_ends = tw_starts + tw_lengths

    # Create empty arrays for time windows
    release_times = np.full((n_jobs, n_weeks), 0)
    delivery_times = np.full((n_jobs, n_weeks), 0)

    release_periods = np.full(n_jobs, 1) # initialize all to week 1
    delivery_periods = np.full(n_jobs, n_weeks) # initialize all to last week

    # Convert tw_starts to release times
    for i, val in enumerate(tw_indexes):
        release_day = tw_starts[i] % 5 # day of week
        release_week = tw_starts[i] // 5
        # for j in range(release_week):
        #     release_times[val, j] = day_minutes*5
        release_times[val, release_week] = release_day * day_minutes # start of day, no need to account for index
        release_periods[val] = release_week+1 # plus one to account for index

    # Convert tw_lengths to delivery times
    for i, val in enumerate(tw_indexes):
        delivery_day = (tw_ends[i]) % 5 # day of week
        delivery_week = (tw_ends[i]) // 5
        for j in range(delivery_week):
            delivery_times[val, j] = day_minutes*5
        delivery_times[val, delivery_week] = (delivery_day+1) * day_minutes # plus one for end of day
        delivery_periods[val] = delivery_week+1 # plus one to account for index
    
    # Set all other jobs to have a delivery time of 5 days
    for i in range(n_jobs):
        if i not in tw_indexes:
            for j in range(n_weeks):
                delivery_times[i, j] = day_minutes*5

    # return tw_starts, tw_lengths, release_times, delivery_times, tw_indexes

    return release_times, delivery_times, release_periods, delivery_periods

def generateMachineEligibilities(n_machines, n_jobs, eligibility_constraint):
    # Create a list for each machine to store the jobs that can be processed on it
    machine_eligibilities = [[] for _ in range(n_machines)]
    
    # randomly select indexes for number of jobs that can only be processed on a given machine
    n_with_eligibility = int(n_jobs * eligibility_constraint)
    indexes = np.random.choice(n_jobs, n_with_eligibility, replace=False)

    # Assign the jobs without contraint to all machines and the jobs with contraint to a single machine
    for i in range(n_jobs):
        if i in indexes:
            machine_eligibilities[np.random.randint(0, n_machines)].append(i)
        else:
            for j in range(n_machines):
                machine_eligibilities[j].append(i)

    return machine_eligibilities

def generateInitialSetupTimes(n_machines, n_jobs, upper_bound):
    # create an array of initial setup times for each job
    initial_setup_times = np.full((n_machines, n_jobs), upper_bound)

    return initial_setup_times

def generatePersonnelTimes(n_staff, n_weeks, week_minutes):
    people_times = np.full((n_staff, n_weeks), week_minutes)

    # Generate Personnel Assignments
    personnel_assignments = [[] for i in range(n_staff)]
    for i in range(n_staff):
        personnel_assignments[i].append(2*i+1)
        personnel_assignments[i].append(2*i+2)


    return people_times, personnel_assignments

def writeExcelOld(index, n_machines, n_weeks, job_proc_times, seq_dep_setup_times, release_times, delivery_times, release_periods, delivery_periods, people_times, machine_eligibilities, initial_setup_times, tw_density, weekly_personnel_availability, mean_value, upper_bound, lower_bound, machine_eligibility, personnel_assignments):

    # Job processing times
    df_job_proc_times = pd.DataFrame(job_proc_times.T)
    df_job_proc_times.columns = [f'Machine {i}' for i in range(n_machines)]
    df_job_proc_times.index.name = "Job"

    # Sequence Dependant Setup Times
    dfs_seq_dep_setup_times = {}
    for i in range(n_machines):
        dfs_seq_dep_setup_times[i] = pd.DataFrame(seq_dep_setup_times[i,:,:])

    # Time Windows
    df_periods = pd.DataFrame()
    df_periods['Release Periods'] = release_periods
    df_periods['Delivery Periods'] = delivery_periods
    df_periods.index.name = 'Job'

    df_release_times = pd.DataFrame(release_times)
    df_release_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_release_times.index.name = 'Job'

    df_delivery_times = pd.DataFrame(delivery_times)
    df_delivery_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_delivery_times.index.name = 'Job'

    # Personnel Times
    df_people_times = pd.DataFrame(people_times)
    df_people_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_people_times.index.name = 'Personnel'

    # Initial Setup Times
    df_initial_setup_times = pd.DataFrame(initial_setup_times.T)
    df_initial_setup_times.columns = [f'Machine {i}' for i in range(n_machines)]
    df_initial_setup_times.index.name = "Job"


    # Check for data dir
    if not os.path.exists('data'):
        os.makedirs('data')

    # Write to excel
    writer = pd.ExcelWriter(f'./data/Scheduling_Instance_{index}.xlsx', engine='xlsxwriter')
    # workbook = writer.book

    # Write instance info sheet
    worksheet = writer.book.add_worksheet('Instance Info')
    worksheet.write_string(0, 0, 'Instance Info')
    worksheet.write_string(1, 0, 'Instance Number')
    worksheet.write_string(1, 1, str(index))
    worksheet.write_string(2, 0, 'Number of Jobs')
    worksheet.write_string(2, 1, str(n_jobs))
    worksheet.write_string(3, 0, 'Number of Machines')
    worksheet.write_string(3, 1, str(n_machines))
    worksheet.write_string(4, 0, 'Number of Weeks')
    worksheet.write_string(4, 1, str(n_weeks))
    worksheet.write_string(5, 0, 'Personnel Available')
    worksheet.write_string(5, 1, str(n_staff))
    worksheet.write_string(6, 0, 'Max Time Window Length (Days)')
    worksheet.write_string(6, 1, str(n_weeks*5))
    worksheet.write_string(7, 0, 'Time Window Density')
    worksheet.write_string(7, 1, str(tw_density))
    worksheet.write_string(8, 0, 'Weekly Personnel Availability')
    worksheet.write_string(8, 1, str(weekly_personnel_availability))
    worksheet.write_string(9, 0, 'Mean Time Value')
    worksheet.write_string(9, 1, str(mean_value))
    worksheet.write_string(10, 0, 'Upper Bound')
    worksheet.write_string(10, 1, str(upper_bound))
    worksheet.write_string(11, 0, 'Lower Bound')
    worksheet.write_string(11, 1, str(lower_bound))
    worksheet.write_string(12, 0, 'Machine Eligibility Constraint')
    worksheet.write_string(12, 1, str(machine_eligibility))

    df_job_proc_times.to_excel(
        writer, 'Job Info',
        startcol=0,
        startrow=1
    )
    worksheet = writer.sheets['Job Info']
    worksheet.write_string(0, 0, 'Job Processing Times')

    # Initial Setup Times
    df_initial_setup_times.to_excel(
        writer, 'Job Info',
        startcol=n_machines+2,
        startrow=1
    )
    worksheet.write_string(0, n_machines+2, 'Initial Setup Times')

    # Machine Eligibilities
    worksheet.write_string(0, 2*n_machines+4, 'Machine Eligibilities')
    for i in range(n_machines):
        worksheet.write_string(i+1, 2*n_machines+4, f'Machine {i}')
        worksheet.write_string(i+1, 2*n_machines+5, str(machine_eligibilities[i]))

    # Personnel Times
    df_people_times.to_excel(
        writer, 'Job Info',
        startcol=2*n_machines+4,
        startrow=n_machines+3
    )
    worksheet.write_string(n_machines+2, 2*n_machines+4, 'Personnel Times')

    # Personnel Assignments
    for i in range(n_staff):
        worksheet.write_string(i+n_machines + n_staff + 6, 2*n_machines+4, f'Personnel {i}')
        worksheet.write_string(i+n_machines + n_staff + 6, 2*n_machines+5, str(personnel_assignments[i]))
    worksheet.write_string(n_machines + n_staff + 5, 2*n_machines+4, 'Personnel Assignments')

    # Time Windows
    df_periods.to_excel(
        writer, 'Time Windows',
        startcol=0,  # n_machines+2,
        startrow=1
    )
    worksheet = writer.sheets['Time Windows']
    worksheet.write_string(0, 0, 'Time Periods')

    #worksheet.write_string(0, n_machines+2, 'Periods')

    df_release_times.to_excel(
        writer, 'Time Windows',
        startcol=4,
        startrow=1
    )
    worksheet.write_string(0, 4, 'Release Times')

    df_delivery_times.to_excel(
        writer, 'Time Windows',
        startcol=n_weeks+6,
        startrow=1
    )
    worksheet.write_string(0, n_weeks+6, 'Delivery Times')

    for i, df in dfs_seq_dep_setup_times.items():
        df.to_excel(
            writer, f'Setup Times Machine {i}',
            startcol=0,
            startrow=0
        )

    # writer.save()
    writer.close()

def writeExcel(index, n_machines, n_weeks, job_proc_times, seq_dep_setup_times, release_times, delivery_times, release_periods, delivery_periods, people_times, machine_eligibilities, initial_setup_times, tw_density, weekly_personnel_availability, mean_value, upper_bound, lower_bound, machine_eligibility, personnel_assignments):

    # Job processing times
    df_job_proc_times = pd.DataFrame(job_proc_times.T)
    df_job_proc_times.columns = [f'Machine {i}' for i in range(n_machines)]
    df_job_proc_times.index.name = "Job"

    # Sequence Dependant Setup Times
    dfs_seq_dep_setup_times = {}
    for i in range(n_machines):
        dfs_seq_dep_setup_times[i] = pd.DataFrame(seq_dep_setup_times[i,:,:])

    # Time Windows
    df_periods = pd.DataFrame()
    df_periods['Release Periods'] = release_periods
    df_periods['Delivery Periods'] = delivery_periods
    df_periods.index.name = 'Job'

    df_release_times = pd.DataFrame(release_times)
    df_release_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_release_times.index.name = 'Job'

    df_delivery_times = pd.DataFrame(delivery_times)
    df_delivery_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_delivery_times.index.name = 'Job'

    # Personnel Times
    df_people_times = pd.DataFrame(people_times)
    df_people_times.columns = [f'Week {i+1}' for i in range(n_weeks)]
    df_people_times.index.name = 'Personnel'

    # Initial Setup Times
    df_initial_setup_times = pd.DataFrame(initial_setup_times.T)
    df_initial_setup_times.columns = [f'Machine {i}' for i in range(n_machines)]
    df_initial_setup_times.index.name = "Job"


    # Check for data dir
    if not os.path.exists('data'):
        os.makedirs('data')

    # Write to excel
    writer = pd.ExcelWriter(f'./data/Scheduling_Instance_{index}.xlsx', engine='xlsxwriter')
    # workbook = writer.book

    # Write instance info sheet
    worksheet = writer.book.add_worksheet('Instance Info')
    worksheet.write(0, 0, 'Instance Info')
    worksheet.write(1, 0, 'Instance Number')
    worksheet.write(1, 1, index)
    worksheet.write(2, 0, 'Number of Jobs')
    worksheet.write(2, 1, n_jobs)
    worksheet.write(3, 0, 'Number of Machines')
    worksheet.write(3, 1, n_machines)
    worksheet.write(4, 0, 'Number of Weeks')
    worksheet.write(4, 1, n_weeks)
    worksheet.write(5, 0, 'Personnel Available')
    worksheet.write(5, 1, n_staff)
    worksheet.write(6, 0, 'Max Time Window Length (Days)')
    worksheet.write(6, 1, n_weeks*5)
    worksheet.write(7, 0, 'Time Window Density')
    worksheet.write(7, 1, tw_density)
    worksheet.write(8, 0, 'Weekly Personnel Availability')
    worksheet.write(8, 1, weekly_personnel_availability)
    worksheet.write(9, 0, 'Mean Time Value')
    worksheet.write(9, 1, mean_value)
    worksheet.write(10, 0, 'Upper Bound')
    worksheet.write(10, 1, upper_bound)
    worksheet.write(11, 0, 'Lower Bound')
    worksheet.write(11, 1, lower_bound)
    worksheet.write(12, 0, 'Machine Eligibility Constraint')
    worksheet.write(12, 1, machine_eligibility)

    # Write Sets Sheet
    worksheet = writer.book.add_worksheet('Sets')
    worksheet.write(0, 0, 'NumberOfOrders')
    worksheet.write(1, 0, n_jobs+1)

    worksheet.write(3, 0, 'EndOrder')
    worksheet.write(4, 0, n_jobs+1)

    worksheet.write(6, 0, 'NumberOfProductionLines')
    worksheet.write(7, 0, n_machines)

    worksheet.write(9, 0, 'NumberOfCrewsWithReassignment')
    worksheet.write(10, 0, n_staff*2)

    worksheet.write(12, 0, 'NumberOfCrewsWithoutReassignment')
    worksheet.write(13, 0, n_staff)

    worksheet.write(15, 0, 'NumberOfWeeks')
    worksheet.write(16, 0, n_weeks)

    # Machine Eligibilities
    worksheet.write(0, 2, 'Machine Eligibilities')
    for i in range(n_machines):
        worksheet.write(i+1, 2, f'Machine {i+1}')
        # worksheet.write(i+1, 3, machine_eligibilities[i])
    # worksheet.write(1, 3, f'data') # TODO: fix this to write the actual data
    formatted_machine_eligibilities = [f'{{{", ".join(map(str, machine_eligibilities[i]))}}}' for i in range(len(machine_eligibilities))]
    formatted_machine_eligibilities = str(dict(zip([n+1 for n in range(len(formatted_machine_eligibilities))],formatted_machine_eligibilities)))
    formatted_machine_eligibilities = formatted_machine_eligibilities.replace("'",'')
    formatted_machine_eligibilities = f"data {formatted_machine_eligibilities};"
    worksheet.write(1, 3, formatted_machine_eligibilities)

    # lists_temp = [f'{{{", ".join(map(str, lists[i]))}}}' for i in range(len(lists))]
    # lists_temp = str(dict(zip([n+1 for n in range(len(lists_temp))],lists_temp)))
    # lists_temp = lists_temp.replace("'",'')
    # print(f'data {lists_temp};')
    
    # Personnel Assignments
    worksheet.write(0, 5, 'Personnel Assignments')
    for i in range(n_staff):
        worksheet.write(i+1, 5, f'Personnel {i+1}')
        # worksheet.write(i+1, 6, str(personnel_assignments[i]))
    # worksheet.write(n_staff + 1, 5, 'data') # TODO: fix this to write the actual data
    formatted_personnel_assignments = [f'{{{", ".join(map(str, personnel_assignments[i]))}}}' for i in range(len(personnel_assignments))]
    formatted_personnel_assignments = str(dict(zip([n+1 for n in range(len(formatted_personnel_assignments))],formatted_personnel_assignments)))
    formatted_personnel_assignments = formatted_personnel_assignments.replace("'",'')
    formatted_personnel_assignments = f"data {formatted_personnel_assignments};"
    worksheet.write(1, 6, formatted_personnel_assignments)

    # Write OrderInfo Sheet
    worksheet = writer.book.add_worksheet('OrderInfo')
    # Order Processing Times
    worksheet.write(0, 0, 'OrdersProcessingTimes')
    worksheet.write(1, 0, 'Orders')
    worksheet.write(1, 1, 'Machines')
    for i in range(n_machines):
        worksheet.write(2, i+1, i+1)
    for i in range(n_jobs):
        worksheet.write(i+3, 0, i+1)
        for j in range(n_machines):
            worksheet.write(i+3, j+1, job_proc_times[j, i])

    # Order Initial Setup Times
    worksheet.write(0, 10, 'Time_Begin')
    worksheet.write(1, 10, 'Orders')
    worksheet.write(1, 11, 'Machines')
    for i in range(n_machines):
        worksheet.write(2, i+11, i+1)
    for i in range(n_jobs):
        worksheet.write(i+3, 10, i+1)
        for j in range(n_machines):
            worksheet.write(i+3, j+11, initial_setup_times[j, i])

    # Release Weeks
    worksheet.write(0, 20, 'ReleaseWeeks')
    worksheet.write(1, 20, 'Orders')
    worksheet.write(1, 21, 'ReleaseWeek')
    for i in range(n_jobs):
        worksheet.write(i+2, 20, i+1)
        worksheet.write(i+2, 21, release_periods[i])

    # Delivery Weeks
    worksheet.write(0, 23, 'DeliveryWeeks')
    worksheet.write(1, 23, 'Orders')
    worksheet.write(1, 24, 'Deadline')
    for i in range(n_jobs):
        worksheet.write(i+2, 23, i+1)
        worksheet.write(i+2, 24, delivery_periods[i])

    # Release Times
    worksheet.write(0, 26, 'ReleaseTimes')
    worksheet.write(1, 26, 'Orders')
    worksheet.write(1, 27, 'Weeks')
    for i in range(n_weeks):
        worksheet.write(2, i+27, i+1)
    for i in range(n_jobs):
        worksheet.write(i+3, 26, i+1)
        for j in range(n_weeks):
            worksheet.write(i+3, j+27, release_times[i, j])

    # Delivery Times
    worksheet.write(0, 34, 'DeliveryTimes')
    worksheet.write(1, 34, 'Orders')
    worksheet.write(1, 35, 'Weeks')
    for i in range(n_weeks):
        worksheet.write(2, i+35, i+1)
    for i in range(n_jobs):
        worksheet.write(i+3, 34, i+1)
        for j in range(n_weeks):
            worksheet.write(i+3, j+35, delivery_times[i, j])

    # Write CleanTimes Sheet
    worksheet = writer.book.add_worksheet('CleanTimes')
    worksheet.write(0, 0, 'Machine')
    worksheet.write(0, 1, 'Order')
    for i in range(n_jobs):
        worksheet.write(0, i+2, i+1)
    for i in range(n_machines):
        for j in range(n_jobs):
            worksheet.write(i*n_jobs+j+1, 0, i+1)
            worksheet.write(i*n_jobs+j+1, 1, j+1)
            for j2 in range(n_jobs):
                # if j != j2:
                worksheet.write(i*n_jobs+j+1, j2+2, seq_dep_setup_times[i, j, j2])
    
    writer.close()

# Main

# Check Valid Combinations

# Initialize a list to store valid combinations
valid_combinations = []

# Iterate through all combinations
for instance_number, (
    num_jobs, num_machines, num_weeks, personnel_available, tw_percent, machine_eligibility
) in enumerate(
        itertools.product(
            job_numbers, machine_numbers, scheduling_weeks, personnel_capacity, TW_density, machine_eligibility_constraint
        ), start=1):
    # Rule 1: The ratio of (number of jobs) / (number of machines * number of weeks) is not less than 15.
    if num_jobs >= 15 * num_machines * num_weeks:
        # Rule 2: The number of personnel available is less than the number of machines.
        if personnel_available < num_machines:
            # Rule 3: The ratio of machines to personnel is less than or equal to 3.
            if num_machines / personnel_available <= 3:
                # Calculate the total processing time for this instance
                jobs_per_machine = num_jobs / (num_machines)
                total_completion_time = personnel_available * num_weeks * weekly_personnel_availability
                mean_value = total_completion_time / (num_machines*(jobs_per_machine*2 - 1))
                upper_bound = int((4/3)*mean_value)
                lower_bound = int((2/3)*mean_value)
                # Create a dictionary to store the parameters, instance number, and relevant values
                instance_data = {
                    'Instance Number': instance_number,
                    'Number of Jobs': num_jobs,
                    'Number of Machines': num_machines,
                    'Number of Weeks': num_weeks,
                    'Personnel Available': personnel_available,
                    'Time Window Length': num_weeks, #todo remove
                    'Time Window Density': tw_percent,
                    'Weekly Personnel Availability': weekly_personnel_availability,
                    'Mean Value' : mean_value,
                    'Uppwer Bound' : upper_bound,
                    'Lower Bound' : lower_bound,
                    'Machine Eligibility Constraint' : machine_eligibility
                }
                valid_combinations.append(instance_data)

# Create a DataFrame from the list of valid combinations
df = pd.DataFrame(valid_combinations)

for index, instance in df.iterrows():
    # if index not in [138, 45, 67, 101]:
    #     continue

    n_jobs = int(instance['Number of Jobs'])
    n_machines = int(instance['Number of Machines'])
    week_minutes = weekly_personnel_availability
    n_staff = int(instance['Personnel Available'])
    n_weeks = int(instance['Number of Weeks'])
    tw_density = instance['Time Window Density']
    tw_length = instance['Time Window Length']
    mean_value = instance['Mean Value']
    upper_bound = instance['Uppwer Bound']
    lower_bound = instance['Lower Bound']
    machine_eligibility = instance['Machine Eligibility Constraint']

    min_processing_time, max_processing_time, = lower_bound, upper_bound

    # Generate job processing times
    job_proc_times, job_initial_setup_times, machine_eligibilities = generateJobProcessingTimes(min_processing_time, max_processing_time, n_machines, n_jobs, machine_eligibility)

    # Generate machine eligibilities
    # machine_eligibilities = generateMachineEligibilities(n_machines, n_jobs, machine_eligibility)

    # Generate sequence dependant setup times
    seq_dep_setup_times = generateSequenceDependantSetupTimes(n_machines, n_jobs, lower_bound, upper_bound)

    # Generate time windows
    release_times, delivery_times, release_periods, delivery_periods = generateTimeWindows(n_jobs, n_weeks, tw_density, tw_length, week_minutes)

    # Generate personnel times
    people_times, personnel_assignments = generatePersonnelTimes(n_staff, n_weeks, week_minutes)

    # Generate initial setup times
    initial_setup_times = generateInitialSetupTimes(n_machines, n_jobs, upper_bound)

    # Write to excel
    writeExcel(index, n_machines, n_weeks, job_proc_times, seq_dep_setup_times, release_times, delivery_times, release_periods, delivery_periods, people_times, machine_eligibilities, initial_setup_times, tw_density, weekly_personnel_availability, mean_value, upper_bound, lower_bound, machine_eligibility, personnel_assignments)

    print(f'Instance {index} created')
print('Done')