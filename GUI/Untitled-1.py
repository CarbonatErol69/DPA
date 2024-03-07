# %%
import pandas as pd

# %%
df_wahl = pd.read_csv(r'C:\Users\DE88342\OneDrive - Grunenthal Group\Desktop\VSC - Py\DPArenamed\Data\IMPORT BOT2_Wahl.csv', delimiter= ';')
df_event = pd.read_csv(r'C:\Users\DE88342\OneDrive - Grunenthal Group\Desktop\VSC - Py\DPArenamed\Data\IMPORT BOT1_Veranstaltungsliste.csv', delimiter= ';')
df_raum = pd.read_csv(r'C:\Users\DE88342\OneDrive - Grunenthal Group\Desktop\VSC - Py\DPArenamed\Data\IMPORT BOT0_Raumliste.csv', delimiter= ';')

# %%
df_wahl.head(20)

# %%
value_counts = pd.concat([df_wahl['Wahl 1'].value_counts(),
                         df_wahl['Wahl 2'].value_counts(),
                         df_wahl['Wahl 3'].value_counts(),
                         df_wahl['Wahl 4'].value_counts(),
                         df_wahl['Wahl 5'].value_counts(),
                         df_wahl['Wahl 6'].value_counts()], axis=1)

column_names = ['Wahl1', 'Wahl2', 'Wahl3', 'Wahl4', 'Wahl5', 'Wahl6']
value_counts.columns = column_names

value_counts.insert(0, 'event_number', value_counts.index)

print(value_counts)


# %%
df_event = df_event.rename(columns={'Nr. ': 'event_number'})
df_event

# %%
value_counts

# %%
def count_company_votes(df_wahl, df_event):
    
    # Sum the votes for each company
    company_votes = value_counts.sum(axis=1)
    
    # Add the votes for companies with different indices
    company_votes['Finanzamt'] = value_counts['Wahl1'].get(18, 0) + value_counts['Wahl1'].get(19, 0)
    
    # Divide the total votes by 20 to get the needed slots
    needed_slots = company_votes.sum() / 20
    
    return needed_slots

# Call the function with the loaded DataFrames
needed_slots = count_company_votes(df_wahl, df_event)


# %%
import math
merged_df = pd.merge(df_event, value_counts, on='event_number')

merged_df['Total 1-3'] = merged_df[['Wahl1', 'Wahl2', 'Wahl3']].sum(axis=1).astype(int)

merged_df['Total 4-6'] = merged_df[['Wahl4', 'Wahl5', 'Wahl6']].sum(axis=1).astype(int)

merged_df['Total'] = merged_df['Total 1-3'] + merged_df['Total 4-6']

merged_df['Nötige Slots'] = merged_df['Total 1-3'] / 20

merged_df['Optionale Slots'] = merged_df['Total 4-6'] / 20

# Round the 'Nötige Slots' column up or down
merged_df['Rundet Nötige Slots'] = merged_df['Nötige Slots'].apply(lambda x: math.ceil(x) if x - math.floor(x) >= 0.25 else math.floor(x))

# Round the 'Optionale Slots' column up or down
merged_df['Rundet Optionale Slots'] = merged_df['Optionale Slots'].apply(lambda x: math.ceil(x) if x - math.floor(x) >= 0.25 else math.floor(x))

merged_df.info()

# %%
df_raum

df_slots = pd.DataFrame(columns=['A', 'B', 'C', 'D', 'E'])



df_timetable = pd.concat([df_raum, df_slots[['A', 'B', 'C', 'D', 'E']]], axis=1)

df_timetable


# %%
# Sort df_merged by 'Rundet nötige slots' in descending order
df_merged_sorted = merged_df.sort_values('Total 1-3', ascending=False)

# Specify the columns to be modified
columns_to_modify = ['A', 'B', 'C', 'D', 'E']

# Iterate over the event_numbers
for event_number in df_merged_sorted['event_number']:
    # Get the corresponding 'Rundet nötige slots' value
    rundet_slots = df_merged_sorted.loc[df_merged_sorted['event_number'] == event_number, 'Rundet Nötige Slots'].values[0]
    
    # Iterate over the rows in df_timetable
    for row in df_timetable.index:
        # Check if the event_number has been assigned 'rundet_slots' times in the row and does not exist in the column
        if df_timetable.loc[row].eq(event_number).sum() < rundet_slots and event_number not in df_timetable[columns_to_modify].values:
            # Assign the event_number to the datapoint
            df_timetable.loc[row, columns_to_modify[df_timetable.loc[row].eq(event_number).sum()]] = event_number
        # Break the loop if the event_number has been assigned 'rundet_slots' times
        if df_timetable.loc[row].eq(event_number).sum() >= rundet_slots:
            break

df_timetable

# %%
# Sort df_merged by 'Rundet nötige slots' in descending order
df_merged_sorted = merged_df.sort_values('Total 1-3', ascending=False)

# Specify the columns to be modified
columns_to_modify = ['A', 'B', 'C', 'D', 'E']

# Iterate over the event_numbers
for event_number in df_merged_sorted['event_number']:
    # Get the corresponding 'Rundet nötige slots' value
    rundet_slots = df_merged_sorted.loc[df_merged_sorted['event_number'] == event_number, 'Total 1-3'].values[0]
    
    # Iterate over the rows in df_timetable
    for row in df_timetable.index:
        # Iterate over the specified columns only
        for col in columns_to_modify:
            # Check if the event_number has been assigned five times in the row and does not exist in the column
            if df_timetable.loc[row].eq(event_number).sum() < 5 and event_number not in df_timetable[col].values:
                # Assign the event_number to the datapoint
                df_timetable.loc[row, col] = event_number
                # Break the inner loop if the event_number has been assigned 'rundet_slots' times
                if df_timetable.loc[row].eq(event_number).sum() >= rundet_slots:
                    break
        # Break the outer loop if the event_number has been assigned 'rundet_slots' times
        if df_timetable.loc[row].eq(event_number).sum() >= rundet_slots:
            break

df_timetable

# %%
# Sort df_merged by 'Rundet nötige slots' in descending order
df_merged_sorted = merged_df.sort_values('Total 1-3', ascending=False)

# Specify the columns to be modified
columns_to_modify = ['A', 'B', 'C', 'D', 'E']

# Iterate over the event_numbers
for event_number in df_merged_sorted['event_number']:
    # Get the corresponding 'Rundet nötige slots' value
    rundet_slots = df_merged_sorted.loc[df_merged_sorted['event_number'] == event_number, 'Total 1-3'].values[0]
    
    # Iterate over the rows in df_timetable
    for row in df_timetable.index:
        # Iterate over the specified columns only
        for col in columns_to_modify:
            # Check if the event_number has been assigned five times in the row and does not exist in the column
            if df_timetable.loc[row].eq(event_number).sum() < 5 and event_number not in df_timetable[col].values:
                # Assign the event_number to the datapoint
                df_timetable.loc[row, col] = event_number
                # Break the inner loop if the event_number has been assigned 'rundet_slots' times
                if df_timetable.loc[row].eq(event_number).sum() >= rundet_slots:
                    break
        # Break the outer loop if the event_number has been assigned 'rundet_slots' times
        if df_timetable.loc[row].eq(event_number).sum() >= rundet_slots:
            break

df_timetable

# %%



