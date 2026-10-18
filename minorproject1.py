from openpyxl import load_workbook

workbook = load_workbook("12610019.xlsx", data_only=True)
sheet = workbook["table"]




#   TECH PRODUCTIVITY (TPI)

coding_col = 0

for cell in sheet[1]:
    if cell.value == "Coding (min)":
        coding_col = cell.column
        break

coding_minutes = []

for row in range(3, sheet.max_row + 1):
    value = sheet.cell(row, coding_col).value
    if value is not None:
        coding_minutes.append(value)
    

tpi = sum(coding_minutes) / len(coding_minutes)

print("TPI:",tpi, "min/day")





# ACADEMIC ACTIVITY (AAI)

study_col = 0
class_col = 0

for cell in sheet[1]:
    if cell.value == "Study (min)":
        study_col = cell.column
    if cell.value == "Class (min)":
        class_col = cell.column

daily_totals_aai = []

for row in range(2, sheet.max_row + 1):
    study = sheet.cell(row, study_col).value
    class_time = sheet.cell(row, class_col).value
    if study is not None and class_time is not None:
        daily_totals_aai.append(study + class_time)


aai = sum(daily_totals_aai) / len(daily_totals_aai)
print("AAI:", aai, "min/day")





#  PHYSICAL ACTIVITY (PhAI)

fitness_colm = 0

for cell in sheet[1]:
    if cell.value == "Fitness (min)":
        fitness_colm = cell.column

daily_totals_phai = []

for row in range(2, sheet.max_row + 1):
    fitness = sheet.cell(row, fitness_colm).value
    if fitness is not None:
        daily_totals_phai.append(fitness)


PhAI = sum(daily_totals_phai) / len(daily_totals_phai)
print("PhAI:", PhAI, "min/day") 





#    Sleep and Recovery (sri)

sleep_colm = 0

for cell in sheet[1]:
    if cell.value == "Sleep (min)":
        sleep_colm = cell.column

daily_totals_sri = []

for row in range (2, sheet.max_row + 1):
    sleep = sheet.cell(row, sleep_colm).value
    if sleep is not None:
        daily_totals_sri.append(sleep)

SRI = sum(daily_totals_sri) / len(daily_totals_sri)
print("SRI: ", SRI, "min/day")




#    activity balance (ABI)

unaccounted_colm = 0

for cell in sheet[1]:
    if cell.value == "Free/Unaccounted (min)":
        unaccounted_colm = cell.column

daily_totals_abi = []

for row in range (2, sheet.max_row +1):
    unaccounted = sheet.cell(row, unaccounted_colm).value
    if unaccounted is not None:
        daily_totals_abi.append(unaccounted)

ABI = sum(daily_totals_abi)/ len(daily_totals_abi)
print("ABI: ", ABI, "min/day")




#  time utilization (TUI)

tracked_colm = 0

for cell in sheet[1]:
    if cell.value == ("Total Tracked (min)"):
        tracked_colm = cell.column

daily_totals_tui = []

for row in range (2, sheet.max_row +1):
    tracked = sheet.cell(row, tracked_colm).value
    if tracked is not None:
        daily_totals_tui.append(tracked)

TUI = sum(daily_totals_tui) / len(daily_totals_tui)
print("TUI: ", TUI, "min/day")





#   experiance index (EI)

feeling_col = 0
satisfaction_col = 0
energy_col = 0

for cell in sheet[1]:
    if cell.value == "Day's Feeling":
        feeling_col = cell.column
    if cell.value == "Satisfaction Level":
        satisfaction_col = cell.column
    if cell.value == "Energy Level":
        energy_col = cell.column

feeling_scores = {
    "Excellent": 5,
    "Good": 4,
    "Neutral": 3,
    "Low": 2,
    "Stressed": 1
}

satisfaction_scores = {
    "Very Satisfied": 5,
    "Satisfied": 4,
    "Neutral": 3,
    "Unsatisfied": 2,
    "Very Unsatisfied": 1
}

energy_scores = {
    "High": 5,
    "Medium": 3,
    "Low": 1
}

total_score = 0



for row in range(2, sheet.max_row + 1):
    feeling = sheet.cell(row, feeling_col).value
    satisfaction = sheet.cell(row, satisfaction_col).value
    energy = sheet.cell(row, energy_col).value

    if feeling in feeling_scores and satisfaction in satisfaction_scores and energy in energy_scores:
        total_score += feeling_scores[feeling]
        total_score += satisfaction_scores[satisfaction]
        total_score += energy_scores[energy]
        
valid_days = sheet.max_row+1

ei = total_score / (3 * valid_days)
print("EI:", ei, "out of 5")




#    data contnuity (DCI)


from datetime import datetime, date

valid_dates = []


for row in range(2, sheet.max_row + 1):
    recorded_date = sheet.cell(row, 1).value

    if isinstance(recorded_date, (datetime, date)):
        if recorded_date not in valid_dates:
            valid_dates.append(recorded_date)

expected_days = 44
dci = len(valid_dates) / expected_days * 100

print("Valid recorded days:", len(valid_dates))
print("DCI:", round(dci, 2), "%")





#   Personal Activity Index (PAI)


pai = (
    0.15 * tpi
    + 0.20 * aai
    + 0.15 * PhAI
    + 0.20 * SRI
    + 0.15 * TUI
    + 0.10 * ei
    + 0.05 * dci
)

print("PAI:", pai)