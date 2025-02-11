from openpyxl import load_workbook
import grab_data_functions
from grab_data_functions import combine_firstname_with_school_code_by_team, get_team_count
import nametag_gen_func
from nametag_gen_func import generate_name_tags

wb_file_path = "act_groups3.xlsx" # This will be given by the GUI
ws_name = "Sheet1" # Make GUI option to change this
work_book = load_workbook(wb_file_path)
work_sheet = work_book[ws_name]
team_count = get_team_count(work_sheet=work_sheet)

full_list = combine_firstname_with_school_code_by_team(wb_file_path, ws_name, team_count, "D", "I")

team_one = (full_list[0])
team_two = (full_list[1])
team_three = (full_list[2])
print(team_one)
print(team_two)
print(team_three)
print(team_count)

train_options = nametag_gen_func.get_train_options()
name_options = nametag_gen_func.get_name_options()
generate_name_tags(team_count, train_options, name_options, team_one, team_two, team_three)

teams = []
for team_number in range(len(full_list) + 1):
    if team_number == 0:
        pass
    else:
        teams.append('team ' + str(team_number))



