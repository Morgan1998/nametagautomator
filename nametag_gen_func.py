import xlsxwriter

def get_train_options(font='OPTIVagRound-Bold', color='#20b79f', bold=True, size=18): # Replace this font with the actual font used at office
   train_options = {
      'width': 220,
      'height': 58,
      'x_offset': 10,
      'y_offset': 3,
      'font': {'name': font, 'color': color, 'bold': bold, 'size': size},
      'align': {'vertical': 'middle',
                'horizontal': 'left'
                },
      'fill': {'none': True},
      'border': {'none': True},
   }
   return train_options

def get_name_options(font='OPTIVagRound-Bold', color='#20b79f', bold=True, size=28):
   name_options = {
      'width': 313,
      'height': 224,
      'font': {'name': font, 'color': color, 'bold': bold, 'size': size},
      'align': {'vertical': 'middle',
                'horizontal': 'center'
                },
      'fill': {'none': True},
      'border': {'none': True},
      'x_offset': 0,
   }
   return name_options

def get_logo_options():
   logo_options = {
      'x_offset': 185,
      'y_offset': 0,
      'x_scale': 1,
      'y_scale': 1,
   }
   return logo_options

def set_column(sheet_var):
   sheet_var.set_column(0, 1, 43.93)

def set_row(sheet_var, row_count_var):
   for row in range(row_count_var):
      sheet_var.set_row(row, 168)

def get_front_cell_list(group_row_count_var): # Gets list of cell names for front side of name tag e.g. [A1, A2, A3, A4]
   cells = []
   for i in range(group_row_count_var + 1):
      if i == 0:
         pass
      else:
         cells.append("A" + str(i))
   return cells

def get_bundles(group_front_cells_var, group_list_var): # bundles front cell values with corresponding (name, train group) pairs
   bundles = []
   for cell_value, pairs, in zip(group_front_cells_var, group_list_var):
      bundles.append([cell_value, pairs])
   return bundles

def evenly_bundle(group_bundles_var): # Unpacks the (name, train group) pairs to make an even list of bundles. Can probably find a way to do this all in one code within the get_bundles function
   evenly_bundled = []
   for bundle in group_bundles_var:
      x = bundle[1]
      evenly_bundled.append([bundle[0], *x])
   return evenly_bundled

def insert_name_and_traingroup(ws, even_bundles, train_options_var, name_options_var): # Places first name and train groups onto each name tag as a textbox
   for bundle in even_bundles:  # Find way to do this code with enumerate
      ws.insert_textbox(bundle[0], bundle[1], name_options_var)
   for bundle in even_bundles:
      ws.insert_textbox(bundle[0], bundle[2], train_options_var)

def insert_background_image(ws, row_count, img_var, column=0): # add image param later
   for row in range(row_count):
      ws.embed_image(row, column, img_var)

def insert_back_side_image(ws, row_count, img_var, column=1):
   for row in range(row_count):
      ws.embed_image(row, column, img_var)

def insert_logo(ws, row_count, img, column=0):
   options = get_logo_options()
   for row in range(row_count):
      ws.insert_image(row, column, img, options)

def generate_name_tags(number_of_groups: int, train_options, name_options, back_side_image, logo, *args):  # Considering making parameter to let the user make output file name
   wb = xlsxwriter.Workbook('FinishedNameTags.xlsx')
   number_of_groups_range = number_of_groups + 1
   unpacked_args = [*args]
   index_number = 0
   for group_number in range(1, number_of_groups_range):
      generate_name_tags_by_group(group_number, wb, back_side_image, logo, unpacked_args[index_number], unpacked_args[index_number + 1], get_train_options(*(train_options[group_number - 1])), get_name_options(*(name_options[group_number - 1])))
      index_number += 2
   wb.close()

def generate_name_tags_by_group(group_number, wb, back_side_image, logo, group_list, group_image, train_options, name_options):
   ws = wb.add_worksheet(f'group{str(group_number)}')
   row_count = len(group_list)
   front_cells = get_front_cell_list(row_count)
   set_column(ws)
   set_row(ws, row_count)
   bundles = get_bundles(front_cells, group_list)
   even_bundles = evenly_bundle(bundles)
   insert_name_and_traingroup(ws, even_bundles, train_options, name_options)
   if type(group_image) == int:
       pass
   else:
      insert_background_image(ws, row_count, group_image)
   if back_side_image is None:
      pass
   else:
      insert_back_side_image(ws, row_count, back_side_image)
   if logo is None:
      pass
   else:
      insert_logo(ws, row_count, logo)
