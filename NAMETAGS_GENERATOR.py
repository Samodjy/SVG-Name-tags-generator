import pandas as pd
import svgwrite
import subprocess
import os

# Read Excel file with name list and other data
df_names = pd.read_excel('name_list.xlsx')  # Assuming the Excel file name is "name_list.xlsx"

# Fill NaN values in 'surname' column with an empty string
#df_names['surname'].fillna('', inplace=True) //from chatgpt
#'df.method({col: value}, inplace=True)' // suggets of modification 
df_names.fillna({'surname': ''}, inplace=True) #modification

# Read Excel file with parameters
df_params = pd.read_excel('parameters.xlsx')  # Assuming the Excel file name is "parameters.xlsx"

# Extract the constant values from the parameters DataFrame
frameX = df_params['frame_size_x'][0]  # Assuming the column name for frame width is "frame_size_x"
frameY = df_params['frame_size_y'][0]  # Assuming the column name for frame height is "frame_size_y"
font_size = df_params['font_size'][0]  # Assuming the column name for font size is "font_size"
font_type = df_params['font_type'][0]  # Assuming the column name for font type is "font_type"
stroke_width = df_params['stroke_width'][0]  # Assuming the column name for stroke width is "stroke_width"
ink_path = df_params['ink_path'][0]

# Conversion factor
px_per_mm = 3.7795275591  # 1 mm = 3.7795275591 px

# Convert frame dimensions to px if they are specified in mm
frameX_px = frameX * px_per_mm
frameY_px = frameY * px_per_mm


output_folder = os.path.join('name_tags')
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print("Folder name_tags created!")
else : print("Folder name_tags exist!")

# Define function to generate SVG paths for text
def generate_svg(room, name, surname):
    # Capitalize first letter of name and surname
    name = name.capitalize()
    surname = surname.capitalize() if surname else ''
    
    # Add leading zeros to room number if necessary
    room_padded = str(room).zfill(3)  # Assuming room numbers are padded to 3 digits
    
    # Create SVG Drawing
    dwg = svgwrite.Drawing(size=(f'{frameX_px}px', f'{frameY_px}px'), profile='tiny')
    
    # Add blue frame
    dwg.add(dwg.rect(insert=(0, 0), size=(f'{frameX_px}px', f'{frameY_px}px'), fill='none', stroke='blue', stroke_width=f'{stroke_width}px'))
    
    # Calculate vertical center for name and surname
    frame_center_x = (frameX_px / 2)  # Center in px
    if surname:
        name_center_y = (frameY_px - (font_size/3)) / 2
        surname_center_y = (frameY_px / 2) + font_size
    else:
        name_center_y = (frameY_px / 2) + (font_size / 3)  # Center name vertically if surname is empty
    
    # Generate SVG paths for name
    name_paths = dwg.text(name, insert=(f'{frame_center_x}px', f'{name_center_y}px'), text_anchor='middle', font_size=f'{font_size}px', font_family=font_type, fill="red")
    
    # Generate SVG paths for surname if it's not empty
    if surname:
        surname_paths = dwg.text(surname, insert=(f'{frame_center_x}px', f'{surname_center_y}px'), text_anchor='middle', font_size=f'{font_size}px', font_family=font_type, fill="red")
    
    # Add paths to SVG group
    group = dwg.g()
    group.add(name_paths)
    if surname:
        group.add(surname_paths)
    
    # Write SVG content to file
    dwg.add(group)
    dwg.saveas(os.path.join('name_tags', f'{room_padded}.svg'))
    
    # Convert SVG text to paths using Inkscape
    output_folder = 'name_tags'    
    subprocess.run([ink_path, '--export-type=svg', '--export-text-to-path', f'{room_padded}.svg'], cwd=output_folder)
    print("SVG files created successfully!")

# Generate SVG files for each entry in the Excel file
for index, row in df_names.iterrows():
    room = row['room']
    name = row['name']
    surname = row['surname']
    
    # Add leading zeros to room number if necessary
    room_padded = str(room).zfill(3)  # Assuming room numbers are padded to 3 digits
    
    # Generate SVG paths for text and write to file
    generate_svg(room, name, surname)
    # Remove the original SVG file
    os.remove(os.path.join('name_tags', f'{room_padded}.svg'))
    
# Specify the folder path
folder_path = 'name_tags'

# Iterate over each file in the folder
for filename in os.listdir(folder_path):
    # Check if the filename contains the "_out" appendix
    if '_out' in filename:
        # Remove the "_out" part from the filename
        new_filename = filename.replace('_out', '')
        # Rename the file
        os.rename(os.path.join(folder_path, filename), os.path.join(folder_path, new_filename))

# Prevent the command prompt from closing automatically
input("Press Enter to exit...")
