import os
import xml.etree.ElementTree as ET
import pandas as pd
import math
import svgwrite

folder_path = 'name_tags_merged'
output_folder = os.path.join(folder_path)
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print("Folder name_tags_merged created!")
else : print("Folder name_tags_merged exist!")

# Function to convert px to mm (1 px = 0.264583 mm)
def px_to_mm(px):
    return px * 0.264583

# Function to convert mm to px (1 mm = 3.7795275591 px)
def mm_to_px(mm):
    return mm * 3.7795275591

def merge_svg_files(work_area_x, work_area_y, tags_spacing, frame_size_x, frame_size_y):
    input_folder = 'name_tags'
    output_folder = 'name_tags_merged'
    os.makedirs(output_folder, exist_ok=True)
    
    merged_svg = svgwrite.Drawing(size=(f'{work_area_x}px', f'{work_area_y}px'), viewBox=f"0 0 {work_area_x} {work_area_y}", profile='tiny')
    
    # Initialize coordinates and file count
    current_x = 0
    current_y = 0
    file_count = 1
    start_tag = 1
    svg_files = sorted([f for f in os.listdir(input_folder) if f.endswith('.svg')], key=lambda x: int(os.path.splitext(x)[0]))

    # Create a new SVG file
    merged_svg = ET.Element('svg', xmlns="http://www.w3.org/2000/svg", version="1.1")
    merged_svg.set('width', f'{work_area_x}px')
    merged_svg.set('height', f'{work_area_y}px')
    
     # Add green frame around the whole working area
    green_frame = ET.Element('rect', {
        'x': '0',
        'y': '0',
        'width': f'{work_area_x}px',
        'height': f'{work_area_y}px',
        'fill': 'none',
        'stroke': 'green',
        'stroke-width': '1'
    })
    merged_svg.append(green_frame)

    for i, svg_file in enumerate(svg_files):
        tree = ET.parse(os.path.join(input_folder, svg_file))
        root = tree.getroot()

        # Adjust the coordinates of the root elements
        transform = f'translate({current_x},{current_y})'
        for child in root:
            if 'transform' in child.attrib:
                child.attrib['transform'] += f' {transform}'
            else:
                child.attrib['transform'] = transform

        # Append root elements to the merged SVG
        merged_svg.extend(root)

        # Update the coordinates for the next SVG file
        current_x += frame_size_x + tags_spacing
        if current_x + frame_size_x > work_area_x:
            current_x = 0
            current_y += frame_size_y + tags_spacing

        # Check if the current working area is full
        if current_y + frame_size_y > work_area_y or i == len(svg_files) - 1:
            # Save the merged SVG file
            end_tag = i + 1
            output_path = os.path.join(output_folder, f'merged{start_tag}-{end_tag}.svg')
            tree = ET.ElementTree(merged_svg)
            tree.write(output_path)

            # Reset coordinates and create a new SVG file
            current_x = 0
            current_y = 0
            file_count += 1
            start_tag = end_tag + 1
            merged_svg = ET.Element('svg', xmlns="http://www.w3.org/2000/svg", version="1.1")
            merged_svg.set('width', f'{work_area_x}px')
            merged_svg.set('height', f'{work_area_y}px')
            
                        # Add green frame around the whole working area for the new SVG
            green_frame = ET.Element('rect', {
                'x': '0',
                'y': '0',
                'width': f'{work_area_x}px',
                'height': f'{work_area_y}px',
                'fill': 'none',
                'stroke': 'green',
                'stroke-width': '1'
            })
            merged_svg.append(green_frame)

def main():
    # Read parameters from Excel file
    parameters_df = pd.read_excel('parameters.xlsx')

    work_area_x = mm_to_px(parameters_df['workAreaX'][0])
    work_area_y = mm_to_px(parameters_df['workAreaY'][0])
    tags_spacing = mm_to_px(parameters_df['tagsSpacing'][0])
    frame_size_x = mm_to_px(parameters_df['frame_size_x'][0])
    frame_size_y = mm_to_px(parameters_df['frame_size_y'][0])

    merge_svg_files(work_area_x, work_area_y, tags_spacing, frame_size_x, frame_size_y)
    print("SVG files merged successfully!")

    # Prevent the command prompt from closing automatically
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
