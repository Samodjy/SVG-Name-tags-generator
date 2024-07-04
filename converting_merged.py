import os
import subprocess

def convert_svg_files(folder_path, inkscape_path):
    # Ensure the output folder exists
    output_folder = os.path.join(folder_path, 'converted')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Iterate over each SVG file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".svg"):
            input_path = os.path.join(folder_path, filename)
            output_path = os.path.join(output_folder, filename)

            # Command to convert the SVG file using Inkscape
            subprocess.run([inkscape_path, input_path, '--export-plain-svg', output_path])

            print(f"Converted {filename} and saved as {output_path}")

    print("All SVG files have been converted and saved.")

if __name__ == "__main__":
    folder_path = 'name_tags_merged'  # Specify your folder containing the SVG files
    inkscape_path = r'C:\Program Files\Inkscape\bin\inkscape.exe'  # Specify the full path to Inkscape executable

    convert_svg_files(folder_path, inkscape_path)
    input("Press Enter to close the program...")
