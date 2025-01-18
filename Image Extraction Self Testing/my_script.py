# Import necessary modules for file operations and Excel file handling
import zipfile  # For handling ZIP files, specifically for extracting images from Excel files
import os  # For operating system-specific functionalities, such as creating directories and file paths
import openpyxl  # For reading and manipulating Excel files
import nltk
# Define the path to the Excel file
excel_file = 'my.xlsx'

# Create a directory to save extracted images, ensuring it exists
output_dir = 'extracted_images'
os.makedirs(output_dir, exist_ok=True)

# Open the Excel file as a ZIP to directly extract images
with zipfile.ZipFile(excel_file, 'r') as zipf:
    # Iterate through each file in the ZIP to locate and extract images
    for file_info in zipf.infolist():
        # Check if the file is an image (png, jpg, or jpeg) located in 'xl/media/'
        if file_info.filename.startswith('xl/media/') and file_info.filename.endswith(('png', 'jpg', 'jpeg')):
            # Extract the image name from the file path
            image_name = os.path.basename(file_info.filename)
            # Construct the full path to save the image
            image_path = os.path.join(output_dir, image_name)
            # Open the file in binary write mode to save the image
            with open(image_path, 'wb') as img_file:
                # Write the image data to the file
                img_file.write(zipf.read(file_info.filename))

# Load the Excel workbook to map images to rows based on tags
workbook = openpyxl.load_workbook(excel_file)
# Select the active worksheet
sheet = workbook.active

# Iterate through each file in the output directory to rename images based on tags
for img_file in os.listdir(output_dir):
    # Construct the full path to the image file
    img_path = os.path.join(output_dir, img_file)

    # Extract the row number from the image file name (e.g., 'image32.png')
    row_number = int(''.join(filter(str.isdigit, img_file)))

    # Adjust row number to account for the headings in the Excel sheet
    adjusted_row = row_number + 1  # Assuming the first row contains headings

    # Retrieve the tag from the Excel sheet (assuming tags are in Column B)
    tag = sheet[f'B{adjusted_row}'].value

    # If a tag exists, rename the image file
    if tag:
        # Construct the new path for the image file using the tag
        new_path = os.path.join(output_dir, f"{tag}.png")

        # Check if the new file name already exists and append a suffix if necessary
        counter = 1
        while os.path.exists(new_path):
            # If the file exists, append a counter to the file name
            new_path = os.path.join(output_dir, f"{tag}_{counter}.png")
            counter += 1

        # Rename the image file to the new path
        os.rename(img_path, new_path)

# Print a success message after completing the process
print("Images extracted and renamed successfully!")
