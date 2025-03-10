import xml.etree.ElementTree as ET
import base64
import re
import struct
import os
import shutil

# Specify the source and destination folders
source_folder = ""  # Replace with your source folder (e.g., "C:/source")
dest_folder = ""  # Replace with your destination folder (e.g., "C:/destination") # Ensure the destination folder exists

os.makedirs(dest_folder, exist_ok=True)

# Namespace (adjust if your XML uses a different one)
namespace = {'my': 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2006-04-19T07:22:55'}

# Find attachment fields (e.g., felt31, felt32, etc.)
attachment_pattern = re.compile(r'felt\d+')

def decode_infopath_attachment(binary_data):
    try:
        # Check if data is long enough for header
        if len(binary_data) < 24:
            raise ValueError("Binary data too short for InfoPath header")
        
        # Read header size (bytes 0-3, typically 24)
        header_size = struct.unpack('<I', binary_data[0:4])[0]
        if header_size != 24:
            print(f"Warning: Unexpected header size {header_size}, expected 24")
        
        # Read filename length (bytes 20-23, in characters)
        filename_length = struct.unpack('<I', binary_data[20:24])[0]
        
        # Calculate filename byte length (Unicode, 2 bytes per char)
        filename_bytes = filename_length * 2
        filename_start = 24
        filename_end = filename_start + filename_bytes
        
        # Extract filename (decode Unicode)
        filename = binary_data[filename_start:filename_end].decode('utf-16-le').rstrip('\x00')
        
        # File content starts after header + filename
        file_content = binary_data[filename_end:]
        
        return filename, file_content
    except Exception as e:
        print(f"Error decoding attachment structure: {e}")
        return None, None

def process_xml_file(file_path, output_dir):
    try:
        # Load the InfoPath XML file
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Create a subfolder named after the XML file (without .xml)
        xml_name = os.path.splitext(os.path.basename(file_path))[0]  # e.g., "O.81" from "O.81.xml"
        xml_output_dir = os.path.join(output_dir, xml_name)
        os.makedirs(xml_output_dir, exist_ok=True)
        
        # Copy the original XML file to the subfolder
        dest_xml_path = os.path.join(xml_output_dir, os.path.basename(file_path))  # e.g., "O.81.xml"
        shutil.copy2(file_path, dest_xml_path)
        print(f"Copied {file_path} to {dest_xml_path}")
        
        attachment_count = 1
        
        for elem in root.findall('.//*', namespaces=namespace):
            tag_name = elem.tag.split('}')[-1]  # Strip namespace prefix
            if attachment_pattern.match(tag_name):
                if elem.text is not None:
                    base64_string = elem.text.strip()
                    if base64_string:
                        try:
                            # Decode base64 to binary
                            binary_data = base64.b64decode(base64_string)
                            
                            # Debug: Print first few bytes
                            print(f"{file_path} - {tag_name} first 10 bytes: {binary_data[:10].hex()}")
                            
                            # Decode InfoPath attachment
                            filename, file_content = decode_infopath_attachment(binary_data)
                            
                            if filename and file_content:
                                # Prefix the filename with the XML name
                                prefixed_filename = f"{xml_name} - {filename}"  # e.g., "O.81 - filename.pdf"
                                output_file = os.path.join(xml_output_dir, prefixed_filename)
                                
                                # Avoid overwriting by adding a number if file exists
                                base_name, ext = os.path.splitext(output_file)
                                counter = 1
                                while os.path.exists(output_file):
                                    output_file = f"{base_name}_{counter}{ext}"
                                    counter += 1
                                with open(output_file, 'wb') as f:
                                    f.write(file_content)
                                print(f"Saved {output_file}")
                                attachment_count += 1
                            else:
                                print(f"Skipping {tag_name} in {file_path}: Could not decode attachment")
                        except Exception as e:
                            print(f"Error processing {tag_name} in {file_path}: {e}")
                    else:
                        print(f"Skipping {tag_name} in {file_path}: Empty content")
                else:
                    print(f"Skipping {tag_name} in {file_path}: No text content")
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")

# Recursively process all XML files in the source folder
for root_dir, _, files in os.walk(source_folder):
    for file in files:
        if file.lower().endswith('.xml'):  # Only process .xml files
            file_path = os.path.join(root_dir, file)
            print(f"\nProcessing {file_path}")
            process_xml_file(file_path, dest_folder)

print("\nRecursive extraction complete.")