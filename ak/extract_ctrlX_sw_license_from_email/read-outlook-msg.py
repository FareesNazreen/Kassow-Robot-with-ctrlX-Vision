# Exercise: 
# 1. How to read Outlook email messages that contains ctrlX OS Software Licenses (SWL) 
# 2. Extract the string from Entitlement ID, Activation ID and Software Product 
# 3. Save it as MS Excel file

import os
import re
import extract_msg
import pandas as pd

#
# Get names of all messages in the folder
#
# Set the folderpath variable to the path of a folder with the .msg file you want to process
folderpath = "CTRLX SOFTWARE LICENSE/"

# 
# Regular expressions for Entitlement ID, Activation ID and Software Product
# 
entitlement_id_pattern = r"Entitlement ID:\s*([\w-]+)"
activation_id_pattern = r"Expiration Date\s+([\w-]+)"
software_product_pattern = r"(ctrlX OS License - [^[]+\[Version: [^\]]+\])"


# Get list of .msg files in the folder
f = []
for (dirpath, dirnames, filenames) in os.walk(folderpath):
    for file in filenames:
        if file.endswith(".msg"):
            f.append(file)

# print(f)

# Create a list to store extracted data
data = []

# Loop through each .msg file and extract data
for email_file in f:
    print(f"Processing: {email_file}")

    try:
        # Open and read the .msg file
        msg = extract_msg.Message(os.path.join(folderpath, email_file))
        msg_message = msg.body

        # print(msg_message)
        # print(msg_message[384:425]) # Entitlement ID No is within this range

        # # Extract email sender
        # print(f"From: {msg.sender}")

        # # Search for Entitlement ID and Activation ID in the email body
        entitlement_id_match = re.search(entitlement_id_pattern, msg_message)
        activation_id_match = re.search(activation_id_pattern, msg_message)
        software_product_match = re.search(software_product_pattern, msg_message)

        entitlement_id = entitlement_id_match.group(1) if entitlement_id_match else "Not Found"
        activation_id = activation_id_match.group(1) if activation_id_match else "Not Found"
        software_product = software_product_match.group(1) if software_product_match else "Not Found"

        # print(f"Entitlement ID: {entitlement_id}")
        # print(f"Activation ID: {activation_id}")
        # print(f"Software Product: {software_product}")

        # Append the extracted data as a tuple to the data list
        data.append((entitlement_id, activation_id, software_product))

    except Exception as e:
        print(f"Can't open or process email {email_file}: {e}")

# Create a DataFrame from the extracted data
df = pd.DataFrame(data, columns=["Entitlement ID", "Activation ID", "Software Product"])

# Save the DataFrame to an Excel file
output_file = "extracted_ctrlX_software_licenses.xlsx"
df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")