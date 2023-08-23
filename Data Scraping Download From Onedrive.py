#!/usr/bin/env python
# coding: utf-8

# In[31]:


import zipfile
import os
import shutil
import time
import requests
import asyncio
import nest_asyncio
from pyppeteer import launch


# In[32]:


folder_path = r'D:\audra'

# Check if the folder path exists
if os.path.exists(folder_path):
    # Get a list of all filenames in the folder
    filenames = os.listdir(folder_path)

    # Initialize a set to store the unique substrings found in filenames
    filename_substrings = set()

    # Extract substrings from filenames and add them to the set
    for filename in filenames:
        substrings = filename.split()
        filename_substrings.update(substrings)
    
    lokal = []
    # Print the filenames containing both "DR" and "DH"
    for filename in filenames:
        if "DR DH" in filename:
            lokal.append(filename)


# In[33]:


path_d = r'D:\audra'

# Check if the folder path exists
if os.path.exists(path_d):
    # Get a list of all filenames in the folder
    zipnyam = os.listdir(path_d)


# In[34]:


path_ = r'C:\Users\ASUS\Downloads'

# Check if the folder path exists
if os.path.exists(path_):
    # Get a list of all filenames in the folder
    a = os.listdir(path_)


# In[35]:


non = []
ya = []
# Allow running asyncio code in Jupyter Notebook
nest_asyncio.apply()

async def main():
    browser = await launch(headless=False)
    page = await browser.newPage()
    
    # goto page sharepoint
    await page.goto("--- Link that you want ---")

    # goto 01. DPR
    await page.waitForSelector('.ms-Breadcrumb-list')
    await page.click('button[title="01. DPR"]')

    # Use asyncio.sleep instead of page.waitForTimeout
    await asyncio.sleep(2)

    # goto 03. EKP
    await page.click('button[title="03. EKP"]')

    # Use asyncio.sleep instead of page.waitForTimeout
    await asyncio.sleep(1)

    # select all rows with same class
    listFileSelector = await page.querySelectorAll('[data-automationid="DetailsRowCheck"]')
    for element in listFileSelector:
        title = await page.evaluate('(element) => element.title', element)
        if title not in lokal and title.startswith('DR DH'):
            # Add to variable
            non.append(title)
            await asyncio.sleep(2)
            # Click on the element with title "DH DPR"
            if non == ya:
                    print('No New File')
                else:
                    await element.click()
                    await page.click('button[name="Download"]')
                
# Run the main function
asyncio.get_event_loop().run_until_complete(main())
#         download_path = [os.path.join(r'C:\Users\ASUS\Downloads', x) for x in a if x not in a]

#         while any("crdownload" in item for item in download_path):
#             await asyncio.sleep(1)  # Wait for a second before checking again


#         # Poll for the presence of the downloaded file
#         # temp_list = os.listdir(r'C:\Users\ASUS\Downloads')
#         # download_path = [os.path.join(r'C:\Users\ASUS\Downloads',x) for x in a if x not in temp_list]
#         # for jan in download_path:
#         # #    while not os.path.exists(download_path):
#         #     while any("crdownload" for item in download_path) == False:
#         #         await asyncio.sleep(1)  # Wait for a second before checking again

#         print("Download complete!")
#         # # print(x)
#         # print("Download complete!")
#         # print(a)
#         # print(download_path)

        # await browser.close()
        # return download_path

        
# Call the main function, but since we're in a Jupyter Notebook, asyncio.run is not used
# await main()


# In[36]:


if os.path.exists(path_):
    # Get a list of all filenames in the folder
    b = os.listdir(path_)
    
    # Initialize a set to store the unique substrings found in filenames
    filename_substrings = set()
    
beda = [x for x in b if x not in a ]


# In[1]:


for value in beda:
    # Source file path
        if value.endswith('.zip'):
            source_file_path = f'C:\\Users\\ASUS\\Downloads\\{value}'  # Replace with the actual source file path

            # Destination directory path
            destination_directory = r'D:\audra'  # Replace with the desired destination directory path

            # Check if the destination directory exists, if not, create it
            if not os.path.exists(destination_directory):
                os.makedirs(destination_directory)

             # Move the file to the destination directory
                shutil.move(source_file_path, destination_directory)

            def unzip_file(zip_file_path, target_directory):
            # Check if the zip file exists
                if not zipfile.is_zipfile(zip_file_path):
                    print("Error: The provided file is not a valid zip file.")
                    return

                # Check if the target directory exists, if not, create it
                if not os.path.exists(target_directory):
                    os.makedirs(target_directory)

                with zipfile.ZipFile(zip_file_path) as zip_ref:
                    zip_ref.extractall(target_directory)

                print(f"Unzipping {zip_file_path} completed successfully!")

            for value in beda:
                zip_file_path = f'D:\\audra\\{value}'
                target_directory = r'D:\audra'

                unzip_file(zip_file_path, target_directory)

                file_path = f'D:\\audra\\{value}'  # Replace this with the actual file path

                try:
                    # Use os.remove() to delete the file
                    os.remove(file_path)
                    print(f"File '{file_path}' has been deleted successfully.")
                except FileNotFoundError:
                    print(f"File '{file_path}' not found.")
                except Exception as e:
                    print(f"Error occurred while deleting the file: {e}")

        else:
            # Source file path
            source_file_path = f'C:\\Users\\ASUS\\Downloads\\{value}'  # Replace with the actual source file path

            # Destination directory path
            destination_directoryz = r'D:\audra'  # Replace with the desired destination directory path

            # Check if the destination directory exists, if not, create it
            if not os.path.exists(destination_directoryz):
                os.makedirs(destination_directoryz)

            # Move the file to the destination directory
            shutil.move(source_file_path, destination_directoryz)
            print ('xlsx')


# In[8]:


for value in beda: 
    file_pathx = f'D:\\audra\\{value}'  # Replace this with the actual file path
    if value.endswith(').xlsx') :
        # Use os.remove() to delete the file
        os.remove(file_pathx)
    else:
        print(f"File '{value}' has been deleted successfully.")

