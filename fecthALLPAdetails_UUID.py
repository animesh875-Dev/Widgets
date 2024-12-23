import json
import requests
from openpyxl import Workbook
import urllib3
from datetime import datetime
import os

# Suppress warnings about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load configuration from config.json
with open("config.json", "r") as config_file:
    config = json.load(config_file)

username = config["username"]
password = config["password"]
server_url = config["server_url"]

# API endpoint
api_url = f"{server_url}/qm/service/com.ibm.team.repository.service.internal.webuiInitializer.IWebUIInitializerRestService/initializationData"

# Function to fetch project areas
def fetch_project_areas():
    try:
        response = requests.get(api_url, auth=(username, password), verify=False)
        response.raise_for_status()  # Check if the request was successful

        # Parse the response as JSON
        data = response.json()

        # Check if 'soapenv:Body' exists in the response
        if 'soapenv:Body' in data:
            body = data['soapenv:Body']
            if "response" in body:
                return body["response"]["returnValue"]["value"]["com.ibm.rqm.planning.service.permissionsWebUIInitializer"]["userProjectAreas"]
            else:
                print("Key 'response' not found in the response")
        else:
            print("Key 'soapenv:Body' not found in the response")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching project areas: {e}")
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON response: {e}")

    return []

# Function to parse JSON and extract project area names and UUIDs
def parse_project_areas(user_project_areas):
    project_areas = []

    # Extract project area names and UUIDs
    for item in user_project_areas:
        name = item.get("name")
        item_id = item.get("itemId")
        if name and item_id:
            project_areas.append({"Project_Area_Name": name, "Project_Area_UUID": item_id})

    return project_areas

# Function to save project areas to an Excel file
def save_to_excel(project_areas):
    # Create a timestamped file name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Reports/Project_Areas_{timestamp}.xlsx"

    # Ensure the Reports folder exists
    os.makedirs("Reports", exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Project Areas"

    # Add header
    sheet.append(["Project Area Name", "Project Area UUID"])

    # Write project areas to Excel
    for area in project_areas:
        sheet.append([area["Project_Area_Name"], area["Project_Area_UUID"]])

    workbook.save(file_name)
    print(f"Project areas saved to {file_name}")

# Main Execution
user_project_areas = fetch_project_areas()
if user_project_areas:
    project_areas = parse_project_areas(user_project_areas)
    project_area_count = len(project_areas)

    # Display the total count in log
    print(f"Total number of project areas fetched: {project_area_count}")

    # Save project areas to an Excel file
    save_to_excel(project_areas)
else:
    print("Failed to fetch or parse project areas.")
