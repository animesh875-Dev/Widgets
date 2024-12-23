import json
import requests
from openpyxl import Workbook
import urllib3
from datetime import datetime
import os
import xml.etree.ElementTree as ET

# Suppress warnings about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load configuration from config.json
with open("config.json", "r") as config_file:
    config = json.load(config_file)

username = config["username"]
password = config["password"]
server_url = config["server_url"]

# API endpoints
api_url_project_areas = f"{server_url}/qm/service/com.ibm.team.repository.service.internal.webuiInitializer.IWebUIInitializerRestService/initializationData"
oslc_api_url_template = f"{server_url}/qm/service/com.ibm.rqm.configmanagement.service.rest.IConfigurationManagementRestService/pagedSearchResult?pageSize=100&page=0&projectArea={{Project_Area_UUID}}"

# Function to fetch project areas
def fetch_project_areas():
    try:
        response = requests.get(api_url_project_areas, auth=(username, password), verify=False)
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

# Function to fetch OSLC details for a project area
def fetch_oslc_details(project_area_uuid):
    oslc_url = oslc_api_url_template.replace("{Project_Area_UUID}", project_area_uuid)
    try:
        response = requests.get(oslc_url, auth=(username, password), verify=False)
        
        # Log the response status and raw content
        print(f"Response Status Code: {response.status_code}")
        print(f"Response Content: {response.text}")
        
        response.raise_for_status()  # Check if the request was successful
        
        # Parse the XML response
        try:
            root = ET.fromstring(response.text)
            
            # Find the rootStream section
            root_stream = root.find(".//rootStream")
            if root_stream is not None:
                item_id = root_stream.find("itemId").text if root_stream.find("itemId") is not None else None
                name = root_stream.find("name").text if root_stream.find("name") is not None else None
                
                # Print the itemId and name for debugging
                print(f"Root Stream - itemId: {item_id}, name: {name}")
                
                return [{"Project_Area_Stream_Name": name, "Project_Area_stream_OSLC_ID": item_id}]
            else:
                print("rootStream not found in the response.")
                return []
        except ET.ParseError as e:
            print(f"Error parsing XML response for Project Area UUID {project_area_uuid}: {e}")
            return []

    except requests.exceptions.RequestException as e:
        print(f"Error fetching OSLC details for Project Area UUID {project_area_uuid}: {e}")
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON response for Project Area UUID {project_area_uuid}: {e}")

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

# Function to save project areas and streams to an Excel file
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
    sheet.append(["Project Area Name", "Project Area UUID", "Project Area Stream Name", "Project Area Stream OSLC ID"])

    # Write project areas and streams to Excel
    for area in project_areas:
        streams = area.get("Streams", [])
        if streams:
            for stream in streams:
                sheet.append([area["Project_Area_Name"], area["Project_Area_UUID"], stream["Project_Area_Stream_Name"], stream["Project_Area_stream_OSLC_ID"]])
        else:
            sheet.append([area["Project_Area_Name"], area["Project_Area_UUID"], "", ""])

    workbook.save(file_name)
    print(f"Project areas and streams saved to {file_name}")

# Main Execution
user_project_areas = fetch_project_areas()
if user_project_areas:
    project_areas = parse_project_areas(user_project_areas)

    # Fetch OSLC details for each project area
    for area in project_areas:
        area["Streams"] = fetch_oslc_details(area["Project_Area_UUID"])

    # Display the total count in log
    total_streams = sum(len(area.get("Streams", [])) for area in project_areas)
    print(f"Total number of project area streams fetched: {total_streams}")

    # Save project areas and streams to an Excel file
    save_to_excel(project_areas)
else:
    print("Failed to fetch or parse project areas.")

# Code is fetching only the default stream 