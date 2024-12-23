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

def fetch_project_areas():
    """
    Fetches project areas using the API.
    """
    try:
        response = requests.get(api_url_project_areas, auth=(username, password), verify=False)
        response.raise_for_status()

        data = response.json()
        if 'soapenv:Body' in data and "response" in data['soapenv:Body']:
            return data['soapenv:Body']["response"]["returnValue"]["value"]["com.ibm.rqm.planning.service.permissionsWebUIInitializer"]["userProjectAreas"]
        else:
            print("Invalid response structure.")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching project areas: {e}")
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON response: {e}")

    return []

def fetch_oslc_details(project_area_uuid):
    """
    Fetches OSLC details for a given project area and retrieves all streams.
    """
    oslc_url = oslc_api_url_template.replace("{Project_Area_UUID}", project_area_uuid)
    try:
        response = requests.get(oslc_url, auth=(username, password), verify=False)
        response.raise_for_status()

        root = ET.fromstring(response.text)
        result_set_size = int(root.find('.//resultSetSize').text)
        results = root.findall('.//results')

        streams = []
        for result in results[:result_set_size]:
            item_id = result.find("itemId").text if result.find("itemId") is not None else None
            name = result.find("name").text if result.find("name") is not None else None
            if item_id and name:
                streams.append({
                    "Project_Area_Stream_Name": name,
                    "Project_Area_Stream_OSLC_ID": item_id
                })

        return streams
    except ET.ParseError as e:
        print(f"Error parsing XML response for Project Area UUID {project_area_uuid}: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching OSLC details for Project Area UUID {project_area_uuid}: {e}")

    return []

def parse_project_areas(user_project_areas):
    """
    Parses the project areas into a simplified format.
    """
    return [
        {"Project_Area_Name": area.get("name"), "Project_Area_UUID": area.get("itemId")}
        for area in user_project_areas if area.get("name") and area.get("itemId")
    ]

def save_to_excel(project_areas):
    """
    Saves project areas and streams to an Excel file.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Reports/Project_Areas_{timestamp}.xlsx"
    os.makedirs("Reports", exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Project Areas"
    sheet.append(["Project Area Name", "Project Area UUID", "Stream Name", "Stream OSLC ID"])

    for area in project_areas:
        streams = area.get("Streams", [])
        for stream in streams:
            sheet.append([area["Project_Area_Name"], area["Project_Area_UUID"], stream["Project_Area_Stream_Name"], stream["Project_Area_Stream_OSLC_ID"]])
        if not streams:
            sheet.append([area["Project_Area_Name"], area["Project_Area_UUID"], "", ""])

    workbook.save(file_name)
    print(f"Project areas and streams saved to {file_name}")

# Main Execution
user_project_areas = fetch_project_areas()
if user_project_areas:
    project_areas = parse_project_areas(user_project_areas)

    for area in project_areas:
        area["Streams"] = fetch_oslc_details(area["Project_Area_UUID"])

    total_streams = sum(len(area.get("Streams", [])) for area in project_areas)
    print(f"Total number of streams fetched: {total_streams}")

    save_to_excel(project_areas)
else:
    print("Failed to fetch or parse project areas.")
