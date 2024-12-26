import json
import requests
from openpyxl import Workbook
import urllib3
from datetime import datetime
import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox

# Suppress warnings about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load configuration from config.json
with open("config.json", "r") as config_file:
    config = json.load(config_file)

username = config["username"]
password = config["password"]
server_url = config["server_url"]
data_governance_TP = config["data_governance_TP"]

# API endpoints
api_url_project_areas = f"{server_url}/qm/service/com.ibm.team.repository.service.internal.webuiInitializer.IWebUIInitializerRestService/initializationData"
oslc_api_url_template = f"{server_url}/qm/service/com.ibm.rqm.configmanagement.service.rest.IConfigurationManagementRestService/pagedSearchResult?pageSize=100&page=0&projectArea={{Project_Area_UUID}}"

# File to save printed messages
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
MESSAGE_LOG_FILE = f"Reports/DataGovernanceReport_{timestamp}.txt"

def log_message_to_file(message):
    """
    Logs the provided message to a text file.
    """
    os.makedirs("Reports", exist_ok=True)
    with open(MESSAGE_LOG_FILE, "a") as file:
        file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

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

def fetch_test_plan_count(project_area_uuid, oslc_id):
    """
    Fetches the test plan count for a given project area and OSLC ID.
    """
    test_plan_api_url = f"{server_url}/qm/service/com.ibm.rqm.planning.common.service.rest.ITestPlanRestService/pagedSearchResult?processArea={project_area_uuid}&page=0&pageSize=500&oslc_config.context={oslc_id}"
    try:
        response = requests.get(test_plan_api_url, auth=(username, password), verify=False)
        response.raise_for_status()

        root = ET.fromstring(response.text)
        result_set_size = int(root.find('.//resultSetSize').text)
        return result_set_size
    except ET.ParseError as e:
        print(f"Error parsing XML response for Test Plan Count: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Error fetching Test Plan Count: {e}")

    return 0

def parse_project_areas(user_project_areas):
    """
    Parses the project areas into a simplified format.
    """
    return [
        {"Project_Area_Name": area.get("name"), "Project_Area_UUID": area.get("itemId")}
        for area in user_project_areas if area.get("name") and area.get("itemId")
    ]

def on_component_select(event):
    """
    Fetch and update test plans based on selected component.
    """
    selected_component = component_combobox.get()
    selected_project_area = project_area_combobox.get()
    if selected_component and selected_project_area:
        project_area_uuid = next(area["Project_Area_UUID"] for area in project_areas if area["Project_Area_Name"] == selected_project_area)
        selected_oslc_id = next(comp["Project_Area_Stream_OSLC_ID"] for comp in components if comp["Project_Area_Stream_Name"] == selected_component)

        # Fetch test plan count
        test_plan_count = fetch_test_plan_count(project_area_uuid, selected_oslc_id)
        max_test_plan_allowed = config["data_governance_TP"]

        # Display message based on test plan count
        if max_test_plan_allowed > test_plan_count:
            remaining = max_test_plan_allowed - test_plan_count
            message = (
                f"Project is allowed to create test plans.\n"
                f"Current Test Plan Count: {test_plan_count}\n"
                f"Max Test Plan Allowed: {max_test_plan_allowed}\n"
                f"Remaining Test Plans: {remaining}"
            )
            messagebox.showinfo("Test Plan Creation Allowed", message)
            log_message_to_file(message)
        else:
            message = (
                f"Test Plan creation limit reached.\n"
                f"Current Test Plan Count: {test_plan_count}\n"
                f"Max Test Plan Allowed: {max_test_plan_allowed}"
            )
            messagebox.showwarning("Test Plan Creation Limit Reached", message)
            log_message_to_file(message)

def on_test_plan_select():
    """
    Handles the logic when the 'Select Test Plan' button is clicked.
    """
    selected_component = component_combobox.get()
    if selected_component:
        messagebox.showinfo("Test Plan Selection", f"Test Plan Selected for Component: {selected_component}")
        log_message_to_file(f"Test Plan Selected for Component: {selected_component}")

def on_project_area_select(event):
    """
    Fetch and update components based on selected project area.
    """
    selected_project_area = project_area_combobox.get()
    if selected_project_area:
        project_area_uuid = next(area["Project_Area_UUID"] for area in project_areas if area["Project_Area_Name"] == selected_project_area)
        global components
        components = fetch_oslc_details(project_area_uuid)

        # Update the components dropdown
        component_combobox['values'] = [comp["Project_Area_Stream_Name"] for comp in components]
        component_combobox.set('')  # Clear previous selection
        component_combobox.grid(row=2, column=1, padx=10, pady=10)
        component_combobox.bind("<<ComboboxSelected>>", on_component_select)

        # Display the selected project area name
        selected_project_area_label.config(text=f"Selected Project Area: {selected_project_area}")

# Set up Tkinter window
window = tk.Tk()
window.title("Project Area and Stream Selector")
window.geometry("800x600")

# Fetch project areas and parse them
user_project_areas = fetch_project_areas()
if user_project_areas:
    project_areas = parse_project_areas(user_project_areas)

    # Project Area selection dropdown
    project_area_combobox = ttk.Combobox(window, values=[area["Project_Area_Name"] for area in project_areas], state="readonly", width=40)
    project_area_combobox.grid(row=0, column=1, padx=10, pady=10)
    project_area_combobox.bind("<<ComboboxSelected>>", on_project_area_select)

    # Label for project area selection
    project_area_label = tk.Label(window, text="Select Project Area:")
    project_area_label.grid(row=0, column=0, padx=10, pady=10)

    # Label for selected project area
    selected_project_area_label = tk.Label(window, text="Selected Project Area: ")
    selected_project_area_label.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

    # Components dropdown
    component_combobox = ttk.Combobox(window, state="readonly", width=40)
    component_combobox.grid(row=2, column=1, padx=10, pady=10)

    # Label for components selection
    component_label = tk.Label(window, text="Select Components:")
    component_label.grid(row=2, column=0, padx=10, pady=10)

    # Add a button to select test plan
    test_plan_button = tk.Button(window, text="Select Test Plan", command=on_test_plan_select)
    test_plan_button.grid(row=3, column=1, padx=10, pady=10)

    # Start Tkinter main loop
    window.mainloop()

    # Save data to Excel after the UI is closed
    # save_to_excel(project_areas)
else:
    messagebox.showerror("Error", "Failed to fetch project areas.")
