import json
import requests
import urllib3
import xml.etree.ElementTree as ET

# Suppress warnings about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load configuration from config.json
with open("config.json", "r") as config_file:
    config = json.load(config_file)

# Extract details from config
username = config["username"]
password = config["password"]
server_url = config["server_url"]
project_area_id = config["project_area_id"]

# Construct the request URL
url = f"{server_url}/qm/oslc_qm/contexts/{project_area_id}/resources/com.ibm.rqm.planning.VersionedTestCase"

# Send the GET request with basic authentication
response = requests.get(url, auth=(username, password), verify=False)

# Check if the request was successful
if response.status_code == 200:
    # Parse the XML response
    root = ET.fromstring(response.content)
    # Find the oslc:totalCount element
    namespace = {"oslc": "http://open-services.net/ns/core#"}
    total_count_element = root.find(".//oslc:totalCount", namespace)
    if total_count_element is not None:
        total_count = total_count_element.text
        print(f"Total Count: {total_count}")
    else:
        print("Total count not found in the response.")
else:
    print(f"Failed to fetch data. Status code: {response.status_code}, Response: {response.text}")
