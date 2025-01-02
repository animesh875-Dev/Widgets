import requests
import xml.etree.ElementTree as ET

# Configuration details (to be replaced with your actual details)
config = {
    "username": "dan7kor",
    "password": "Password@875",
    "server_url": "https://rb-alm-11-p.de.bosch.com",
    "data_governance_TP": 200,
    "data_governance_TC": 5000,
    "project_area_id": "_Lx7fEHaQEeeHQLB3qMZX2g",
    "Project_Area_Stream_OSLC_ID": "_N4VyNHaQEeeHQLB3qMZX2g"
}

api_url = f"{config['server_url']}/qm/service/com.ibm.rqm.planning.common.service.rest.ITestCaseRestService/pagedSearchResult"
username = config["username"]
password = config["password"]

# Headers
headers = {
    "Content-Type": "application/x-www-form-urlencoded; charset=utf-8",
    "Accept": "application/json",
}

def build_request_body(page=0, page_size=100, process_area=""):
    """
    Build the body of the request dynamically based on parameters.
    """
    body = {
        "includeCustomAttributes": "true",
        "includeArchived": "false",
        "processArea": process_area,
        "traceabilityViewType": "true",
        "resolveParentTestPlans": "false",
        "resolveScripts": "false",
        "resolveParentTestSuites": "true",
        "resolveCategories": "true",
        "resolveCustomAttributes": "false",
        "resolveLinkedFiles": "false",
        "resolveDevItem": "false",
        "resolveCopiedArtifactInfo": "false",
        "page": str(page),
        "pageSize": str(page_size),
        "resultLimit": "-1",
        "oslc_config.context": config["Project_Area_Stream_OSLC_ID"],
        "isWebUI": "true",
    }
    return "&".join(f"{key}={value}" for key, value in body.items())

def fetch_total_size(body):
    """
    Sends a POST request to the API and extracts the <totalSize> value from the response.
    """
    try:
        # Make the POST request
        response = requests.post(api_url, data=body, headers=headers, auth=(username, password), verify=False)
        response.raise_for_status()

        # Parse the XML response
        root = ET.fromstring(response.text)

        # Find the <totalSize> element
        total_size_element = root.find(".//totalSize")
        if total_size_element is not None:
            total_size = int(total_size_element.text)
            print(f"Total Size: {total_size}")
            return total_size
        else:
            print("No <totalSize> element found in the response.")
            return None

    except requests.exceptions.RequestException as e:
        print(f"Error making API request: {e}")
        return None
    except ET.ParseError as e:
        print(f"Error parsing XML response: {e}")
        return None

# Example usage
body = build_request_body(page=1, page_size=50, process_area=config["project_area_id"])
total_size = fetch_total_size(body)
if total_size is not None:
    print(f"Successfully fetched total size: {total_size}")
else:
    print("Failed to fetch total size.")
