import json
import requests
from requests.auth import HTTPBasicAuth
import urllib3
import logging

# Suppress warnings about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Load configuration from config.json
try:
    with open("config.json", "r") as config_file:
        config = json.load(config_file)
    logger.info("Configuration loaded successfully.")
except Exception as e:
    logger.error(f"Error loading configuration: {e}")
    raise

# Configuration variables
username = config["username"]
password = config["password"]
server_url = config["server_url"]
project_area_id = config["project_area_id"]
project_area_stream_oslc_id = config["data_governance_TP"]

# Define the API endpoint for the POST request
api_url = f"{server_url}/qm/service/com.ibm.rqm.planning.common.service.rest.ITestCaseRestService/pagedSearchResult?oslc_config.context={project_area_stream_oslc_id}&webContext.projectArea={project_area_id}"
logger.info(f"API URL: {api_url}")

# Define the headers (sending JSON format instead of form-encoded)
headers = {
    "Accept": "application/json",  # Accepting JSON response
    "Content-Type": "application/json",  # Content-Type changed to application/json
}

# Define the payload for the POST request (adjust the query according to your needs)
# payload = {
#     "query": "SELECT * FROM TestCase",  # Example query; adjust according to the server's requirements
#     "maxResults": 500
# }

# Make the POST request to fetch the test case data
logger.info("Sending POST request to the API...")
try:
    response = requests.post(api_url, auth=HTTPBasicAuth(username, password), headers=headers,  verify=False)
    logger.info(f"Request sent. Status code: {response.status_code}")
except Exception as e:
    logger.error(f"Error sending POST request: {e}")
    raise

# Check if the request was successful
if response.status_code == 200:
    try:
        # Parse the JSON response
        logger.debug("Parsing response...")
        data = response.json()

        # Extract the totalSize from the response
        if 'soapenv:Body' in data:
            body = data['soapenv:Body']
            if 'response' in body and 'returnValue' in body['response']:
                total_size = body['response']['returnValue']['value'].get('totalSize', 0)

                if total_size:
                    # Print the total count of test cases
                    logger.info(f"Total test cases: {total_size}")
                else:
                    logger.error("Error: 'totalSize' not found in the response.")
            else:
                logger.error("Error: Unexpected response structure.")
        else:
            logger.error("Error: Missing 'soapenv:Body' in the response.")

    except (KeyError, json.JSONDecodeError) as e:
        logger.error(f"Error parsing response: {e}")
else:
    # Log the full error message from the response for debugging
    logger.error(f"Error: Received status code {response.status_code} - {response.text}")
