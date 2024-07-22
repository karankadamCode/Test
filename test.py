import json
import pandas as pd 
import requests
import numpy as np

def sanitize_json(data):
    """
    Recursively sanitize JSON data to ensure compliance with JSON format.
    """
    if isinstance(data, dict):
        return {key: sanitize_json(value) for key, value in data.items()}
    elif isinstance(data, list):
        return [sanitize_json(item) for item in data]
    elif isinstance(data, float):
        # Handle special float values
        if np.isinf(data) or np.isnan(data):
            return None  # or some placeholder value
        return data
    else:
        return data

def api_fetch(api_url, description, access_token):
    try:
        # Defining the API endpoint URL
        endpoint_url = f"{api_url}/case-classifier"

        # Preparing the headers and data for the POST request
        headers = {
            "accept": "application/json",
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        data = {"msg": description}  # Ensure key matches the API expectations

        # Sending the POST request with authentication
        response = requests.post(endpoint_url, headers=headers, json=data)

        if response.status_code == 200:
            try:
                json_response = response.json()
                
                # Sanitize JSON response to handle non-compliant values
                sanitized_response = sanitize_json(json_response)
                return sanitized_response
            except json.JSONDecodeError:
                print("Error: Response is not in JSON format.")
                print("Response text:", response.text)
                return None
        else:
            print(f"Error: {response.status_code} - {response.text}")
            return None

    except Exception as e:
        print(f"Error: {str(e)}")
        return None



# def store_in_excel():
#     api_url = "https://caseratingv3-web-app.azurewebsites.net"
#     access_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTcyMTY0NTc4Mn0.qLC853h86k881h1tDnHEN7xUG9fSXIL0xfEDTnj0QvM"

#     # Read the input Excel file
#     input_df = pd.read_excel('SSD Intakes for AI Scrub Sample Data.xlsx')
    
#     for description in input_df['Description']:
#         response = api_fetch(api_url, description, access_token)
#         print(response)

# store_in_excel()

# def store_in_excel():
#     api_url = "https://caseratingv3-web-app.azurewebsites.net"
#     access_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTcyMTY0NTc4Mn0.qLC853h86k881h1tDnHEN7xUG9fSXIL0xfEDTnj0QvM"

#     # Read the input Excel file
#     input_df = pd.read_excel('SSD Intakes for AI Scrub Sample Data.xlsx', nrows=10)
    
#     results = []  # List to hold all responses

#     for description in input_df['Description']:
#         response = api_fetch(api_url, description, access_token)
#         if response:  # Only append if the response is valid
#             results.append(response)
    
#     print(results)
    
#     # Convert results to DataFrame
#     # results_df = pd.json_normalize(results)

    
#     # # Save to Excel
#     # results_df.to_excel('test_results.xlsx', index=False)
#     # print(f"Results have been saved to 'test_results.xlsx'.")

# store_in_excel()



# def store_in_excel():
#     api_url = "https://caseratingv3-web-app.azurewebsites.net"
#     access_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTcyMTY1MTg5N30.gRLLZmfJgOr6PGY4oegcdz7j273YFdPjs6Om46tZKnY"

#     # Read the input Excel file
#     input_df = pd.read_excel('SSD Intakes for AI Scrub Sample Data.xlsx', nrows=10)
    
#     output_file = 'test_results.xlsx'
    
#     # Initialize a list to store all responses
#     results = []
    
#     for i, description in enumerate(input_df['Description']):
#         response = api_fetch(api_url, description, access_token)
    
#         if response:  # Only process if the response is valid
#             # Ensure response is a valid JSON string
#             try:
#                 response_dict = json.loads(response.replace("'", '"'))
#                 response_df = pd.DataFrame([response_dict])
#                 response_df['Description'] = description  # Add description column
#                 results.append(response_df)
#             except json.JSONDecodeError as e:
#                 print(f"Error decoding JSON for record {i + 1}: {e}")
        
#         print(f"Processed record {i + 1}")
    
#     # Combine all DataFrames into a single DataFrame
#     final_df = pd.concat(results, ignore_index=True)
    
#     # Save the combined DataFrame to Excel
#     final_df.to_excel(output_file, index=False)
#     print(f"Results have been saved to '{output_file}'.")

# store_in_excel()

def store_in_excel():
    api_url = "https://caseratingv3-web-app.azurewebsites.net"
    access_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTcyMTY1MTg5N30.gRLLZmfJgOr6PGY4oegcdz7j273YFdPjs6Om46tZKnY"

    # Read the input Excel file
    input_df = pd.read_excel('SSD Intakes for AI Scrub Sample Data.xlsx', nrows=100)
    
    output_file = 'test_results.xlsx'
    
    # Initialize a list to store all responses
    results = []
    
    for i, description in enumerate(input_df['Description']):
        response = api_fetch(api_url, description, access_token)
    
        if response:  # Only process if the response is valid
            # Ensure response is a valid JSON string
            try:
                response_dict = json.loads(response.replace("'", '"'))
                response_df = pd.DataFrame([response_dict])
                response_df['Description'] = description  # Add description column
                results.append(response_df)
            except json.JSONDecodeError as e:
                print(f"Error decoding JSON for record {i + 1}: {e}")
        
        print(f"Processed record {i + 1}")
    
    # Combine all DataFrames into a single DataFrame
    final_df = pd.concat(results, ignore_index=True)
    
    # Reorder columns to ensure 'Description' is the first column
    columns_order = ['Description'] + [col for col in final_df.columns if col != 'Description']
    final_df = final_df[columns_order]
    
    # Save the combined DataFrame to Excel
    final_df.to_excel(output_file, index=False)
    print(f"Results have been saved to '{output_file}'.")

store_in_excel()