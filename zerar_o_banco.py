import requests

response = requests.post('http://localhost:5000/reset_database')
print(response.json())