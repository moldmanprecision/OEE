import requests

url = "https://oee-1-4n1e.onrender.com/upload-json"

files = {
    "file": open("output/oee_data.json", "rb")
}

response = requests.post(url, files=files)

print(response.text)