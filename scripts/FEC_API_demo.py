import requests

url = "http://api.open.fec.gov/v1/committee/C00401224"

print(requests.get(url, params={"api_key": "VDkmeFlFlO9ZRao7AyDyPMrgEeSdwJXO8UdN7faS"}).json()["results"][0]["party"])