import requests

KEY = "pop"
URL = "pop"
ENDPOINT="/webservice/rest/server.php"

def call(funzione,**args):
    standard={"wstoken": KEY, 'moodlewsrestformat': 'json', "wsfunction": funzione}
    if args:
        for k in args:
            standard[k]=args[k]
    request = requests.get(URL+ENDPOINT, params=standard)
    
    return request.json()