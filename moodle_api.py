import requests
from requests import get, post
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

KEY = "pop"
URL = "pop"
ENDPOINT="/webservice/rest/server.php"
MAX_RETRIES=10

def rest_api_parameters(in_args, prefix='', out_dict=None):
    if out_dict==None:
        out_dict = {}
    if not type(in_args) in (list,dict):
        out_dict[prefix] = in_args
        return out_dict
    if prefix == '':
        prefix = prefix + '{0}'
    else:
        prefix = prefix + '[{0}]'
    if type(in_args)==list:
        for idx, item in enumerate(in_args):
            rest_api_parameters(item, prefix.format(idx), out_dict)
    elif type(in_args)==dict:
        for key, item in in_args.items():
            rest_api_parameters(item, prefix.format(key), out_dict)
    return out_dict

def call(fname, **kwargs):
    session = requests.Session()
    retry = Retry(connect=MAX_RETRIES, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    parameters = rest_api_parameters(kwargs)
    parameters.update({"wstoken": KEY, 'moodlewsrestformat': 'json', "wsfunction": fname})
    if MAX_RETRIES > 0:
        response = session.post(URL+ENDPOINT, parameters,adapter)
    else:
        response = post(URL+ENDPOINT, parameters)
    response = response.json()
    if type(response) == dict and response.get('exception'):
        raise SystemError("Error calling Moodle API\n", response)
    return response
