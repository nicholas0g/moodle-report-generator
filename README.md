# Moodle Global user report generator
Client for xlsx report generator.  
The report will contain all the individual courses by number of participants and a detail of all users and all their enrollments.  
In order to use the client you will need to cofigure the web service of your moodle installation.  

![](https://img.shields.io/github/stars/nicholas0g/moodle-report-generator.svg) ![](https://img.shields.io/github/forks/nicholas0g/moodle-report-generator.svg) ![](https://img.shields.io/github/tag/nicholas0g/moodle-report-generator.svg)  
## Warning
:exclamation: PLEASE BE AWARE: the GUI will save on your local pc an unencrypted moodle token! (for now it is used for development)
## Enable Webservices
Go to `Site administration -> Plugins -> Web services -> Overview` and follow the instructions to enable web services and generate authorization token. 
## Needed functions
In order to use the client you will need to grant access to the functions `core_course_get_courses` and `core_enrol_get_enrolled_users`

# Utilization
The client is released as a GUI or a CLI script.  
The report will be saved in the same folder of execution.  
For the script you must set the global variables  
```
moodle_api.URL = "your_moodle_utl"
moodle_api.KEY = "your_moodle_token"
mail_skip="mail_to_skip"
```
`mail_skip` can be used to filter out users with particular email, like your company email me@mycompany.org  
If you set `mail_skip="@mycompany.org"`, all the user with an email containing @mycompany.org will be ignored.  
Before running make shure to install all the required pip packages in `requirements.txt`.  

To use the cli simply run: 
```
python run.py
```
To use the gui in development mode:
```
python gui.py
```
## DEVELOPMENT WARNING
:exclamation::exclamation::exclamation: The GUI will automatically save the entered data in a json file for future usage.


