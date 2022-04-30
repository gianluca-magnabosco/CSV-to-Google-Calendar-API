# CSV and Excel formatted file importer to Google Calendar, via API
This script will create a calendar, and insert all of the events present on the [excel/csv formatted file](https://github.com/gianluca-magnabosco/CSV-to-Google-Calendar-API/blob/main/csv_file.csv), taking into consideration all possible outcomes from its parameters, such as `Subject`, `Start Date`, `All Day Event`, `Location`, `Private`, etc. </br>

`There is an example of the formatted file just down below. ↓` </br>

If you have been using Google Calendar and came across the `Import CSV` option, but couldn't automate the insertion of events through a `.csv` file using the Google Calendar API, this [script](https://github.com/gianluca-magnabosco/CSV-to-Google-Calendar-API/blob/main/CSV-to-Google-Calendar-API.py) might help you!
> If you have an excel file that is formatted in the same way, don't worry, the script will work!

##
</br>

In order for the events to be successfully inserted, you'll need to get your `client_secret.json` file from the Google Cloud Platform, if you scroll down a little you can find a  guide on how to get your secret file ↓↓↓.

##


### On this repository there are two sample files, they contain the same data, but are in two different formats:

![Sample File](https://media.discordapp.net/attachments/555940526554218496/935920316654420018/aasdadas.png?width=819&height=559)
> This is how the `.csv` formatted file from the `Import CSV` Google Calendar option looks like.



##
</br >

## Instructions - Getting your client_secret.json file!
### Before you run the script, you'll need to get your `client_secret.json` file, you can do that by signing into the Google Cloud Platform:
### <https://console.cloud.google.com/>

### After logging in or signing into your Google account, you'll need to create a project on the platform:
![calendar1](https://media.discordapp.net/attachments/810687915045814293/935486027412406302/eproc_1.png)

</br >

### Activate the Google Calendar API:
<div>
  
  ![calendar2](https://media.discordapp.net/attachments/810687915045814293/935486027655696384/eproc_2.png)
  
  ![calendar3](https://media.discordapp.net/attachments/810687915045814293/935486027953496105/eproc_3.png)
  
  ![calendar4](https://media.discordapp.net/attachments/810687915045814293/935486029136293908/eproc_4.png)
  
  ![calendar5](https://media.discordapp.net/attachments/810687915045814293/935486028230295562/eproc_5.png)
</div>

</br >

### Configure consent screen:
<div>
  
  ![calendar6](https://media.discordapp.net/attachments/810687915045814293/935486028536500224/eproc_6.png)

  ![calendar7](https://media.discordapp.net/attachments/810687915045814293/935486028809117716/eproc_7.png)

</div>

</br >

### Create an OAuth authentication key, and download its .json file:
<div>
  
  ![calendar8](https://media.discordapp.net/attachments/810687915045814293/935486025202040902/eproc_8.png)

  ![calendar9](https://media.discordapp.net/attachments/810687915045814293/935486025470451762/eproc_9.png)

  ![calendar10](https://media.discordapp.net/attachments/810687915045814293/935486048430067772/eproc_10.png?width=984&height=559)
  
  ![calendar11](https://media.discordapp.net/attachments/810687915045814293/935486048929185842/eproc_11.png?width=973&height=559)
</div>


> Rename the OAuth authentication file to  _`client_secret.json`_ and place it in your current directory.
## 
</br >
