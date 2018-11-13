# -*- coding: utf-8 -*-
"""
Created on Thu Oct 18 18:19:05 2018
@author: fkong
"""
import onedrivesdk
from onedrivesdk.helpers import GetAuthCodeServer
import datetime
from datetime import datetime as dt

def navigate(client, item_id):
    items = client.item(id=item_id).children.get()
    return items

# authorization part
def auth(redirect_uri, client_secret):
    scopes=['wl.signin', 'wl.offline_access', 'onedrive.readwrite']

    client = onedrivesdk.get_default_client(

        client_id='7cb26e08-62f1-44b9-8151-3e16e87611e7', scopes=scopes)

    auth_url = client.auth_provider.get_auth_url(redirect_uri)

    #this will block until we have the code
    code = GetAuthCodeServer.get_auth_code(auth_url, redirect_uri)
    client.auth_provider.authenticate(code, redirect_uri, client_secret)

    return client



# calculate part

def count(client):

    item_id = "root"

    #   navigate to metrix export folder
    items = navigate(client, item_id)
    for item in items:
        if item.name == "Metric Export":
            item_id = item.id
            break
  

    items = navigate(client, item_id) 
    count_dict = {}

    for item in items:
        child = navigate(client, item.id)
        count_dict[item.name] = 0

        if item.name in ["Failed by Wharton QA", "Ready For Wharton QA", "ReadyForQA2"]:
            for child_item in child:
                count_dict[item.name] += 1
        else: # ready for ballmer group case

            # check it is a subfolder, if no, .folder == None
            for child_item in child:
                if child_item.folder == None:
                    count_dict[item.name] += 1
                else:
                    # go to subfolder
                    subchild = navigate(client, child_item.id)
                    for subitem in subchild:
                        count_dict[item.name] += 1
    return count_dict



def count_total(count_dict):
    total = 0 
    for key in count_dict:
        total += count_dict[key]
    return total



# expected count part

def count_expected():
    now = dt.today()
    nearest_friday = now + datetime.timedelta( (4-now.weekday()) % 7 )
    nearest_friday = nearest_friday.strftime("%A, %d %B")
    #now = dt(2018, 12, 13)
    start_date = dt(2018, 9, 7)
    thanksgiving = dt(2018, 11, 23)
    if (thanksgiving - now).days < 7 and (thanksgiving - now).days >= 0:
        expected_count = 390
    elif (now - thanksgiving).days > 0:
        expected_count = 390 + int(((now - thanksgiving).days - 1) / 7 + 1) * 35
        expected_count = 527 if expected_count > 527 else expected_count
    else:
        expected_count = 40 + int(((now - start_date).days - 1) / 7 + 1) * 35
    return nearest_friday, expected_count


def count_gap(expected_count, total):
    return expected_count - total

redirect_uri = "http://localhost:8082/sign-microsoft"
client_secret = 'kVIV6517_!@mrnfzpFQWY3?'
client = auth(redirect_uri, client_secret)
count_dict = count(client)
total = count_total(count_dict)
nearest_friday, expected_count = count_expected()
gap = count_gap(expected_count, total)

#print(count_dict, total, nearest_friday, expected_count, gap)

from flask import Flask, render_template

app = Flask(__name__, template_folder='templates')

@app.route('/')
@app.route('/index')
def index():

    #return "hello world"
    return render_template('index.html', FailedByWhartonQA = count_dict["Failed by Wharton QA"], \
                            ReadyForWhartonQA = count_dict["Ready For Wharton QA"], \
                            ReadyForBallmerGroup = count_dict["Ready for Ballmer Group"], \
                            ReadyForQA2 = count_dict["ReadyForQA2"], \
                            total = total, \
                            nearestFriday = nearest_friday,
                            expectedCount = expected_count, \
                            gap = gap)



if __name__ == "__main__":
    app.run(debug=True)