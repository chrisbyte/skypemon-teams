import json
import os
import requests
import time
import webbrowser

#
# References
#
# https://portal.azure.com/#home
# https://developer.microsoft.com/en-us/graph/graph-explorer
# https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0
# https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code

class MicroAuth:

  # settings
  __auth_url = None
  __client_id = None
  __graph_url = None
  __scopes = None
  __tenant = None
  __token_file = None

  def __init__(self, tenant, client_id):
    self.__client_id = client_id
    self.__tenant = tenant
    self.__auth_url = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/'.format(tenant=tenant)
    self.__graph_url = 'https://graph.microsoft.com/'
    self.__scopes = 'offline_access presence.read presence.read.all'
    self.__token_file = 'token.txt'

  def __create_access_token(self):
    print('request permissions')

    # request user permissions
    validation_codes = self.__request_permissions()
    print('user code: ', validation_codes['user_code'])

    url = self.__auth_url + 'token'
    data = { 
      'tenant':self.__tenant, 
      'client_id':self.__client_id, 
      'grant_type':'device_code', 
      'device_code': validation_codes['device_code'] 
    }
    access_token = None

    while access_token is None:
      print('checking permissions')
      time.sleep(validation_codes['interval'])

      # request
      json = requests.post(url, data).json()

      if 'access_token' in json:
        print('permissions granted')
        access_token = { 
          'access_token': str(json['access_token']), 
          'expires_in': str(json['expires_in']),
          'refresh_token': str(json['refresh_token'])
        }

    # save
    self.__save_access_token(access_token)

    return access_token

  def __graph_host(self, access_token=None, id=None, endpoint='', version='v1.0'):
    if access_token is None:
      access_token = self.load_access_token()['access_token']

    if id == None:
      endpoint = 'me/' + endpoint
    else:
      endpoint = 'users/' + id + '/' + endpoint

    return {
      'url': self.__graph_url + version + '/' + endpoint,
      'headers': {"Authorization": "Bearer " + access_token}
    }

  def __read_access_token(self):
    if os.path.isfile(self.__token_file):
      with open(self.__token_file, "r") as out_file:
        return { 
          'expires_in': int(out_file.readline()),
          'access_token': out_file.readline(), 
          'refresh_token': out_file.readline()
        }

  def __refresh_access_token(self, refresh_token):
    print('refresh token')

    url = self.__auth_url + 'token'
    data = { 
      'tenant':self.__tenant, 
      'client_id':self.__client_id, 
      'refresh_token': refresh_token,
      'grant_type':'refresh_token', 
      'scope': self.__scopes
    }

    # request
    json = requests.post(url, data).json()

    access_token = { 
      'access_token': str(json['access_token']), 
      'expires_in': str(json['expires_in']),
      'refresh_token': str(json['refresh_token'])
    }

    # save
    self.__save_access_token(access_token)

    return access_token

  def __request_permissions(self):
    url = self.__auth_url + 'devicecode'
    data = { 
      'tenant': self.__tenant, 
      'client_id': self.__client_id, 
      'scope': self.__scopes
    }

    # request
    json = requests.post(url, data).json()

    # open browser to verify
    webbrowser.open(json['verification_uri'], new=2)

    return { 
      'user_code': json['user_code'], 
      'device_code': json['device_code'],
      'interval': json['interval']
    }
  
  def __save_access_token(self, data):
    with open(self.__token_file, "wb") as out_file:
      out_file.writelines([
        data['expires_in']
        , '\n', data['access_token']
        , '\n', data['refresh_token']
      ])

  def get_graph_mailbox(self, access_token=None, id=None):
    host = self.__graph_host(access_token, id, 'mailFolders/inbox')
    return requests.get(host['url'], headers=host['headers']).json()

  def get_graph_manager(self, access_token=None, id=None):
    host = self.__graph_host(access_token, id, 'manager')
    return requests.get(host['url'], headers=host['headers']).json()

  def get_graph_presence(self, access_token=None, id=None):
    host = self.__graph_host(access_token, id, 'presence', 'beta')
    return requests.get(host['url'], headers=host['headers']).json()

  def get_graph_user(self, access_token=None, id=None):
    host = self.__graph_host(access_token, id)
    return requests.get(host['url'], headers=host['headers']).json()

  def load_access_token(self):

    # check for a stored access token
    access_token = self.__read_access_token()

    # if no token
    if access_token is None or 'access_token' not in access_token:

      # grant permissions and create token
      access_token = self.__create_access_token()

    else:

      expires = time.time() - os.path.getmtime(self.__token_file)

      # if access token expired
      if expires >= access_token['expires_in']:

        # refresh access token
        access_token = self.__refresh_access_token(access_token['refresh_token'])
        
      else:

        # display time until access token expires
        min = int((access_token['expires_in'] - expires) / 60)
        sec = int((access_token['expires_in'] - expires) % 60)
        print('expires in: {min} min {sec} sec'.format(min=min, sec=sec))

    return access_token

def print_mailbox(access_token):
  mailbox = auth.get_graph_mailbox(access_token)
  print('unread emails: ' + str(mailbox['unreadItemCount']))

def print_manager(access_token):
  manager = auth.get_graph_manager(access_token)
  msg = manager['mail'] + ': ' + manager['id']
  print('manager: ' + msg)

  # presence
  print_presence(access_token, manager['id'], 'manager')

def print_presence(access_token, id=None, label='user'):
  presence = auth.get_graph_presence(access_token, id)
  msg = presence['availability']
  if presence['activity'] != msg:
    msg = msg + ': ' + presence['activity']
  print(label + ' presence: ' + msg)

def print_user(access_token):
  user = auth.get_graph_user(access_token)
  msg = user['mail'] + ': ' + user['id']
  print('user: ' + msg)

#
# Run the app
#
if __name__ == "__main__":

  if os.path.isfile('settings.json'):
    with open('settings.json') as json_file:
      CONFIG = json.load(json_file)

      # create auth object and get access token
      auth = MicroAuth(CONFIG['tenant'], CONFIG['client_id'])
      access_token = auth.load_access_token()['access_token']

      # print tests
      # print_user(access_token)
      print_presence(access_token)
      print_mailbox(access_token)
      # print_manager(access_token)

  else:
    print('no settings file found')