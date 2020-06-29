# -*- coding: utf-8 -*-
import os, json
import PySimpleGUI as sg
from linkedin_v2 import linkedin

'''
To-do
	Package with PyInstaller
'''
class Session:
	def __init__(self):
		self.application_name = None
		self.authenticated = False
		self.application = None
		
		self.consumer_key = None
		self.consumer_secret = None
		self.user_token = None
		self.user_secret = None
		
		self.window = None
		self.sheet_path = None
		self.parsed = 0

	def load_configuration(self):
		if not os.path.isfile('config.json'):
			return []

		with open('config.json', 'r') as config_file:
			try:
				config = json.load(config_file)
			except:
				os.remove('config.json')
				return []

			keys = []
			for key in config:
				if config[key]['app_name'] != '':
					keys.append(config[key]['app_name'])
				else:
					keys.append(key)

			return keys if len(keys) > 0 else ()

	def load_password_from_config(self, consumer_key):
		with open('config.json', 'r') as config_file:
			config = json.load(config_file)

		# if key not found, we're using an app name
		if consumer_key not in config.keys():
			for key in config:
				if config[key]['app_name'] == consumer_key:
					return config[key]

		try:
			return config[consumer_key]
		except:
			sg.popup('Error finding password from configuration!')
			raise Exception('Error finding password from configuration!')

	def store_login(self, app_name, key_dict):
		consumer_key = key_dict['consumer_key']
		if not os.path.isfile('config.json'):
			with open('config.json', 'w') as config_file:
				config = {}
				config[consumer_key] = {}
		else:
			with open('config.json', 'r') as config_file:
				config = json.load(config_file)
				if key_dict['consumer_key'] not in config.keys():
					config[consumer_key] = {}

		config[consumer_key]['app_name'] = app_name
		for key, val in key_dict.items():
			config[consumer_key][key] = val

		with open('config.json', 'w') as config_file:
			json.dump(config, config_file, indent=4)

		return True

	def sign_in(self, app_name, key_dict, remember_login):
		if app_name == '':
			app_name = key_dict['consumer_key'][0:6]

		self.consumer_key = key_dict['consumer_key']
		self.consumer_secret = key_dict['consumer_secret']
		self.user_token = key_dict['user_token']
		self.user_secret = key_dict['user_secret']

		auth_success = self.authenticate()

		print(f'remember_login: {remember_login}, type: {type(remember_login)}')
		if self.authenticated and remember_login:
			self.application_name = app_name

			success = self.store_login(app_name, key_dict)
			if success:
				print('Login stored into config file successfully!')

		return auth_success

	def authenticate(self):
		try:
			RETURN_URL = 'http://localhost:8000'
			authentication = linkedin.LinkedInDeveloperAuthentication(
				self.consumer_key, self.consumer_secret, 
				self.user_token, self.user_secret, 
				RETURN_URL, linkedin.PERMISSIONS.enums.values()
			)
				
			sg.popup(f'Authentication: {authentication}')

			self.authenticated = True
			return linkedin.LinkedInApplication(authentication)

		except Exception as error:
			error_args = error.args

			if 'BAD_EMAIL' in error_args:
				sg.popup('Incorrect email: try again.')
			elif 'CHALLENGE' in error_args:
				sg.popup('Error: LinkedIn requires a sign-in challenge.')
			else:
				print(error)
				print(error_args)
				sg.popup(f'Unhandled exception: {error_args}')

			return False

	def store_contact_information(self, contact_dict):
		pass

	def display_signin_screen(self, VERSION):
		layout = [
			[sg.Text('LinkedIn developer authentication', font=('Helvetica Bold', 14))],
			[sg.Text('Application name\t', font=('Helvetica', 14)), sg.InputText(key="name")],
			[sg.Text('Consumer key\t', font=('Helvetica', 14)), sg.InputText(key="consumer_key")],
			[sg.Text('Consumer secret\t', font=('Helvetica', 14)), sg.InputText(key="consumer_secret")],
			[sg.Text('User token\t', font=('Helvetica', 14)), sg.InputText(key="user_token")],
			[sg.Text('User secret\t', font=('Helvetica', 14)), sg.InputText(key="user_secret")],
			[	
				sg.Text('Select a stored login', font=('Helvetica', 14)),
				sg.Listbox(
					self.load_configuration(), select_mode='LISTBOX_SELECT_MODE_SINGLE', size=(40, 2 + len(self.load_configuration())),
					key='-CONSUMER_KEY-', enable_events=True
					)
			],
			[
				sg.Button('Sign in', font=('Helvetica', 14)), 
				sg.Checkbox('Remember me', key='remember')
			]
		] # [sg.Output(size=(80, 20))]

		return sg.Window(f'Liscrape version {VERSION}', layout=layout, resizable=True, grab_anywhere=True)

	def display_sheet_screen(self, VERSION):
		layout = [
			[sg.Text('Choose the Excel sheet to store contact information in', font=('Helvetica Bold', 14))],
			[sg.FileBrowse(), sg.Input(key="sheet_path")],
			[sg.Button('OK')]
		]

		return sg.Window(f'Liscrape version {VERSION}', layout)

	def display_main_screen(self, VERSION):
		layout = [
			[sg.Text(f'Using application {self.application_name}', font=('Helvetica', 14))],
			[sg.Text('Contact to store (URL)\t', font=('Helvetica', 14)), sg.InputText(key="profile_url")],
			[sg.Button('Store contact', font=('Helvetica', 14)), sg.Text(f'{self.parsed} contacts stored')],
			[sg.Text(f'Storing into {self.sheet_path}', font=('Helvetica', 12))],
			[sg.Text(f'(c) 2020 Icotak Ltd.', font=('Helvetica', 10))]
		] # [sg.Output(size=(60, 20))]

		return sg.Window(
			title=f'Liscrape version {VERSION}', layout=layout, resizable=True, grab_anywhere=True)


if __name__ == '__main__':
	VERSION = 0.3
	
	# set theme
	sg.theme('Reddit')

	# create session, display sign-in screen
	session = Session()
	session.window = session.display_signin_screen(VERSION)

	# sign-in eventloop
	while True:
		event, values = session.window.read()
		print(f'event: {event} | values: {values}')

		if event == sg.WIN_CLOSED:
			break

		if event == 'Sign in':
			login_fields = (
				values['consumer_key'], values['consumer_secret'],
				values['user_token'], values['user_secret'])

			print(f'Login fields: {login_fields}')
			print(f'Values: {values}')
			print(f'Values[consumer_key]: {values["-CONSUMER_KEY-"]}')

			if '' not in login_fields or values['-CONSUMER_KEY-'] != []:
				if values['-CONSUMER_KEY-'] != []:
					consumer_key = values['-CONSUMER_KEY-'][0]
					key_dict = session.load_password_from_config(consumer_key)
					values['remember'] = False
				else:
					key_dict = {}
					for key in ('consumer_key', 'consumer_secret', 'user_token', 'user_secret'):
						key_dict[key] = values[key]

				session.application = session.sign_in(values['name'], key_dict, values['remember'])
				print('Signing in...')

				if session.application is None:
					sg.popup('Incorrect login details!')
				else: 
					# successful sign-in, update UI
					sg.popup('Signed in successfully!')
					session.window.close()

					# request sheet/csv location
					session.window = session.display_sheet_screen(VERSION)
					while session.sheet_path == None:
						event, values = session.window.read()
						print(f'event: {event} | values: {values}')
						if event == sg.WIN_CLOSED:
							break

						if values['sheet_path'] != '':
							if '.xls' in values['sheet_path'] or '.xlsx' in values['sheet_path']:
								session.sheet_path = values['sheet_path']

					session.window.close()

					# main screen
					session.window = session.display_main_screen(VERSION)
					break
			else:
				sg.popup('Please enter your login details!')

		print('You entered ', values)

	if session.sheet_path == None and session.authenticated:
		session.window.close()
		sg.popup('Error: no Excel sheet path defined!')

	# main eventloop
	while True:
		event, values = session.window.read()
		if event == sg.WIN_CLOSED:
			break

		if event == 'Store contact' and 'profile_url' in values:
			print(f'Loading {values["profile_url"]}...')
			profile = values['profile_url'].split('/')[-1]

			profile = session.application.get_profile(profile)
			print(profile)
			print('\n\n')

			contact_info = session.application.get_profile_contact_info(profile)
			print(contact_info)


	session.window.close()



