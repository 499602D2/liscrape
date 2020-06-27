import os, json
import PySimpleGUI as sg

from linkedin_api import Linkedin

'''
Basic idea:
	1. find an interesting profile
	2. drag link into UI
	3. info automatically pulled from Linkedin and inserted into a CSV for automatic import into CRM

To-do
	Package with PyInstaller
'''
class Session:
	'''
	Session stores the current session's log-in cookie, among other things
	'''
	def __init__(self):
		self.username = None
		self.password = None
		self.authenticated = False
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
				return ()

			return tuple(config['users'].keys()) if len(config['users'].keys()) > 0 else ()

	def load_password_from_config(self, username):
		with open('config.json', 'r') as config_file:
			config = json.load(config_file)

		try:
			return config['users'][username]
		except:
			sg.popup('Error finding password from configuration!')
			raise Exception('Error finding password from configuration!')

	def store_login(self, username, password):
		if not os.path.isfile('config.json'):
			with open('config.json', 'w') as config_file:
				config = {}
				config['users'] = {}
		else:
			with open('config.json', 'r') as config_file:
				config = json.load(config_file)

		config['users'][username] = password
		with open('config.json', 'w') as config_file:
			json.dump(config, config_file, indent=4)

		return True

	def sign_in(self, username, password, remember_login, refresh_cookies):
		self.username = username
		self.password = password
		auth_success = self.authenticate(refresh_cookies)

		# DEBUG DEBUG DEBUG DEBUG DEBUG DEBUG 
		self.authenticated = True
		print(f'remember_login: {remember_login}, type: {type(remember_login)}')
		if self.authenticated and remember_login:
			success = self.store_login(username, password)
			if success:
				print('Login stored into config file successfully!')

		return auth_success

	def authenticate(self, refresh_cookies):
		try:
			api = Linkedin(self.username, self.password, debug=True, refresh_cookies=refresh_cookies)
			self.authenticated = True
			print(f'api: {api}')
			return api
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
			[sg.Text('Sign in to LinkedIn to continue', font=('Helvetica Bold', 14))],
			[sg.Text('Username (email)\t', font=('Helvetica', 14)), sg.InputText(key="username")],
			[sg.Text('Password\t\t', font=('Helvetica', 14)), sg.InputText(key="password")],
			[	
				sg.Text('Select a stored login', font=('Helvetica', 14)),
				sg.Listbox(
					self.load_configuration(), select_mode='LISTBOX_SELECT_MODE_SINGLE', 
					enable_events=True, size=(40, 1 + len(self.load_configuration())),
					key='-USERNAME-'
					)
			],
			[
				sg.Button('Sign in', font=('Helvetica', 14)), 
				sg.Checkbox('Remember me', key='remember'),
				sg.Checkbox('Refresh cookies', key='cookies')]
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
			[sg.Text(f'Signed in as {self.username}', font=('Helvetica', 14))],
			[sg.Text('Contact to store (URL)\t', font=('Helvetica', 14)), sg.InputText(key="profile_url")],
			[sg.Button('Store contact', font=('Helvetica', 14)), sg.Text(f'{self.parsed} contacts stored')],
			[sg.Text(f'Storing into {self.sheet_path}', font=('Helvetica', 12))]
		] # [sg.Output(size=(60, 20))]

		return sg.Window(
			title=f'Liscrape version {VERSION}', layout=layout, resizable=True, grab_anywhere=True)


if __name__ == '__main__':
	VERSION = 0.2
	
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
			if values['username'] != '' and values['password'] != '' or values['-USERNAME-'] != []:
				if values['-USERNAME-'] != []:
					username = values['-USERNAME-'][0]
					password = session.load_password_from_config(username)
					values['remember'] = False

				api = session.sign_in(values['username'], values['password'], values['remember'], values['cookies'])
				print('Signing in...')

				if api is False:
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

	if session.sheet_path == None:
		session.window.close()
		sg.popup('Error: no Excel sheet path defined!')

	# main eventloop
	while True:
		event, values = session.window.read()
		if event == sg.WIN_CLOSED:
			break

		if 'profile_url' in values:
			print(f'Loading {values["profile_url"]}...')
			profile = values['profile_url'].split('/')[-1]

			profile = api.get_profile(profile)
			print(profile)
			print('\n\n')

			contact_info = api.get_profile_contact_info(profile)
			print(contact_info)


	session.window.close()



