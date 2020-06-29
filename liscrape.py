import os, json, csv, time, logging, traceback
import PySimpleGUI as sg
from linkedin_api import Linkedin

class History:
	def __init__(self):
		self.call_count = 0
		self.hourly_limit = 60
		self.history = {}

	def load(self):
		if not os.path.isfile('config.json'):
			with open('config.json', 'w') as config_file:
				config = {'users': {}, 'history': {}}
				json.dump(config, config_file, indent=4)
			
			return {}

		with open('config.json', 'r') as config_file:
			try:
				config = json.load(config_file)
				self.call_count = len(config['history'].keys())
				return config['history']
			except Exception as error:
				logging.exception(error)
				os.remove('config.json')
				return {}

	def store(self):
		if os.path.isfile('config.json'):
			with open('config.json', 'r') as config_file:
				config = json.load(config_file)

			with open('config.json', 'w') as config_file:
				config['history'] = self.history
				json.dump(config, config_file, indent=4)
		else:
			with open('config.json', 'w') as config_file:
				config = {'users': {}, 'history': {}}
				config['history'] = self.history
				json.dump(config, config_file, indent=4)

	def add(self, user_id):
		self.call_count += 1
		not_added = False if user_id in self.history.values() else True

		self.history[time.time()] = user_id
		return not_added

	def check_validity(self):
		if self.call_count > self.hourly_limit:
			filtered_history = {}
			for key, val in self.history.items():
				if time.time() - float(key) <= 3600:
					filtered_history[key] = val

			self.call_count = len(filtered_history.keys())
			if self.call_count > self.hourly_limit:
				earliest_call = min([float(key) for key in self.history.keys()])
				since_call = time.time() - earliest_call
				return False, f'{int((3600 - since_call) / 60)} minutes'

		return True, None

class Session:
	'''
	Session stores the current session's log-in cookie, among other things
	'''
	def __init__(self):
		self.version = '1.0.1'
		self.username = None
		self.password = None
		self.authenticated = False
		self.window = self.display_signin_screen()
		self.sheet_path = None
		self.total_parsed = 0
		self.parsed = 0

		self.history = History()
		self.history.history = self.history.load()
		self.history.check_validity()

	def load_sheet_length(self):
		if not os.path.isfile(self.sheet_path):
			self.total_parsed = 0
		else:
			with open(self.sheet_path, 'r') as csv_file:
				csv_reader =  csv.reader(csv_file)
				self.total_parsed = sum(1 for row in csv_reader)

	def load_configuration(self):
		if not os.path.isfile('config.json'):
			return []

		with open('config.json', 'r') as config_file:
			try:
				config = json.load(config_file)
			except Exception as error:
				logging.exception(error)
				os.remove('config.json')
				return ()

			return tuple(config['users'].keys()) if len(config['users'].keys()) > 0 else ()

	def load_password_from_config(self, username):
		with open('config.json', 'r') as config_file:
			config = json.load(config_file)

		try:
			return config['users'][username]
		except:
			sg.popup('Error finding password from configuration!', title='Error', keep_on_top=True)
			raise Exception('Error finding password from configuration!')

	def store_login(self, username, password):
		if not os.path.isfile('config.json'):
			with open('config.json', 'w') as config_file:
				config = {'users': {}, 'history': {}}
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
		if self.authenticated and remember_login:
			success = self.store_login(username, password)
			if success:
				print('Login stored into config file successfully!')

		return auth_success

	def authenticate(self, refresh_cookies):
		try:
			self.application = Linkedin(self.username, self.password, debug=True, refresh_cookies=refresh_cookies)
			self.authenticated = True
			return True
		except Exception as error:
			logging.exception(error)
			if 'BAD_EMAIL' in error.args:
				sg.popup('Incorrect email: try again.', title='Incorrect email', keep_on_top=True)
			elif 'CHALLENGE' in error.args:
				sg.popup('Error: LinkedIn requires a sign-in challenge.', title='Linkedin error', keep_on_top=True)
			else:
				sg.popup(f'Unhandled exception: {traceback.format_exc(error)}', title='Unhandled exception', keep_on_top=True)

			return False

	def store_profile(self, profile, contact_info):
		def set_diff(dict, full_set):
			ignored_keys = {key for key in full_set if key not in dict.keys()}
			return full_set.difference(ignored_keys)

		profile_keys = {
			'firstName', 'lastName', 'profile_id', 'headline', 
			'summary', 'industryName', 'geoCountryName', 'languages'}
		contact_keys = {'birthdate', 'email_address', 'phone_numbers'}

		profile_keys = set_diff(profile, profile_keys)
		contact_keys = set_diff(contact_info, contact_keys)

		column_map = {
			'firstName': 'First name',
			'lastName': 'Last name',
			'profile_id': 'Linkedin profile ID',
			'headline': 'Linkedin headline',
			'summary': 'Linkedin summary',
			'industryName': 'Industry',
			'geoCountryName': 'Location',
			'languages': 'Languages',
			'birthdate': 'Birthday',
			'email_address': 'Email address',
			'phone_numbers': 'Phone number'
		}

		profile_dict = {}

		for key in profile_keys:
			if key == 'languages':
				profile[key] = ','.join(profile['languages'])

			profile_dict[column_map[key]] = profile[key]

		for key in contact_keys:
			if key == 'phone_numbers':
				profile[key] = ','.join(contact_info['phone_numbers'])
			
			profile_dict[column_map[key]] = contact_info[key]

		field_names = [key for key in profile_dict.keys()]
		if not os.path.isfile(self.sheet_path):
			with open(self.sheet_path, 'w', newline='') as csv_file:
				csv.DictWriter(csv_file, fieldnames=field_names).writeheader()
				print(f'Created file: {self.sheet_path}')

		if not self.history.add(profile_dict['Linkedin profile ID']):
			sg.popup(f'This profile has already been added: avoiding duplicate.', font=('Helvetica', 14), title='Duplicate', keep_on_top=True)
			print(f'Duplicate detected ({profile_dict["Linkedin profile ID"]})\n')
			return

		with open(self.sheet_path, 'a', newline='') as csv_file:
			csv.DictWriter(csv_file, fieldnames=field_names).writerow(profile_dict)
			print(f'Stored {profile_dict["Linkedin profile ID"]}\n')
			self.parsed += 1
			self.total_parsed += 1

	def display_signin_screen(self):
		layout = [
			[sg.Text('Sign in to LinkedIn to continue', font=('Helvetica Bold', 14))],
			[sg.Text('Username (email)\t', font=('Helvetica', 14)), sg.InputText(key="username")],
			[sg.Text('Password\t\t', font=('Helvetica', 14)), sg.InputText(key="password")],
			[	
				sg.Text('Select a stored login', font=('Helvetica', 14)),
				sg.Listbox(
					self.load_configuration(), select_mode='LISTBOX_SELECT_MODE_SINGLE', 
					enable_events=True, size=(40, 1 + len(self.load_configuration())),
					key='-USERNAME-', no_scrollbar=True
					)
			],
			[
				sg.Button('Sign in', font=('Helvetica', 14)), 
				sg.Checkbox('Remember me', key='remember'),
				sg.Checkbox('Refresh cookies', key='cookies'),
				sg.Checkbox('Debug mode', key='debug_mode')
			],
				[sg.Output(size=(80, 20), font=('Helvetica', 12))]
		]

		return sg.Window(f'Liscrape version {self.version}', layout=layout, resizable=True, grab_anywhere=True)

	def display_sheet_screen(self):
		layout = [
			[sg.Text('Choose the Excel sheet to store contact information in (.CSV)', font=('Helvetica Bold', 14))],
			[sg.FileBrowse(), sg.Input(key="sheet_path")],
			[sg.Button('OK'), sg.Button('Use default')]
		]

		return sg.Window(f'Liscrape version {self.version}', layout)

	def display_main_screen(self):
		layout = [
			[sg.Text(f'Signed in as:', font=('Helvetica Bold', 14)), sg.Text(f'{self.username}', font=('Helvetica', 14))],
			[sg.Text('Contact to store (URL)\t', font=('Helvetica', 14)), sg.InputText(key="profile_url")],
			[sg.Button('Store contact', font=('Helvetica', 14)), sg.Text(f'{self.parsed} contacts stored (this session)', key='parsed', font=('Helvetica', 14))],
			[sg.Text(f'Session path: {self.sheet_path}', font=('Helvetica', 14))],
			[sg.Output(size=(60, 20))],
			[sg.Text(f'Contacts in file: {self.total_parsed}', font=('Helvetica', 12), key='total_parsed')]
		]
		return sg.Window(
			title=f'Liscrape version {self.version}', layout=layout, resizable=True, grab_anywhere=True)


if __name__ == '__main__':
	# logging and debug
	log = 'log.log'
	logging.basicConfig(filename=log,level=logging.DEBUG,format='%(asctime)s %(message)s', datefmt='%d/%m/%Y %H:%M:%S')
	logging.info(f'Program started')
	debug = False
	
	# set theme
	sg.theme('Reddit')

	# create session, display sign-in screen
	session = Session()

	# sign-in eventloop
	while True:
		event, values = session.window.read()

		if event == sg.WIN_CLOSED:
			logging.info(f'Sign-in window closed')
			break

		if values['debug_mode']:
			debug = True

		if event == 'Sign in':
			if values['username'] != '' and values['password'] != '' or values['-USERNAME-'] != [] or debug:
				if values['-USERNAME-'] != []:
					username = values['-USERNAME-'][0]
					password = session.load_password_from_config(username)
					values['remember'] = False

				print('Signing in...')
				if not debug:
					auth_success = session.sign_in(values['username'], values['password'], values['remember'], values['cookies'])
				else:
					logging.info(f'Authenticated with debug mode enabled')
					session.authenticated = True
					session.username = 'debug user'
					auth_success = True

				if not auth_success:
					sg.popup('Incorrect login details!', title='Incorrect login', keep_on_top=True)
				else: 
					sg.popup('Signed in successfully!', title='Success', keep_on_top=True)
					session.window.close()

					# request sheet/csv location
					session.window = session.display_sheet_screen()
					while session.sheet_path == None:
						event, values = session.window.read()
						if event == sg.WIN_CLOSED:
							if session.sheet_path == None:
								sg.popup('No Excel sheet defined. Using default path: linkedin_scrape.csv', title='No path defined', keep_on_top=True)
								session.sheet_path = 'linkedin_scrape.csv'
							break

						if event == 'Use default':
							session.sheet_path = 'linkedin_scrape.csv'
							break

						if values['sheet_path'] != '':
							if '.csv' in values['sheet_path'] or '.csv' in values['sheet_path']:
								session.sheet_path = values['sheet_path']

					session.load_sheet_length()
					session.window.close()
					session.window = session.display_main_screen()
					break
			else:
				sg.popup('Please enter your login details!', title='Incorrect login', keep_on_top=True)

	# main eventloop
	try:
		while True and session.authenticated:
			event, values = session.window.read()
			if event == sg.WIN_CLOSED:
				logging.info(f'Main window closed')
				session.history.store()
				session.window.close()
				logging.info(f'Exiting main event loop gracefully')
				break

			if event == 'Store contact' and 'profile_url' in values:
				print(f'Loading {values["profile_url"]}...')
				if values['profile_url'][-1] == '/':
					values['profile_url'] = values['profile_url'][0:-1]

				profile = values['profile_url'].split('/')[-1]
				logging.info(f'Parsing profile {profile}')

				validity_status, time_until_next = session.history.check_validity()
				if validity_status:
					if not debug:
						profile = session.application.get_profile(profile)
						contact_info = session.application.get_profile_contact_info(profile)
						session.store_profile(profile, contact_info)
					else:
						profile = {'lastName': 'Meikäläinen', 'firstName': 'Matti', 'industryName': 'Professional retard', }
						contact_info = {'email_address': 'matti@meikalainen.com', 'websites': ['matti.fi'], 'twitter': '@matti', 'phone_numbers': ['+358400123456']}
						session.store_profile(profile, contact_info)

					# clear input
					session.window['profile_url'].update('')
					session.window['parsed'].update(f'{session.parsed} {"contact" if session.parsed == 1 else "contacts"} stored (this session)')
					session.window['total_parsed'].update(f'Contacts in file: {session.total_parsed}')

					print(f'Call count: {session.history.call_count}')
				else:
					sg.popup(f'API call limit reached. Try again in {time_until_next}.', font=('Helvetica', 14), title='Limit reached', keep_on_top=True)
					logging.info(f'API call limit reached: time until next call {time_until_next}. Limit: {session.history.hourly_limit} calls per hour.')

	except Exception as error:
		logging.exception(error)
		session.history.store()
		session.window.close()
