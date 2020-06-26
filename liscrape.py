import requests
import PySimpleGUI as sg

from pprint import pprint
from linkedin_api import Linkedin

'''
Basic idea:
	1. find an interesting profile
	2. drag link into UI
	3. info automatically pulled from Linkedin and inserted into a CSV for automatic import into CRM
'''
class Session:
	'''
	Session stores the current session's log-in cookie, among other things
	'''
	def __init__(self):
		self.username = None
		self.password = None
		self.authenticated = False

	def sign_in(self, username, password):
		self.username = username
		self.password = password
		return self.authenticate()

	def authenticate(self):
		try:
			api = Linkedin(self.username, self.password)
			self.authenticated = True
			print(f'api: {api}')
			return api
		except Exception as error:
			error_args = error.args
			if 'BAD_EMAIL' in error_args:
				sg.popup('Incorrect email: try again.')
			elif 'CHALLENGE' in error_args:
				sg.popup('Challenge')
			else:
				sg.popup(error_args)

			return False

	def display_signin_screen(self, VERSION):
		layout = [
			[sg.Text('Sign in to LinkedIn to continue')],
			[sg.Text('Username (email)\t'), sg.InputText(key="username")],
			[sg.Text('Password\t\t'), sg.InputText(key="password")],
			[sg.Button('Sign in')],
			[sg.Output(size=(60, 20))]
		]

		return sg.Window(f'Liscrape version {VERSION}', layout)


if __name__ == '__main__':
	VERSION = 0.1
	headers = {
		"user-agent": 
			"Mozilla/5.0 \
			(Macintosh; Intel Mac OS X 10_13_5) \
			AppleWebKit/537.36 (KHTML, like Gecko) \
			Chrome/66.0.3359.181 \
			Safari/537.36"
	}
	
	# set theme
	sg.theme('Reddit')

	# create session, display sign-in screen
	session = Session()
	window = session.display_signin_screen(VERSION)

	# Event Loop to process "events" and get the "values" of the inputs
	while True:
		event, values = window.read()
		print(f'event: {event} | values: {values}')

		if event == sg.WIN_CLOSED:
			break

		elif event == 'Sign in':
			if values['username'] != '' and values['password'] != '':
				api = session.sign_in(values['username'], values['password'])
				print('Signing in...')

				if api is False:
					sg.popup('Incorrect login details!')

			else:
				sg.popup('Please enter your login details!')

		print('You entered ', values)


	window.close()



