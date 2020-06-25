import PySimpleGUI as sg

'''
Basic idea:
	1. find an interesting profile
	2. drag link into UI
	3. info automatically pulled from Linkedin and inserted into a CSV for automatic import into CRM
'''

if __name__ == '__main__':
	VERSION = 0.1

	# All the stuff inside your window.
	layout = [
				[sg.Text('Drag Linkedin profile link into the form beow')],
				[sg.Text('Linkedin profile URL\t\t'), sg.InputText()],
				[sg.Text('File the data is saved to\t'), sg.InputText()],
				[sg.Button('Process'), sg.Button('Cancel')] ]

	# Create the Window
	window = sg.Window(f'Liscrape {VERSION}', layout)

	# Event Loop to process "events" and get the "values" of the inputs
	while True:
	    event, values = window.read()
	    
	    if event == sg.WIN_CLOSED or event == 'Cancel':	# if user closes window or clicks cancel
	        break

	    print('You entered ', values[0])

	window.close()