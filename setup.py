import os

#########################
#			#
#EASY SETUP FILE 	#
#AUTOMATICALLY DELETES	#
#AFTER SUCCESSFUL	#
#INSTALL		#
#########################
for x in range(0,2):
	
	try:
		os.system('sudo pip install pyinstaller')
		os.system('pyinstaller --onefile cl.py')
	except:
		os.system('sudo easy_install pip')
