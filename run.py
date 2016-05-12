import os

#########################
#LAZY			#
#EASY SETUP FILE 	#
#AUTOMATICALLY DELETES	#
#AFTER SUCCESSFUL	#
#INSTALL		#
#########################


os.system('/usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"')
os.system('brew install python')
os.system('sudo pip install pyinstaller')
os.system('pyinstaller --onefile cl.py')
