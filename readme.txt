Highlights documentation

- navigate to the C drive and create a directory there called 'Scripts'. Save your python scripts in that location
- open up the command line. You can do this by typing 'cmd' in the Windows search bar
- the command line will open and display with a default directory location. You will need to change this directory to the one where your scripts reside. To do this, type 'cd c:\Scripts'. NB: 'cd' stands for change directory. 
- to run the scripts you will need to type their filename then a space, then give them one of two variables, 'weekly' or 'monthly', and press enter. This tells the script whether you want to collect data from the Aicer report for the last week, or for the last month. For example:
	- 'newcontent.py weekly' will run a report to show you all the new content that's been published in the last week
	- 'newcontent.py monthly' will run a report to show you all the new content that's been published in the last month
	- 'updatedcontent.py weekly' will run a report to show you all the content that's been updated in the last week
	- 'updatedcontent.py monthly' will run a report to show you all the content that's been updated in the last month
- if you receive a ModuleNotFound error, you will need to type 'pip install <name of the module>'. This will install the missing python module. You will only need to do this once for every missing module.
- be patient when running these scripts as they can take up to 20 minutes each time as they load the most recent Aicer report which is over 130Mb.
- the scripts will output the progress on the command line and will tell you when it is finished. The reports will be output to the following location: \\atlas\lexispsl\Highlights\Automatic creation\New and Updated content report
	
For any other issues contact Daniel Hutchings: daniel.hutchings.1@lexisnexis.co.uk