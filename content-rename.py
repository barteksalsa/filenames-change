import os
import re

pptMacroFile = "content-rename.pptm" # znajdz sobie skrypt
pptMacro = "ReplaceSKBR" # taka jest na stale
powerpointExt = "(\.pptx$)|(\.ppt$)"
batFile = "content-rename.bat"
vbsFileName = "content-rename.vbs"

''' create a bat file which will call the actual VBScript script '''
batFile = open(batFile, 'w')
batFile.write("cscript " + vbsFileName + "\r\n");
batFile.close()

''' produce lines from list '''
def getLinesFromList(aList):
  commands = ""
  commands = commands.join([aListItem + "\r\n"   for aListItem in aList])
  return commands

''' define command sequence for a single file '''
def getVBSlinesForFile(fileName):
  commandLines = []
  commandLines.append("*** Rem File: " + fileName)
  commandLines.append("Set oApp = CreateObject(\"Powerpoint.Application\")")
  commandLines.append("oApp.Presentations.Open(strCurDir & \"\\" " & \"content-rename.pptm\")")
  commandLines.append("oApp.Presentations.Open(strCurDir & \"\\" " & " + fileName)
  commandLines.append("oApp.Run \"" + pptMacroFile + "\"!" + pptMacro + "\"")
  commandLines.append("oApp.ActivePresentation.Save")
  commandLines.append("oApp.Quit")
  commandLines.append("WScript.Sleep(2000)")
  commandLines.append(" ")
  return getLinesFromList(commandLines)
  
''' define script header '''
def getVBSScriptHeader():
  commandLines = []
  commandLines.append("Set WshShell = CreateObject(\"Wscript.Shell\")")
  commandLines.append("strCurDir = WshShell.CurrentDirectory")
  commandLines.append(" ")
  return getLinesFromList(commandLines)
  
  
''' write commands for each powerpoint file '''
powerpointFilter = re.compile(powerpointExt)

vbsFile = open(vbsFileName, 'w')
vbsFile.write(getVBSScriptHeader())

for path,dirs,files in os.walk("."):
  for fileName in files:
    if powerpointFilter.search(fileName) != None:
      vbsFile.write(getVBSlinesForFile(os.path.join(path,fileName)))

vbsFile.close()
