import os
import re

msofficePowerpoint = "C:\Program Files (x86)\Microsoft Office\Office14\POWERPNT.EXE" # znajdz sobie swoja sciezke
pptMacroFile = "content-rename.pptm" # znajdz sobie skrypt
pptMacro = "ReplaceSKBRAndSaveExit" # taka jest na stale
powerpointExt = "(\.pptx$)|(\.ppt$)"
scriptFile = "content-rename.bat"

file = open(scriptFile, 'w')

powerpointFilter = re.compile(powerpointExt)
for path,dirs,files in os.walk("."):
  for fileName in files:
    if powerpointFilter.search(fileName) != None:
      command = "\"" + msofficePowerpoint + "\"" \
	    + " " + "/M" + " " + "\"" + pptMacroFile + "\"" + " " + "\"" + pptMacro + "\"" \
	    + " " + "\"" + os.path.join(path,fileName) + "\""
      file.write(command + "\r\n")

file.close()
