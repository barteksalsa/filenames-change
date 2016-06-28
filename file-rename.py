import os
import re

searchText = "sk {0,1}-{0,1} {0,1}br( {0,1}-{0,1} {0,1}3){0,1}"
replaceText = "HCT15"
scriptFile = "file-rename.bat"

searchRegexp = re.compile(searchText, re.IGNORECASE)
file = open(scriptFile, 'w')

allFiles = []
for path,dirs,files in os.walk("."):
  for fileName in files:
    if re.search(searchRegexp, fileName) != None:
      fullFileName = os.path.join(path,fileName)
      fileNameReplaced = re.sub(searchRegexp, replaceText, fileName, count=1)
      fullFileNameReplaced = os.path.join(path,fileNameReplaced)
      allFiles.append((fullFileName, fullFileNameReplaced))
    
for (old,new) in allFiles:
  command = "move " + "\"" + old + "\"" + " " + "\"" + new + "\""
  file.write(command + "\r\n")
  #os.rename(old,new)

file.close()
