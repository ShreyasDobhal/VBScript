set shell = createObject("wscript.shell")
pass = inputBox("Enter a Password")
if pass = "1234" then shell.Run("Notepad.exe") else msgBox("Incorrect !")