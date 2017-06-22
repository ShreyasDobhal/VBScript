message=inputBox("Enter text to be read ","Speak")
set sapi=createObject("sapi.spvoice")
sapi.speak message