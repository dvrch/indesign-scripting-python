from win32com.client import Dispatch

app = Dispatch("InDesign.Application")
# What to find?
app.FindTextPreferences.FindWhat = 'consequatur'
# Change to what?
app.ChangeTextPreferences.ChangeTo = 'molutem'

app.ActiveDocument.ChangeText()
# app.Selection[0].ChangeText()

idNothing = 1851876449  #from enum idNothingEnum, see doc_reference
#reset Preferences
app.FindTextPreferences.FindWhat = idNothing
app.ChangeTextPreferences.ChangeTo = idNothing
