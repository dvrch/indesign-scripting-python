import win32com.client
app = win32com.client.Dispatch("InDesign.Application.CC.2023")  # ou "InDesign.Application"
print(app.Name)



