__author__ = "R.K."

import win32com.client
import os

inputFolder = r'C:\Users\RK\Downloads\costing_remits'
outputFolder = r'C:\Users\RK\Downloads\costing_remits\pdf'

for file in os.listdir(inputFolder):
    if file.endswith(".msg"):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        filePath = inputFolder + '\\' + file
        msg = outlook.OpenSharedItem(filePath)
        att = msg.Attachments
        for i in att:
            i.SaveAsFile(os.path.join(outputFolder, i.FileName))


print(f"\n{len(os.listdir(outputFolder))} attachments from {len(os.listdir(inputFolder))} emails "
      f"were saved to folder: {outputFolder}")
