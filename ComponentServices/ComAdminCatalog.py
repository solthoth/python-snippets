import logging
import os
import win32com.client as comClient

def adminApplicationCatalog():
    return comClient.Dispatch("COMAdmin.COMAdminCatalog")
    

def applicationCollection(catalog):
    applicationCollection = catalog.GetCollection("Applications")
    applicationCollection.Populate()
    return applicationCollection

def applicationExists(comCollection, appName: str):
    logging.debug('Listing out existing applications')
    for application in comCollection:
        existingAppName = application.Name
        logging.debug(existingAppName)
        if appName == existingAppName:
            return True
    return False

def createApplication(adminCatalog, appName: str):
    collection = applicationCollection(adminCatalog)
    if not applicationExists(collection, appName):
        defineApplication(collection.Add(), appName)
        collection.SaveChanges()
        return True
    return False

def defineApplication(newApp, name: str):
    # Reference to all Application properties:
    # https://docs.microsoft.com/en-us/windows/win32/cossdk/applications
    logging.debug('Reference %s for additional properties', 'https://docs.microsoft.com/en-us/windows/win32/cossdk/applications')
    newApp.SetValue('Name', name)
    newApp.SetValue('ApplicationAccessChecksEnabled', 0)
    newApp.SetValue('AccessChecksLevel', 0)

def main(appName: str):
    catalog = adminApplicationCatalog()
    created = createApplication(catalog, appName)
    logging.info('Application %s created: %s', appName, created)

if __name__ == '__main__':
    logging.basicConfig(level=os.environ.get("LOGLEVEL", "DEBUG"))
    main('TestService')