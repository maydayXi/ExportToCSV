import os
import unicodedata

from com.sun.star.beans import PropertyValue


def csv_properties():
    '''Build the dialog parameter for UTF-8 CSV'''
    props = []
    p = PropertyValue()
    p.Name = 'FilterName'
    p.Value = 'Text - txt - csv (StarCalc)'
    props.append(p)
    p = PropertyValue()
    p.Name = 'FilterOptions'
    p.Value = '44,34,76,1,,0,false,true,true,false'
    props.append(p)
    return tuple(props)


def export_sheets_to_csv():
    '''Iter over each sheet and save it as CSV file. '''
    desktop = XSCRIPTCONTEXT.getDesktop()  # noqa
    model = desktop.getCurrentComponent()
    controller = model.getCurrentController()
    dirname = os.path.dirname(model.URL)
    for sheet in model.Sheets:
        controller.setActiveSheet(sheet)
        name = sheet.getName().lower().replace(' ', '-')
        name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore')
        filename = '{0}/{1}.csv'.format(dirname, name.decode('ascii'))
        model.storeToURL(filename, csv_properties())

g_exportedScripts = export_sheets_to_csv,