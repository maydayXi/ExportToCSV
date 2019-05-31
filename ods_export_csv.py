import os
from com.sun.star.beans import PropertyValue

def export_to_csv():
	
	doc = XSCRIPTCONTEXT.getDocument()
	sheets = doc.Sheets
	dirname = os.path.dirname(doc.URL)
	controller = doc.getCurrentController()

	for sheet in sheets:
		controller.setActiveSheet(sheet)

		prop = []
		p = PropertyValue()
		p.Name = "FilterName"
		p.Value = "Text - txt - csv (StarCalc)"
		prop.append(p)
		p = PropertyValue()
		p.Name = "FilterOptions"
		p.Value = "44,34,76,1,,0,false,true,true,false"
		prop.append(p)

		name = sheet.getName()
		filename = "{0}/{1}.csv".format(dirname, name)
		doc.storeToURL(filename, tuple(prop))