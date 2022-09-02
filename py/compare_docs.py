#!/usr/bin/env python

import uno
from com.sun.star.beans import PropertyValue

url = uno.systemPathToFileUrl('d:\\changed_document.doc')
url_original = uno.systemPathToFileUrl('d:\\original_document.doc')
url_save = uno.systemPathToFileUrl('d:\\the_diff.pdf')

### Get Service Manager
context = uno.getComponentContext()
resolver = context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", context)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager

### Load document

properties = []
p = PropertyValue()
p.Name = "Hidden"
p.Value = True
properties.append(p)
properties = tuple(properties)

desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL(url, "_blank", 0, properties)

### Compare with original document
properties = []
p = PropertyValue()
p.Name = "URL"
p.Value = url_original
properties.append(p)
properties = tuple(properties)

dispatch_helper = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
dispatch_helper.executeDispatch(doc.getCurrentController().getFrame(), ".uno:CompareDocuments", "", 0, properties)

### Save File
properties = []
p = PropertyValue()
p.Name = "Overwrite"
p.Value = True
properties.append(p)
p = PropertyValue()
p.Name = "FilterName"
p.Value = 'writer_pdf_Export'
properties.append(p)
properties = tuple(properties)

doc.storeToURL(url_save, properties)
doc.dispose()