#https://stackoverflow.com/questions/29417113/how-to-call-an-existing-libreoffice-python-macro-from-a-python-script

#!/usr/bin/python3
# -*- coding: utf-8 -*-
##
# a python script to run a libreoffice python macro externally
# NOTE: for this to run start libreoffice in the following manner
# soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;" --writer --norestore
# OR
# nohup soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;" --writer --norestore &
#
import uno
#from com.sun.star.connection import NoConnectException
#from com.sun.star.uno  import RuntimeException
#from com.sun.star.uno  import Exception
#from com.sun.star.lang import IllegalArgumentException

def uno_directmacro(*args):
    localContext = uno.getComponentContext()
    localsmgr = localContext.ServiceManager
    resolver = localsmgr.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        print ("LibreOffice is not running or not listening on the port given - ("+e.Message+")")
        return
    msp = ctx.getValueByName("/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory")
    sp = msp.createScriptProvider("")
    scriptx = sp.getScript('vnd.sun.star.script:directmacro.py$directmacro?language=Python&location=user')
    try:
        scriptx.invoke((), (), ())
    except Exception as e:
        print ("Script error ( "+ e.Message+ ")")
        print(e)
        return
    return(None)

uno_directmacro()