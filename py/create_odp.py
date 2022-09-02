# http://geoff-thumbingatthemuse.blogspot.com/2008/08/my-introduction-to-pyuno-create-odp.html


# UNO stuff
import uno
import unohelper

# UNO exceptions
from com.sun.star.uno import Exception as UnoException, RuntimeException
from com.sun.star.connection import NoConnectException
from com.sun.star.lang import IllegalArgumentException, DisposedException
from com.sun.star.container import NoSuchElementException


class Outlayer(object)

    def __init__(self, schedule, options):
        self._bootstrap_uno()

    def _bootstrap_uno(self):
        """
        UNO initialisation, sets up the context
        """
        try:
            localContext = uno.getComponentContext()
            resolver = localContext.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", localContext )
            self.smgr = resolver.resolve("uno:socket,host=localhost,port=%s;"
                                         "urp;StarOffice.ServiceManager" % UNOPORT )
            remoteContext = self.smgr.getPropertyValue("DefaultContext")
            self.desktop = self.smgr.createInstanceWithContext("com.sun.star."
                                                               "frame.Desktop",
                                                               remoteContext)
        except NoConnectException, e:
            sys.stderr.write("OpenOffice process not found or "
                             "not listening (" + e.Message + ")\n")
            sys.exit(1)
        except IllegalArgumentException, e:
            sys.stderr.write("The url is invalid ( " + e.Message + ")\n")
            sys.exit(1)
        except RuntimeException, e:
            sys.stderr.write("An unknown error occured: " + e.Message + "\n")


def new_document(self):
        """
        returns a new document
        """
        # no start-up dialog
        document = self.desktop.loadComponentFromURL("private:factory/simpress",
                                                     "_blank",
                                                     0, ())
        # set a nasty class reference to save passing documents around
        # willy nilly
        self.document = document
        return document

def _new_page(self):
        """
        Create and return a new page
        """
        pages = self.document.getDrawPages()
        pages.insertNewByIndex(pages.getCount())
        return pages.getByIndex(pages.getCount() - 1)


def _insert_image(self, page, path_to_image):
    """
    Add an image to document, by scaling and centring it
    path_to_image -> filesystem path to a screenshot
    page -> created draw page
    """
    # convert filesystem path to OO url
    imageURL = unohelper.systemPathToFileUrl(path_to_image)
    # set the page name to the image filename (with the .png removed)
    pagename = os.path.splitext(os.path.basename(path_to_image))[0]
    # set the page name to match the filename (without the suffix)
    page.setName(pagename)
    # create the shape
    imageShape = self.document.createInstance("com.sun.star.drawing.GraphicObjectShape")
    # set the GraphicURL to be the filesystem URL
    imageShape.GraphicURL = imageURL
    # add the graphic shape to the page
    page.add(imageShape)
    # get the object dimensions
    imageBitmapSize = imageShape.GraphicObjectFillBitmap.getSize()
    dImageRatio = float(imageBitmapSize.Height) / float(imageBitmapSize.Width)
    dPageRatio = float(page.Height) / float(page.Width)
    # new size object
    oNewSize = Size()
    #Compare the ratios to see which is wider, relatively speaking.
    if dPageRatio > dImageRatio:
        oNewSize.Width = page.Width
        oNewSize.Height = int(float(page.Width) * dImageRatio)
    else:
        oNewSize.Width = int(float(page.Height) / dImageRatio)
        oNewSize.Height = page.Height
    # new position object
    oPosition = Point()
    # Center the image
    oPosition.X = (page.Width - oNewSize.Width)/2
    oPosition.Y = (page.Height - oNewSize.Height)/2
    # resize the image
    imageShape.setSize(oNewSize)
    # and bang it in the middle
    imageShape.setPosition(oPosition)
