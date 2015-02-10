
# --------------------------------------------------------
# Python batch image processing using Photoshop Scripting
# --------------------------------------------------------
# * Photoshop Scripting docs:
#   http://www.adobe.com/devnet/photoshop/scripting.html
# * Currently tested with Ps CS3 - CS6

from comtypes.client import CreateObject
import os
import logging

from fnmatch import fnmatch


def logToFile():
    LOG_FILE = 'log_file.txt'
    logging.basicConfig(filename=LOG_FILE, level=logging.DEBUG)
    logging.debug('This message should go to the log file')
    logging.exception('Exception!')

def imageProcess():
    try:
        while True:
            fPath = raw_input('Files path(source): ')
            # if folder exists
            if not os.path.exists(fPath):
                print 'Path doesn`t exist!'
                imageProcess()
            # if there're files in source folder'
            elif not len(os.listdir(fPath)):
                print 'No files found in ' + fPath + '!'
                imageProcess()
            else:
                svPath = raw_input('Destination(if no will be created): ')
                if not os.path.exists(svPath):
                    os.mkdir(svPath, 0777)
                quVal = raw_input('JPEG Quality(0-10, 10 better): ')
                docRes = raw_input('Resolution(px): ')
                docWidth = raw_input('New width(px): ')
                for root, dirs, files in os.walk(fPath):
                    for f in files:
                        for ext in ('*.jpg', '*jpeg', '*.jpe', '*.tif',
                        '*.tiff', '*.png', '*.psd', '*.pdd', '*.gif'):
                            if fnmatch(f, ext):
                                pShop = CreateObject('Photoshop.Application')
                                # no dialogs when open file
                                pShop.DisplayDialogs = 3
                                # set ruler units to pixels. for resizing
                                pShop.Preferences.RulerUnits = 1
                                f = os.path.join(root, f)
                                psDoc = pShop.Open(f)
                                psDoc = pShop.Application.ActiveDocument
                                # document proportion calculation
                                docRatio = psDoc.Width / psDoc.Height
                                psDoc.Flatten()
                                # Change to RGB - 2
                                psDoc.ChangeMode(2)
                                # resizing image proportionally
                                # last argument set filenaming lowercase
                                if docWidth != '':
                                    psDoc.ResizeImage(int(docWidth),
                                    int(docWidth) / docRatio, int(docRes), 4)
                                else:
                                    psDoc.ResizeImage(psDoc.Width,
                                    psDoc.Height, int(docRes), 4)
                                opts = CreateObject('Photoshop.JPEGSaveOptions')
                                opts.EmbedColorProfile = False
                                opts.Quality = int(quVal)
                                opts.Matte = 1
                                psDoc.SaveAs(svPath, opts, False, 2)
                                print 'Saving... ', f
                                # close document without y/n prompt
                                psDoc.Close(2)
                print 'Total files saved: ' + str(len(os.listdir(svPath)))
    except:
            # uncomment below line to write traceback to logfile
            # logToFile()
            raise

if __name__ == '__main__':
    imageProcess()