'''
Created on 3/11/2017

@author: bgw
'''
import getopt
import sys

from dmpfile import DMPFile as dmp


def helpme():
    ''' Prints out how to use this program '''

    print "DMPreader.py"
    print "\n\nCommands: \n"
    print "-h: Brings up help (this), or"
    print "-e 'STRINGTOENCRYPY': Encrypts 'STRINGTOENCRPYT'"
    print "-d 'STRINGTODECRYPT': Decrypts 'STRINGTODECRYPT'"
    print "-c 'path/to/dmpfile.DMP': Complete tidy up (network property, overlap, etc...)"
    print "-o 'path/to/dmpfile.DMP': Moves nodes in .DMP file to stop overlap"
    print "-r 'path/to/dmpfile.DMP': Cleans .DMP file"
    exit(1)

def main(argv):


    ''' Performs the required function call depending on the input arguments '''

    try:
        opts, args = getopt.getopt(argv, "hr:e:d:o:c:s:")
    except getopt.GetoptError:
        helpme()
    if opts == []:
        helpme()
    for opt, arg in opts:

        if opt == '-s':
            dmpf = dmp(arg)
            dmpf.sortsources()
            # dmpf.setsource()
        elif opt == '-r':
            dmpf = dmp(arg)
            dmpf.correctdmp()

        elif opt == '-c':
            dmpf = dmp(arg)
            dmpf.correctdmp()
            dmpf.offsetnodes()
            dmpf.setheader()
            dmpf.sortsources()

        elif opt == '-e':
            dmpf = dmp(r'C:\Temp\TEST.DMP')
            dmpf.encryptcode(arg)
        elif opt == '-d':
            dmpf = dmp(r'C:\Temp\TEST.DMP')
            dmpf.decryptcode(arg)
        elif opt == '-o':
            dmpf = dmp(arg)
            dmpf.offsetnodes()

        else:
            helpme()


if __name__ == '__main__':
    main(sys.argv[1:])
