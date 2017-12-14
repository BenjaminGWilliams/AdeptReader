'''
Created on 6/11/2017

@author: bgw
'''
from os.path import basename
# import glob
import random
# import numpy
import datetime
import csv
# from _sqlite3 import Row



class DMPFile(object):
    '''
    classdocs
    '''


    def __init__(self, dmpfilepath):
        '''
        Constructor
        '''
        self.dmpfilepath = dmpfilepath
        self.dmpfilename = basename(dmpfilepath).replace(".dmp", "").replace(".DMP", "")
        self.LOOKUPLIST = [['9916', 'Group', 2, 0, 1, "-", 1400000],
                           ['9921', 'Line', 4, 2, 3, 23, 300000],
                           ['9922', 'Load', 3, 0, 1, 18, 500000],
                           ['9923', 'Node', 2, 0, 1, 10, 100000],
                           ['9926', 'Source', 3, 0, 1, 18, 400000],
                           ['9927', 'Switch', 4, 2, 3, 18, 600000],
                           ['9928', 'Transformer', 4, 2, 3, 39, 900000]]

        self.listofcodes = [item[0] for item in self.LOOKUPLIST]
        self.listoffullnames = [item[1] for item in self.LOOKUPLIST]
        self.listofnameindex = [item[2] for item in self.LOOKUPLIST]
        self.fromlist = [item[3] for item in self.LOOKUPLIST]
        self.tolist = [item[4] for item in self.LOOKUPLIST]
        self.groupposlist = [item[5] for item in self.LOOKUPLIST]

    def setheader(self, **kwargs):

        stringtime = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        createdbystring = self.encryptcode('Created by bgw on {}'.format(stringtime))


        defaultcommentlist = [createdbystring, '', '', '', '', '', '']

        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()
        writelist = []
        for rowstr in textlist:
            splitrowstring = rowstr.split(" ")

            if splitrowstring[0] == '9914':
                # rowstr = rowstr.split(" ")
                splitrowstring[2] = kwargs.get('basekva', '100000')
                for commentplace in range(1, 8):
                    splitrowstring[commentplace + 2] = '"' + kwargs.get('comment' + str(commentplace),
                                                          defaultcommentlist[commentplace - 1]) + '"'
                # writestr = ""
                # for item in rowstr:
                    # writestr += str(item) + " "
                # rowstr = writestr[:-1] + '\n'
            writelist.append(' '.join(splitrowstring))

        dmpfile = open(self.dmpfilepath, "w")
        dmpfile.writelines(writelist)
        print "Completed setting network properties"

    def findhighestcode(self, objecttype):
        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()

        if objecttype == 'Group':
            highestcode = 1400000
        if objecttype == 'Line':
            highestcode = 300000
        for rowstr in textlist:
            rowstr = rowstr.split(" ")

            if rowstr[0] == '9916' and objecttype == 'Group':
                if int(rowstr[1]) > highestcode:
                    highestcode = int(rowstr[1])
            elif rowstr[0] == '9921' and objecttype == 'Line':
                if int(rowstr[1]) > highestcode:
                    highestcode = int(rowstr[1])

        return highestcode

    def findnodespecs(self, nodename):

        "Returns node id, x and y coordinates"
        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()
        for rowstr in textlist:

            if rowstr[0:4] == '9952':
                rowstr = rowstr.split(" ")
                if rowstr[2] == nodename:
                    return [rowstr[1], rowstr[3], rowstr[4]]
        return

    def findgroupspecs(self, groupname):

        "Returns group id"

        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()
        for rowstr in textlist:

            if rowstr[0:4] == '9916':
                rowstr = rowstr.split(" ")
                if rowstr[2] == groupname:
                    return rowstr[1]
        return


    def addalltogroup(self, groupname):
        listofgroups = self.listobjects('Group')
        quotedgroupname = '"' + groupname + '"'

        if quotedgroupname not in listofgroups:
            self.addgroup(groupname)





        writelist = []

        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()

        groupid = self.findgroupspecs(quotedgroupname)

        for rowstr in textlist:
            splitrowstr = rowstr.split(" ")

            if splitrowstr[0] in self.listofcodes and splitrowstr[0] != '9916':
                numgroupsindex = self.groupposlist[self.listofcodes.index(splitrowstr[0])]
                splitrowstr[numgroupsindex] = str(int(splitrowstr[numgroupsindex]) + 1)
                splitrowstr.insert(-1, groupid)

                rowstr = ''
                for item in splitrowstr:
                    rowstr += str(item) + " "
                rowstr = rowstr[:-1]


            writelist.append(rowstr)

        dmpfile = open(self.dmpfilepath, "a")
        for item in writelist:
            dmpfile.writelines(item)
        dmpfile.close()

            # dmpfile.close




    def addgroup(self, groupname, description=''):
        listofgroups = self.listobjects('Group')
        highestcode = self.findhighestcode('Group')
        if listofgroups == []:
            highestcode -= 1
        if groupname not in listofgroups:
            dmpfile = open(self.dmpfilepath, "a")
            dmpfile.write('\n9916 {} "{}" "{}" \n'.format(str(highestcode + 1), groupname, description))

    def addline(self, linename, node1, node2, **kwargs):
        listofobjects = self.listallobjects()
        listofnodes = self.listobjects('Node')
        node1 = '"' + node1 + '"'
        node2 = '"' + node2 + '"'
        if linename not in listofobjects and node1 in listofnodes and node2 in listofnodes:
            linecode = self.findhighestcode('Line')
            dmpfile = open(self.dmpfilepath, "a")
            node1specs = self.findnodespecs(node1)
            node2specs = self.findnodespecs(node2)

            linelength = kwargs.get('linelength', 1)

            posseqres = kwargs.get('posseqres', 0.15)
            posseqrea = kwargs.get('posseqrea', 0.61)
            zeroseqres = kwargs.get('zeroseqres', 0.36)
            zeroseqrea = kwargs.get('zeroseqrea', 2)
            posseqadm = kwargs.get('posseqadm', 7.2)
            zeroseqadm = kwargs.get('zeroseqadm', 3.6)

            current1 = kwargs.get('current1', 280)
            current2 = kwargs.get('current2', 300)
            current3 = kwargs.get('current3', 320)
            current4 = kwargs.get('current4', 340)

            contype = kwargs.get('contype', "USER")

            linestr = '\n9921 {} {} {} "{}" 1 {} {} {} {} {} {} {} 7 999 999 999 999 {} {} {} {} "{}" 0\n'.format(linecode, node1specs[0], node2specs[0], linename, linelength, posseqres, posseqrea, posseqadm, zeroseqres, zeroseqrea, zeroseqadm, current1, current2, current3, current4, contype)
            linestr2 = '9953 {} "{}" {} {} 0 \n'.format(linecode, linename, node1specs[1], node1specs[2])
            linestr3 = '9954 {} "{}" {} {} {} {} \n'.format(linecode, linename, node1specs[1], node1specs[2], node2specs[1], node2specs[2])

            dmpfile.write(linestr)
            dmpfile.write(linestr2)
            dmpfile.write(linestr3)

            dmpfile.close()
            return True


        return False

    def listobjects(self, objectname):

        if objectname not in self.listoffullnames:
            print "%s is not a valid objectname. Please select from:" % objectname
            for item in self.listoffullnames:
                print item
            return
        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()

        listofobjects = []
        objectcode = self.listofcodes[self.listoffullnames.index(objectname)]
        rowstringoffset = self.listofnameindex[self.listoffullnames.index(objectname)]

        for rowstr in textlist:
            rowstr = rowstr.split(" ")

            if rowstr[0] == objectcode:
                listofobjects.append(rowstr[rowstringoffset])
        return listofobjects

    def listallobjects(self):
        dmpfile = open(self.dmpfilepath, "r")
        textlist = dmpfile.readlines()

        objectlist = []
        for rowstr in textlist:
            rowstr = rowstr.split(" ")
            rowcode = rowstr[0]
            if rowcode in self.listofcodes:
                objectlist.append(rowstr[self.listofnameindex[self.listofcodes.index(rowcode)]])
        return objectlist


    def correctdmp(self):

        dmpfile = open(self.dmpfilepath, "r")
        listofnames = []



        writelist = []

        rowcount = 0
        rowstring = 'START'
        while rowstring != '':
            rowcount += 1
            rowstring = dmpfile.readline()

            splitrowstring = rowstring.split(" ")
            # print rowstring
            objectcode = splitrowstring[0]

            try:
                positionindex = self.listofcodes.index(objectcode)
            except ValueError:
                writelist.append(rowstring)
                continue

            objecttypename = self.listoffullnames[positionindex]
            objectname = splitrowstring[self.listofnameindex[positionindex]]

            fromname = splitrowstring[self.fromlist[positionindex]]
            toname = splitrowstring[self.tolist[positionindex]]

            if fromname == toname:
                print ("{} name: {} has same from node as to node ({}) on line: {}. "
                       "Removing...".format(objecttypename, objectname, fromname, rowcount))
                continue



            if objectname in listofnames:
                print("Found duplicate {} name: {} on line: {}. "
                      "Adding '~' to end".format(objecttypename, objectname, rowcount))
                oldobjectname = objectname
                while objectname in listofnames:
                    objectname = objectname[:-1] + 'GN"'
                listofnames.append(objectname)
                newrowstring = rowstring.replace(oldobjectname, objectname)
                writelist.append(newrowstring)
            else:
                listofnames.append(objectname)
                writelist.append(rowstring)

        dmpfile = open(self.dmpfilepath, "w")
        dmpfile.writelines(writelist)

        print "Completed removing lines that have same from and to node and renaming duplicate names"

    #===============================================================================================
    # def lookup(self, filetype, value):
    #     for f in glob.glob(filetype):
    #         r = numpy.recfromcsv(f)
    #         for item in r:
    #             if item[0] == value:
    #                 return item[1]
    #         return
    #===============================================================================================

    def createlookupdic(self):

        returndic = {}

        with open("encrypt.csv") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                returndic[row['Whatitis']] = row['Code']
            return returndic




    def decryptcode(self, codestring):

        codedic = self.createlookupdic()
        returnstring = ""
        mergelist = [i + j for i, j in zip(codestring[::2], codestring[1::2])]
        for items in mergelist:
            for name, code in codedic.items():
                if code == items:
                    returnstring += name

            # returnstring += self.lookup("decrypt.csv", items)
        print "Your de-crypted code is: %s" % returnstring
        return returnstring

    def encryptcode(self, rawstring):

        codedic = self.createlookupdic()

        returnstring = ""
        for items in rawstring:
            returnstring += codedic[items]
        print "Your encrypted code is: %s" % returnstring
        return returnstring

    def offsetnodes(self, **kwargs):
        dmpfile = open(self.dmpfilepath, "r")
        listofcoords = []
        writelist = []

        lowerbound = int(kwargs.get('lowerbound', -10) * 10)
        upperbound = int(kwargs.get('upperbound', 10) * 10)
        scaling = kwargs.get('scaling', 0.05)


        rowcount = 0
        rowstring = 'START'
        while rowstring != '':
            rowcount += 1

            rowstring = dmpfile.readline()

            splitrowstring = rowstring.split(" ")
            # print rowstring
            objectcode = splitrowstring[0]
            if objectcode != '9952':
                writelist.append(rowstring)
            else:
                xcoord = float(splitrowstring[2])
                newxcoord = xcoord
                ycoord = float(splitrowstring[3])
                newycoord = ycoord

                coord = [xcoord, ycoord]
                while coord in listofcoords:

                    print "Duplicate co-ords: {:10.2f},{:10.2f} on row:{} for node {}. Shifting...".format(coord[0], coord[1], rowcount, splitrowstring[1])
                    xoffset = random.randint(lowerbound, upperbound) / 10.0
                    yoffset = random.randint(lowerbound, upperbound) / 10.0

                    newxcoord += xoffset
                    newycoord += yoffset
                    coord = [newxcoord, newycoord]
                listofcoords.append(coord)
                # rowstring.replace(str(xcoord), str(newxcoord))
                # rowstring.replace(str(ycoord), str(newycoord))
                splitrowstring[2] = str(newxcoord * scaling)
                splitrowstring[3] = str(newycoord * scaling)
                writelist.append(' '.join(splitrowstring))

        dmpfile = open(self.dmpfilepath, "w")
        dmpfile.writelines(writelist)
        print "Completed scaling model and offseting nodes that sit on top of each other"
