# -*- coding: utf-8 -*-
# hugues.bernard@gmail.com
# Pour utiliser ce script :
# * Installer python 2.6 pour windows
# * Installer pywin 2.6 pour windows
# * (optionnellement) enregistrer la dll com de notes : "regsvr32 c:\notes\nlsxbe.dll"
# * Ouvrir le client Notes et les bases qu'il faut convertir
# * en ligne de commande (cmd) :
#   SET PATH=%PATH%;C:\Python26
#   python notes2mbox.py mot_de_passe_lotus c:\chemin\de\la\base.nsf
# => un fichier .mbox sera créé qu'il suffit de copier dans le répertoire ad-hoc de Thunderbird (ou d'un autre client...)

import sys
import os 
import mailbox
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.header
import mimetypes
from email import encoders
import re
import tempfile
import win32com.client
#LES APPELS COM SE FONT avec une majuscule

#Regexp
reAddressNotes = re.compile(r'CN=(.*?)\s+(.*?)\/OU=DGI\/OU=FINANCES\/O=GOUV\/C=FR', re.IGNORECASE)
reAddressMail = re.compile(r'([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,6})', re.IGNORECASE)

class NotesDocumentReader(object):
    def __init__(self):
        self.tempname = os.path.join(tempfile.gettempdir(), 'nlc.tmp')

    def get(self, doc, itemname):
        return doc.GetItemValue(itemname)

    def getDocumentType(self, doc):
        return self.get1(doc, 'Form')

    def get1(self, doc, itemname):
        itemvalue = self.get(doc, itemname)
        if len(itemvalue):
            return itemvalue[0]
        else :
            return u''

    def checkDocumentType(self, doc):
        return True
        
    def debug(self, doc):
        for it in doc.Items:
            try:
                print it, doc.GetItemValue(it)
            except:
                print it, "!! can't display item value !!"

    def matchAddress(value):
        res = reAddressNotes.search(value)
        if res == None:
            res = reAddressMail.search(value)
            if res == None:
                return value.lower()
            else :
                return res.group(1)
        else :
            return u"%s.%s@dgfip.finances.gouv.fr" % (res.group(1).lower(), res.group(2).lower())

    def listAttachments(self, doc):
        return filter(lambda x : x != None and x != u'', self.get(doc, "$FILE"))

    def hasAttachments(self, doc):
        return len(self.listAttachments) > 0
        
    def extractAttachment(self, doc, f):
        a = doc.GetAttachment(f)
        #FIXME : tester le \xa0
        a.ExtractFile(self.tempname)

class NotesMemoReader(NotesDocumentReader):
    def checkDocumentType(self, doc):
        return self.getDocumentType(doc) == u'Memo'
      
class NotesDocumentConverter(NotesDocumentReader):
    pass

class NotesMemoToMimeConverter(NotesDocumentConverter):
    charset = 'iso-8859-15'
    
    def stringToHeader(self, value):
        return email.header.Header(value, self.charset)
        
    def header(self, doc, itemname):
        return self.stringToHeader(self.get1(doc, itemname))
    
    def matchAddress(self, value):
        res = reAddressNotes.search(value)
        if res == None:
            res = reAddressMail.search(value)
            if res == None:
                return value.lower()
            else :
                return res.group(1)
        else :
            return u"%s.%s@dgfip.finances.gouv.fr" % (res.group(1).lower(), res.group(2).lower())

    def buildMailBody(self, doc):
        return email.mime.text.MIMEText(mc.get1(doc, 'Body'), _charset=self.charset)

    def addressHeader(self, doc, item):
        items = self.get(doc, item)
        return self.stringToHeader(",".join(map(self.matchAddress, items)))
    
    def messageHeaders(self, doc, m):
        m['Subject'] = self.header(doc, "Subject")
        m['From'] = self.addressHeader(doc, "From")
        m['To'] = self.addressHeader(doc, "sendto")
        m['Cc'] = self.addressHeader(doc, "copyto")
        m['Date'] = self.get1(doc, "PostedDate")
        if m['Date'] == u'':
            m['Date'] = self.get1(doc, "DeliveredDate")
        ccc = self.addressHeader(doc, "BlindCopyTo")
        if ccc != u'':
            m['Ccc'] = ccc
        m['User-Agent'] = self.header(doc, "$Mailer")
        m['Message-ID'] = self.header(doc, "$MessageID")

    def buildAttachment(self, doc, f):
        self.extractAttachment(doc, f)
        fp = open(self.tempname, 'rb')
        msg = email.mime.base.MIMEBase('application', 'octet-stream')
        msg.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(msg)
        msg.add_header('Content-Disposition', 'attachment', filename=f.encode(self.charset))
        return msg

    def buildMessage(self, doc):
        main = self.buildMailBody(doc)
        
        #files
        files = self.listAttachments(doc)
        if len(files) > 0 :
            m = email.mime.multipart.MIMEMultipart(charset=self.charset)
            m.set_charset(self.charset)
            self.messageHeaders(doc, m)
            m.attach(main)
            for f in files :
                msg = self.buildAttachment(doc, f)
                m.attach(msg)
        else:
            m = main
            self.messageHeaders(doc, m)
        return m

#Constantes
notesPasswd = "foobar"
notesNsfPath = "C:\\archive.nsf"
mailboxName = notesNsfPath+".mbox"

#Connection à Notes
session = win32com.client.Dispatch(r'Lotus.NotesSession')
session.Initialize(notesPasswd)
db = session.GetDatabase("", notesNsfPath)

#Création du fichier mbox
mbox = mailbox.mbox(mailboxName, None, True)

#all = tous les documents
all=db.AllDocuments
print "Nombre de documents :", all.Count

c = 0 #compteur de documents
e = 0 #compteur d'erreur à la conversion

mc = NotesMemoToMimeConverter()

doc = all.GetFirstDocument()
while doc and c < 200 and e < 99999:
    try:
        m = mc.buildMessage(doc)
        mbox.add (m)
        
    except Exception as ex:
        e += 1 #compte les exceptions
        print "\n--Exception for message %d (%s)" % (c, ex)
        mc.debug(doc)

    finally:
        doc = all.GetNextDocument(doc)
        c += 1

print "Exceptions a traiter manuellement:", e
mbox.close()
#FIXME : session OLE a cloturer
