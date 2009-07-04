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

#Helpers pour accéder à COM
def get(doc, itemname):
    return doc.GetItemValue(itemname)

def get1(doc, itemname):
    itemvalue = get(doc, itemname)
    if len(itemvalue):
        return itemvalue[0]
    else :
        return u''

def makeheader(value, charset = 'iso-8859-15'):
    return email.header.Header(value, charset)

def header(doc, itemname):
    return makeheader(get1(doc, itemname))

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
        
def addressHeader(doc, item):
    items = get(doc, item)
    return makeheader(",".join(map(matchAddress, items)))

def messageHeaders(doc, m):
    m['Subject'] = header(doc, "Subject")
    m['From'] = addressHeader(doc, "From")
    m['To'] = addressHeader(doc, "sendto")
    m['Cc'] = addressHeader(doc, "copyto")
    m['Date'] = get1(doc, "PostedDate")
    if m['Date'] == u'':
        m['Date'] = get1(doc, "DeliveredDate")
    ccc = addressHeader(doc, "BlindCopyTo")
    if ccc != u'':
        m['Ccc'] = ccc

#Constantes
notesPasswd = "foobar"
notesNsfPath = "C:\\archive.nsf"
mailboxName = notesNsfPath+".mbox"
tempdir = tempfile.mkdtemp('_nlconverter')

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

doc = all.GetFirstDocument()
while doc and c < 99999 and e < 5:
    try:
        main = email.mime.text.MIMEText(doc.GetItemValue("Body")[0], _charset='iso-8859-15')
        
        #files
        files = filter(lambda x : x != None and x != u'', get(doc, "$FILE"))
        if len(files) > 0 :
            m = email.mime.multipart.MIMEMultipart(charset='iso-8859-15')
            m.set_charset('iso-8859-15')
            messageHeaders(doc, m)
            m.attach(main)
            for f in files :
                a = doc.GetAttachment(f)
                if a == None :
                    continue
                fpath = os.path.join(tempdir, f)
                a.ExtractFile(fpath)
                ctype, encoding = mimetypes.guess_type(fpath)
                if ctype is None or encoding is not None:
                    # No guess could be made, or the file is encoded (compressed), so
                    # use a generic bag-of-bits type.
                    ctype = 'application/octet-stream'
                maintype, subtype = ctype.split('/', 1)
                fp = open(fpath, 'rb')
                msg = email.mime.base.MIMEBase(maintype, subtype)
                msg.set_payload(fp.read())
                fp.close()
                encoders.encode_base64(msg)

                msg.add_header('Content-Disposition', 'attachment', filename=f.encode('utf-8'))
                m.attach(msg)
                os.remove(fpath)
        else:
            m = main
            messageHeaders(doc, m)

        mbox.add (m)
        
    except Exception as ex:
        e += 1 #compte les exceptions
        print "-----------Exception, message %d" % c
        print ex
        for it in doc.Items:
            print it, doc.GetItemValue(it)

    finally:
        doc = all.GetNextDocument(doc)
        c += 1

print "Exceptions a traiter manuellement:", e
mbox.close()
os.rmdir(tempdir)
#FIXME : session OLE a cloturer
