# -*- coding: utf-8 -*-
# hugues.bernard@gmail.com
# Pour utiliser ce script :
# * Installer python 2.6 pour windows
# * Installer pywin 2.6 pour windows
# * (optionnellement) enregistrer la dll com de notes : "regsvr32 c:\notes\nlsxbe.dll"
# * en ligne de commande (cmd) :
#   SET PATH=%PATH%;C:\Python26
#   **pour l'instant** fixer notesPasswd et notesNsfPath plus bas
#   python notes2mbox.py 
# => un fichier .mbox sera créé qu'il suffit de copier dans le répertoire ad-hoc de Thunderbird (ou d'un autre client...)

import sys
import mailbox
import win32com.client #NB : Les appels COM se font avec une majuscule en début de méthode
import NlconverterLib

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
ac = all.Count
print "Nombre de documents :", ac

c = 0 #compteur de documents
e = 0 #compteur d'erreur à la conversion

mc = NlconverterLib.NotesMemoToMimeConverter()

doc = all.GetFirstDocument()
while doc and c < 99999 and e < 99999:
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
        if (c % 100) == 0:
            sys.stderr.write("%.1f%%\n" % float(100.*c/ac) )
print "Exceptions a traiter manuellement:", e
mbox.close()
#FIXME : session OLE a cloturer
