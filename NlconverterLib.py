import os 
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.header
import mimetypes
from email import encoders
import re
import tempfile

#Regexp
reAddressNotes = re.compile(r'CN=(.*?)\s+(.*?)\/OU=DGI\/OU=FINANCES\/O=GOUV\/C=FR', re.IGNORECASE)
reAddressMail = re.compile(r'([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,6})', re.IGNORECASE)

class NotesDocumentReader(object):
    """Base class for all documents"""
    def __init__(self):
        #compute a name for temporary files (attachments)
        self.tempname = os.path.join(tempfile.gettempdir(), 'nlc.tmp')

    def get(self, doc, itemname):
        """Helper to get an Item value in a document"""
        return doc.GetItemValue(itemname)

    def getDocumentType(self, doc):
        """Helper to get the 'Form' name"""
        return self.get1(doc, 'Form')

    def get1(self, doc, itemname):
        """Helper to get the first item value"""
        itemvalue = self.get(doc, itemname)
        if len(itemvalue):
            return itemvalue[0]
        else :
            return u''

    def checkDocumentType(self, doc):
        """Check if the document handling by the class"""
        return True
        
    def debug(self, doc):
        """Debug method : print all items values"""
        for it in doc.Items:
            try:
                print it, doc.GetItemValue(it)
            except:
                print it, "!! can't display item value !!"

    def matchAddress(value):
        """Convert Notes Address Name Space into emails"""
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
        """Return the list of the attachments, striping None and void names"""
        return filter(lambda x : x != None and x != u'', self.get(doc, "$FILE"))

    def hasAttachments(self, doc):
        """True if theyre are any attachments"""
        return len(self.listAttachments) > 0
        
    def extractAttachment(self, doc, f):
        """Extract an attachment from the document"""
        a = doc.GetAttachment(f)

        #FIXME : bug when there is \xa0 (non breaking space) in the filename. What to do then ?
        a.ExtractFile(self.tempname)

class NotesMemoReader(NotesDocumentReader):
    """Subclass for reading 'Memo' Notes Documents"""
    def checkDocumentType(self, doc):
        return self.getDocumentType(doc) == u'Memo'
      
class NotesDocumentConverter(NotesDocumentReader):
    """Base class for all converters"""
    pass

class NotesMemoToMimeConverter(NotesDocumentConverter):
    """Convert a Memo Document to a Mime Message"""
    charset = 'iso-8859-15' #default charset
    charsetAttachment = 'utf-8' #attachment filename charset. Because Linux and Windows seems to use Utf-8 for filenames...
    
    def stringToHeader(self, value):
        """Build a Mail header value from a string""" 
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
        """Build Mime Attachment 'f' from doc""" 
        self.extractAttachment(doc, f)
        fp = open(self.tempname, 'rb')
        msg = email.mime.base.MIMEBase('application', 'octet-stream')
        msg.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(msg)
        try:
          fname = f.encode(self.charsetAttachment)
        except :
            fname = f.encode(self.charset)
        msg.add_header('Content-Disposition', 'attachment', filename=fname)
        return msg

    def buildMessage(self, doc):
        """Build a message from doc"""
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
