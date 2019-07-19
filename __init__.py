from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile,io, re
name = "docscan"
class Docscan():
    '''
        Class that scans documents (pdf, doc and docx) and returns their strings,
        as well as can execute regular expressions.
    '''
    def __init__(self,fileName):
        self.filePathName=fileName
        self.fileName=fileName.rsplit('/', 1)[-1]
        self.fileFormat=fileName.rsplit('.',1)[-1]

    def returnFileText(self, xmlHeader='word/document.xml'):
        '''
        Returns the file's text in a large string
        Usage:
            returnFileText()
        Returns:
            'Lorem ipsum ....'
        '''
        if self.fileFormat == 'pdf':
            fp = open(self.filePathName, 'rb')
            rsrcmgr = PDFResourceManager()
            retstr = io.StringIO()
            codec = 'utf-8'
            laparams = LAParams()
            device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for page in PDFPage.get_pages(fp):
                interpreter.process_page(page)
                file =  retstr.getvalue()

        elif self.fileFormat=='doc':
            file=open(self.filePathName,encoding='latin-1').read()

        elif self.fileFormat=='docx':
            WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            PARA = WORD_NAMESPACE + 'p'
            TEXT = WORD_NAMESPACE + 't'
            """
            Take the path of a docx file as argument, return the text in unicode.
            """
            document = zipfile.ZipFile(self.filePathName)
            xml_content = document.read(xmlHeader)
            document.close()
            try:
                tree = XML(xml_content)
                paragraphs = []
                for paragraph in tree.getiterator(PARA):
                    texts = [node.text
                            for node in paragraph.getiterator(TEXT)
                            if node.text]
                    if texts:
                        paragraphs.append(''.join(texts))
            except:
                return Exception
            file='\n\n'.join(paragraphs)

        return file

    
    def executeRegex(self, regularExpression):
        '''
            Executes specified regular expression on any given document that is a doc, docx or pdf.
            Usage:
                executeRegex('regularexpression')
            Returns:
                listOfElements-satisfied-by-regex
        '''
        fileText = self.returnFileText()
        listOfRegex = re.findall(regularExpression, fileText)
        return listOfRegex

    def executeHeaderRegex(self, regularExpression):
        '''
            Executes specified regular expression on given file that SUPPORTS header XML.
            Usage:
                executeHeaderRegex('regularexpression')
            Returns:
                listOfElements-satisfied-by-regex
        '''
        if self.fileFormat != "docx":
            return ('[WARN]   Cannot check header, '+self.fileFormat+' has no XML properies')
        elif self.fileFormat.lower() == 'docx':
            listOfRegex=[]
            count=1
            while len(listOfRegex) <= 0 and count < 3:
                headerText=self.returnFileText('word/header'+str(count)+'.xml')
                listOfRegex = re.findall(regularExpression, headerText)
                count +=1
            if len(listOfRegex) <= 0:
                return '[WARN]   No Headers found'
            else:
                return listOfRegex
        
    def executeFooterRegex(self, regularExpression):
        '''
            Executes specified regular expression on given file that SUPPORTS footer XML.
            Usage:
                executeFooterRegex('regularexpression')
            Returns:
                listOfElements-satisfied-by-regex
        '''
        if self.fileFormat != "docx":
            return ('[WARN]   Cannot check footer, File Format "'+self.fileFormat+'" has no XML properies')
        elif self.fileFormat.lower() == 'docx':
            listOfRegex=[]
            count=1
            while len(listOfRegex) <= 0 and count < 3:
                footerText=self.returnFileText('word/footer'+str(count)+'.xml')
                listOfRegex = re.findall(regularExpression, footerText)
                count +=1
            if len(listOfRegex) <= 0:
                return '[WARN]   No Footers found'
            else:
                return listOfRegex
