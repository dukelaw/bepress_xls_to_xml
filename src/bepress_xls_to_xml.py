#!/usr/bin/python

#=============================================================================
# Transform a Excel 97-2003 spreadsheet to an XML file suitable for loading 
# via XML batch upload for   
#=============================================================================

from xlrd import open_workbook, xldate_as_tuple
from lxml import etree
import optparse

def update_text(s, document, record):
    element = etree.SubElement(document, s)
    if type(record[s]) is float:
        element.text = "%d" % record[s]
    elif type(record[s]) is int:
        element.text =" %s" % record[s]
    else:
        element.text = "%s" % record[s]
    return element

def main():
    usage = "usage: %prog [options] arg"
    parser = optparse.OptionParser(usage)
    parser.add_option("-f", "--filename", dest="filename",
                      help="read data from FILENAME")
    parser.add_option("-o", "--output", dest="output",
                      help="read data from FILENAME")
    parser.add_option("-j", "--journal", dest="journal",
                      help="bepress directory name for journal")
    parser.add_option("-s", "--sheet", dest="sheet_index", default=0,
                      help="sheet index") 
    (options, args) = parser.parse_args()
    #print options    
    if len(args) != 0:
        parser.error("incorrect number of arguments")
    
    filename = options.filename
    output = options.output
    
    xls = open_workbook(filename)
    xls_sheet = xls.sheet_by_index(options.sheet_index)
    
    labels = xls_sheet.row(0)
    # the xml wants hyphen
    labels = [s.value.replace('_', '-') for s in labels] 
    
    #print labels
    
    data = []
    XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
    XSD_NS = "http://www.w3.org/2001/XMLSchema"
    documents = etree.Element('documents', nsmap = {'xsi': XSI_NS,
                                                    'xsd': XSD_NS})
    documents.attrib['{%s}noNamespaceSchemaLocation' % XSI_NS] = 'http://www.bepress.com/document-import.xsd'
    print "Found %s rows in %s." % (xls_sheet.nrows, filename) 
    for row_index in range(1, xls_sheet.nrows):
        # be careful with order in the output xml. schema validates extremely
        # tightly.
        row = xls_sheet.row(row_index)
        record = {}
        for (i, label) in enumerate(labels):
            # force floats DATE and NUMBER into a string
            if row[i].ctype in [2,4]:
                value = u"%d" % row[i].value
            elif row[i].ctype in [3]:
                value = xldate_as_tuple(row[i].value, xls.datemode)
            else:
                value = row[i].value
            record[label] = value
        data.append(record)
        document = etree.SubElement(documents, 'document')         
            
        update_text('title', document, record)
        # TODO Handle Seasons: Override if there is a season. Map each season
        # to a month date
        
        record['publication-date'] = '%04s-%02s-%s' % (record['year'],
                             record['month'],
                             '01')
        publication_date = update_text('publication-date', document, record)        
        season = update_text('season', document, record)
    
        authors = etree.SubElement(document, 'authors')
        i = 1
        while record.get('author%s-lname' % i, None):
            author_n = 'author%s-' % i
    
            author = etree.SubElement(authors, 'author')         
            author.attrib['{%s}type' % XSI_NS] = 'individual'
    #        
    #        email = etree.SubElement(author, 'email')
    #        institution = etree.SubElement(author, 'institution')
            lname = etree.SubElement(author, 'lname')
            lname.text = record[author_n + 'lname']
            
            fname = etree.SubElement(author, 'fname')
            fname.text = record[author_n + 'fname']
            
            mname = etree.SubElement(author, 'mname')
            mname.text = record[author_n + 'mname']
            
            mname = etree.SubElement(author, 'suffix')
            mname.text = record[author_n + 'suffix']      
            i+=1      
        
        disciplines = etree.SubElement(document, 'disciplines')
        for d in record['disciplines'].split(', '):
            discipline = etree.SubElement(disciplines, 'discipline')
            discipline.text = "%s" % d        
        
               
        keywords = etree.SubElement(document, 'keywords')
        for kw in record['keywords'].split(', '):
            keyword = etree.SubElement(keywords, 'keyword')
            keyword.text = "%s" % kw
        
        abstract = etree.SubElement(document, 'abstract')        
        for text in record['abstract'].split('\n'):
            p = etree.SubElement(abstract, 'p')
            p.text = "%s" % text
        if record['fpage']:
            fpage = update_text('fpage', document, record)
        if record['lpage']:        
            lpage = update_text('lpage', document, record)
    
        update_text('fulltext-url', document, record)
        update_text('document-type', document, record)                
        issue = etree.SubElement(document, 'issue')
        
        
        issue.text = "%s/vol%s/iss%s" % (options.journal, record['volume'], 
                                         record['issue'])
                                  
    documents = etree.ElementTree(documents)
    
    xml_file = open(output, 'wb')
    #print etree.tostring(documents, encoding='utf-8', xml_declaration=True, pretty_print=True)
    documents.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)
    
    print "Wrote: %s records." % len(data) 

if __name__ == "__main__":
    main()