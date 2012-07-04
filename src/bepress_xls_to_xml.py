#! /usr/bin/python
#=============================================================================
# Transform a Excel 97-2003 spreadsheet to an XML file suitable for loading 
# via XML batch upload for   
#=============================================================================

from xlrd import open_workbook, xldate_as_tuple
from lxml import etree
import lxml.html as HT
from lxml.html.clean import Cleaner
import optparse
import unidecode
import urllib

def get_bepress_elements():
    xsd = etree.parse('http://www.bepress.com/document-import.xsd')
    element_names = xsd.xpath('//xsd:element[@name="documents"]//xsd:element/@name', namespaces = {'xsd': 'http://www.w3.org/2001/XMLSchema'})
    return element_names


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
                      help="Sheet index to extract from, startinf from '0'") 
    (options, args) = parser.parse_args()
    #print options
    if not options.filename:
        parser.error('-f or --filename: Need a filename!')
    elif not options.output:
        parser.error('-o or --output: Need an output file!')
    elif not options.journal:
        parser.error('-j or --journal: Need journal!')    
    
    try:
        sheet_index = int(options.sheet_index)
    except:
        parser.error('-s or --sheet: Sheet index must be an integer!')
           
    
    filename = options.filename
    output = options.output
    
    xls = open_workbook(filename)
    xls_sheet = xls.sheet_by_index(sheet_index)
    
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
            if row[i].ctype in [2]:
                value = u"%d" % row[i].value
            elif row[i].ctype in [3]:
                value = xldate_as_tuple(row[i].value, xls.datemode)
            elif row[i].ctype in [4]:
                if row[i].value:
                    value = True
                else:
                    value = False                                                                
            else:
                value = row[i].value
            record[label] = value
        data.append(record)
        document = etree.SubElement(documents, 'document')         
            
        update_text('title', document, record)
        # Seasons need to have publication dates
        if 'publication-date' in record:            
            record['publication-date'] = '%04d-%02d-%02d' % record['publication-date'][0:3]
            
        else:
            record['publication-date'] = '%04s-%s-%s' % (record['year'],
                                 record['month'].zfill(2),
                                 '01')
        update_text('publication-date', document, record)        
        
        #season
        update_text('season', document, record)        
        if 'publication-date-date-format' in record:
            # Insconsistent underscore    
            pddf = etree.SubElement(document, 'publication_date_date_format')            
            pddf.text = "%s" % record['publication-date-date-format']
            
        authors = etree.SubElement(document, 'authors')
        i = 1            
        while record.get('author%s-fname' % i, None):
            author_n = 'author%s-' % i
            author = etree.SubElement(authors, 'author')
            
            if record.get(author_n + 'is-corporate', False) == True:                
                author.attrib['{%s}type' % XSI_NS] = 'corporate'
                name = etree.SubElement(author, "name")                
                name.text = record[author_n +'fname']
            else:                
                author.attrib['{%s}type' % XSI_NS] = 'individual'
                email = etree.SubElement(author, 'email')
                email.text = record[author_n + 'email']
                institution = etree.SubElement(author, 'institution')
                institution.text = record[author_n + 'institution'] 
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
        for d in record['disciplines'].split('; '):
            discipline = etree.SubElement(disciplines, 'discipline')
            discipline.text = "%s" % d        
               
        keywords = etree.SubElement(document, 'keywords')
        for kw in record['keywords'].split(', '):
            keyword = etree.SubElement(keywords, 'keyword')
            keyword.text = "%s" % kw
        
        abstract = etree.SubElement(document, 'abstract')
        if record['abstract']:
            if '<p>' in record['abstract'] or '<div>' in record['abstract']:
                cleaner = Cleaner(allow_tags=['a', 'img', 'p', 'br', 'b', 'i', 'em', 'sub', 
                                              'sup', 'u', 'strong'],
                                  remove_unknown_tags=False)
                html = cleaner.clean_html(record['abstract'])
                if html == "<div></div>":
                    html = record['abstract']                
                for element in HT.fragments_fromstring(html):
                    abstract.append(element)
                divs = abstract.xpath('//div')
                for div in divs:
                    div.drop_tag()
                if not abstract.xpath('.//text()'):                    
                    print "Could not pick up abstract in %s" % unidecode.unidecode(record['title'])
            else:
                for text in record['abstract'].split('\n'):
                    p = etree.SubElement(abstract, 'p')            
                    p.text = "%s" % text
        if record['fpage']:
            update_text('fpage', document, record)
        if record['lpage']:        
            update_text('lpage', document, record)
        if 'fulltext-url' in record:
            record['fulltext-url'] = urllib.quote(record['fulltext-url'], '/:')
        update_text('fulltext-url', document, record)
        update_text('document-type', document, record)                
        issue = etree.SubElement(document, 'issue')        
        issue.text = "%s/vol%s/iss%s" % (options.journal, record['volume'], 
                                         record['issue'])
        # source field
        fields = etree.SubElement(document, 'fields')
        
        if 'source-citation' in record:            
            source = etree.SubElement(fields, 'field')
            source.attrib['name'] = 'source'
            source.attrib['type'] = 'string'
            value = etree.SubElement(source, 'value')
            value.text = record['source-citation']
            
        if 'publisher' in record:                        
            publisher = etree.SubElement(fields, 'field')
            publisher.attrib['name'] = 'publisher'
            publisher.attrib['type'] = 'string'
            value = etree.SubElement(publisher, 'value')
            value.text = record['publisher']
                    
    documents = etree.ElementTree(documents)
    
    xml_file = open(output, 'w')
    #print etree.tostring(documents, encoding='utf-8', xml_declaration=True, pretty_print=True)
    documents.write(xml_file, encoding='utf-8', xml_declaration=True, pretty_print=True)
    
    print "Wrote: %s records." % len(data) 

if __name__ == "__main__":
    main()
