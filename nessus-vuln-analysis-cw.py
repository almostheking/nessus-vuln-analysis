from lxml import etree

f = open('scanreport.nessus', 'r')
xml_content = f.read()
f.close()

parz = etree.XMLParser(huge_tree=True)
root = etree.fromstring(text=xml_content, parser=parz)
for block in root:
