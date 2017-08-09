import sys
import xlrd
import xmltodict
import csv
from collections import OrderedDict

input = sys.argv[1]

output = open(sys.argv[1] + "_flattened.csv", "w")
column = int(sys.argv[2])
dataRowStart = int(sys.argv[3])
workbook = xlrd.open_workbook(input)
sheet = workbook.sheet_by_name('Sheet1')
x = []
i = 1
tagName = ''
attrName = ''
loop = 0
count = 0

def flatten(d, parent_key='', sep='.'):
    items = []
    for k, v in d.items():
        new_key = parent_key + sep + k if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return OrderedDict(items)

def flatten_dict(d):
    def items():
        for key, value in d.items():
            if isinstance(value, dict):
                for subkey, subvalue in flatten_dict(value).items():
                    yield key + "." + subkey, subvalue
            else:
                yield key, value

    return OrderedDict(items())

# Flatten dict
printme = ''
for rownum in range(dataRowStart - 1, sheet.nrows):
    alltagsDict = dict()
    cell = sheet.cell(rownum, column - 1)
    xml_content = xmltodict.parse(cell.value)
    data = ' '.join(cell.value.split())
    flattened_xml = flatten_dict(xml_content)
    for k in flattened_xml:
        tagname = ''
        if isinstance(flattened_xml.get(k), list):
            for lst in flattened_xml.get(k):
                for t in lst:
                    printme += "KEY: " + k
                    printme += "." + t
                    tagname = k + "." + t
                    printme += "\nVALUE: " + lst.get(t) + '\n'
                    alltagsDict[tagname] = lst.get(t)
        else:
            printme += "KEY: " + k
            tagname = k
            printme += "\nVALUE: " + flattened_xml.get(k) + '\n'
            alltagsDict[tagname] = flattened_xml.get(k)
    x.append(alltagsDict)
    printme = ''

y = [[]]
counter = 1
for doc in x:
    y.append([])
    temp = y[0]
    if len(temp) < 1:
        for k in doc.keys():
            y[counter].append(doc.get(k))
            y[0].append(k)
    else:
        for tag in temp:
            found = 0
            for k in doc.keys():
                if tag == k:
                    y[counter].append(doc.get(k))
                    found = 1
            if found == 0:
                y[counter].append(doc.get(k))
                y[0].append(k)
    counter+=1


temp = ''

with open(sys.argv[1] + "_flattened.csv", "wb") as f:
    for sublist in y:
        for item in sublist:
	        temp += item + ','
	f.write(temp + '\n')
        temp = ''


'''
	listWriter = csv.DictWriter(
		open(sys.argv[1] + "_flattened.csv", "wb"),
		fieldnames=flattened_xml.keys(),
		delimiter=',',
		quotechar='|',
		quoting=csv.QUOTE_MINIMAL
	)

	listWriter.writerow(flattened_xml)

	x[i].append(data + ',')
	'''
'''
	for k,v in flattened_xml.items():
		new = 1
		for count in range(0, len(x[0])):
			if x[0][count] == k:
				new = 0
		if new == 1:
			x[0].append(k)
	for k, v in flattened_xml.items():
		for count in range(0, len(x[0])):
			if x[0][count] == k:
				print "TEST"
				x[rownum].append(v)
			else:
				x[rownum].append('')

with open(sys.argv[1] + "_flattened.csv", "wb") as f:
	for sublist in x:
	    for item in sublist:
	        f.write(item)
		f.write('\n')
'''



'''
		while count < len(x[0]) + 1:
			print x
			print len(x[0])
			print count
			for k,v in flattened_xml.items():
				if count == len(x[0]):
					x.append([])
					x[0].append(k)
					x[count].append(v)
				elif x[0][count] == k:
						x[count].append(v)
				else:
					x[count].append('')
			count += 1
		for sublist in x:
		    for item in sublist:
		        f.write(item)
			f.write('\n')
'''

'''
	for char in data:
		if char == '<':
			loop = true
		else if loop == true:
			if char != ' ':
				tagName += char
			else:
				x[0].append(tagName + ',')
				loop = false
				tagName = ''
	x.append([])
	i+=1

with open(sys.argv[1] + "_flattened.csv", "wb") as f:
	for sublist in x:
	    for item in sublist:
	        f.write(item)
		f.write('\n')

lines = file.readlines()
printMe = -1
printed = ""
transTyp = ""
counter = 0
for line in lines:
	print line
	if line[0] == '<':
		if line[1] == '/':
			printed += line.rstrip()
			if printed[1] == 'F':
				print printed, "|", transTyp
			printed = ""
			printMe *= -1
			transTyp = ""
		else:
			printMe *= -1
			printed += line.rstrip()
	elif printMe == 1:
		if line[5] == "A":
			for i in line:
				counter += 1
				if line[counter: counter + 8] == "TransTyp":
					transTyp = line[counter + 10]
		printed += line.rstrip()
		counter = 0
'''
