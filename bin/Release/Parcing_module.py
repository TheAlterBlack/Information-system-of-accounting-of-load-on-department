import camelot
print("Read PDF...")
tables=camelot.read_pdf('ITP.pdf', pages='all')
print("Read success! Go to parse...")
for index in range(len(tables)):
    tables[index].to_excel("ITP{0}.xlsx".format(index), encoding='utf-8', index=False, header=False)
    print('In progress...')
print('Parcing success!')
input("Press Enter to continue!")
