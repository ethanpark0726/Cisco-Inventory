import re

class Parse:
    def __init__(self, data):

        self.data = data
        self.result = []

    def getInventory(self):     
        
        print('--- Removing null values in the data')
        self.data = list(filter(None, self.data))
        
        for elem in self.data:
            
            name = dict()
            description = dict()
            pid = dict()
            serialNumber = dict()

            if elem.startswith('NAME') and elem.find('PID', 0) > 0:
                continue
            elif elem.startswith('NAME')  and elem.find('PID', 0) == -1:
                name['NAME'] = elem.split(',')[0].split(':')[-1].strip().replace('"', "")
                self.result.append(name)
                description['DESCR'] = elem.split(',')[1].split(':')[-1].strip().replace('"', "")
                self.result.append(description)

                
            elif elem.startswith('PID'):
                pid['PID'] = elem.split(',')[0].split(':')[-1].strip().replace('"', "")
                self.result.append(pid)
                serialNumber['SN'] = elem.split(',')[-1].split(':')[-1].strip().replace('"', "")
                self.result.append(serialNumber)
          
        return self.result
