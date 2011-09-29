import jarvis
import re
import unittest

class get_name_test(unittest.TestCase):
    first_name = 'John'
    last_name = 'Doe'
    title = 'Mr.'
    formal_address = 'sir'
    
    def test_get_name_sanity(self):
        """get_name should return a combination of the parameters or empty string"""
        name = jarvis.get_name(self.first_name, self.last_name, self.title, self.formal_address)
        pattern = \
            '^(' + \
            self.title + ' ' + self.last_name + '|' + \
            self.first_name + '|' + \
            self.formal_address + '|' + \
            '' + \
            ')$'
        result = re.search(pattern, name)
        self.assertTrue(result)
        
    def test_get_name_all_results(self):
        """get_name should return every possible combination"""
        flag = [False, False, False, False]
        for i in range(1,100):
            name = jarvis.get_name(self.first_name, self.last_name, self.title, self.formal_address)
            if (re.search(self.title + ' ' + self.last_name, name)):
                flag[0] = True
            elif (re.search(self.first_name, name)):
                flag[1] = True
            elif (re.search(self.formal_address, name)):
                flag[2] = True
            elif (re.search('', name)):
                flag[3] = True
        self.assertTrue(flag)