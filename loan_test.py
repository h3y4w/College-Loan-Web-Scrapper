



class loan (object):
        simple = { }
        compound = { }
	collegeTuition = start.collegeTuition
	rate = 0.04
	
	def __init__(self, rate):
		self.rate = rate

        def find_simple (self):
                self.simple['equation'] = 'Interest = ' + str(self.collegeTuition) + '(' + str(self.rate) + '(year)'
		
		for year in range(1,11):
			if year == 1 or year == 5 or year == 10:
				full = self.collegeTuition * self.rate * year
				self.simple[year] = str(full)

	def find_compound (self):
		self.compound['equation'] = 'Interest = ' + str(self.collegeTuition) + '(' + str(self.rate) + ')^' + 'year)'
	
		for year in range(1, 11):
			if year == 1 or year == 5 or year == 10:
				full = self.collegeTuition * ( (self.rate+1) ** year)
				self.compound[year] = str(full)



test = loan()

test.find_simple()
test.find_compound()
print test.simple_loan
print test.compound_loan
