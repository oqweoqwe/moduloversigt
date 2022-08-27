import openpyxl
from datetime import datetime
import time

# flerklassehold: 2cf Mu, 2fi MA, 3c4i EN 1
# valgfag

holdrapport = "data/Samlet holdrapport.xlsx"
modulfordeling = "data/Samlet modulfordeling.xlsx"

classes = ["2a","2b","2c","2d","2e","2f","2h","2i","3a","3b","3c","3d","3e","3f","3h","3i","4i"]
subjects = {"da":"dansk","dr":"drama", "bi":"biologi", "en":"engelsk","hi":"historie","id":"idræt","ma":"matematik","mu":"musik","sa":"samfundsfag",
			"fy":"fysik", "bt":"biotek","ke":"kemi","ap":"ap","apla":"apla","ps":"psykologi","la":"latin","bk":"billedkunst","fi":"filosofi",
			"ty":"tysk","fr":"fransk","sp":"spansk","ng":"naturgeografi","nv":"nv","ol":"oldtidskundskab","re":"religion"}

# load modulfordeling?
load_big = True

class Main:

	def __init__(self):

		self.teams = []
		self.result_path = "Oversigt.xlsx"
		self.result = openpyxl.Workbook()
		self.result_sheet = self.result.active

		# load files into memory
		self.load_files()

		# get all teams from team report, filtering out unimportant ones
		self.load_teams()

		# load the correct amount of modules for each team from the plan
		self.load_correct_amounts()

		# write
		self.write()

	def write(self):

		error = []

		before = time.time()
		self.log(f"Started writing to {self.result_path}")

		index = 2
		self.result_sheet["A1"].value = "Hold"
		self.result_sheet["B1"].value = "Planlagt/afholdt"
		self.result_sheet["C1"].value = "burde have"
		
		for team in self.teams:
			self.result_sheet[f"A{index}"].value = team.name
			self.result_sheet[f"B{index}"].value = team.total
			planned_amount = team.planned_amount
			self.result_sheet[f"C{index}"].value = planned_amount

			if planned_amount == -1:
				error.append(team.name)
			index = index+1

		self.result.save(self.result_path)

		self.log(f"Results saved at {self.result_path}. Writing took {round(time.time()-before,1)} seconds")
		if len(error) > 0:
			self.log(f"An error occured while parsing the following teams ({len(error)}/{self.result_sheet.max_row-1}): {error}")

	def log(self, msg):
		now = datetime.now().strftime("%H:%M:%S")
		print(f"[{now}] {msg}")

	def load_files(self):
		
		before = time.time()
		now = datetime.now().strftime("%H:%M:%S")
		self.log(f"Loading {holdrapport}...")
		self.team_report = openpyxl.load_workbook(holdrapport)
		self.team_report_sheet = self.team_report.active
		self.log(f"Load took {round(time.time()-before,1)} seconds")

		if load_big:
			before = time.time()
			self.log(f"Loading {modulfordeling}...")
			self.module_distribution = openpyxl.load_workbook(modulfordeling)
			now = datetime.now().strftime("%H:%M:%S")
			self.log(f"Load took {round(time.time()-before,1)} seconds")
		else:
			self.log(f"Running without loading {modulfordeling}")
		
		self.log(f"Using the following team report: {self.team_report_sheet['A1'].value}")

	def load_teams(self):
		
		for i in range(4,self.team_report_sheet.max_row+1-5):
			name = self.team_report_sheet[f'A{i}'].value
			name_lower = name.lower()
			total = self.team_report_sheet[f'G{i}'].value

			if self.is_name_legal(name_lower):
				self.teams.append(Team(name,total))

	def get_level(self, team_name):
		subject_id = team_name.split(" ",2)[1]

		if subject_id.islower():
			return "C"
		if subject_id.isupper():
			return "A"
		return "B"

	def set_planned_amount(self, column, occurences, cell, sheet, subject, team):

		# hvis faget står flere gange, vælg det rigtige niveau
		if occurences > 2:
			for cell in sheet['D']:
				if subject in str(cell.value).lower():
					if f" {team.level}" in str(cell.value) or f"/{team.level}" in str(cell.value):
						if team.planned_amount == -1:
							team.planned_amount = sheet[f'{column}{cell.row}'].value

		elif team.planned_amount == -1:
			team.planned_amount = sheet[f'{column}{cell.row}'].value

	# also loads subject levels
	def load_correct_amounts(self):
		for team in self.teams:
			if team.name.split(' ',1)[0] in classes:
				# team er stamklasseundervisning

				class_id = self.get_class_id(team.name)
				subject = self.get_subject_name(team.name)
				sheet_name = self.get_sheet_by_class_id(class_id)
				sheet = self.module_distribution[sheet_name]

				team.level = self.get_level(team.name)

				# håndter 2. fremmedsprog i stamklasser, eks. 2a Ty
				if subject == "tysk" or subject == "fransk":
					subject = "fremmedsprog"

				occurences = 0

				for cell in sheet['D']:
					if subject in str(cell.value).lower():
						occurences = occurences+1

				for cell in sheet['D']:
					if subject in str(cell.value).lower():

						if self.get_year_from_class_id(class_id) == 1:
							
							if not sheet[f"E{cell.row}"].value == "Undervisning":
								continue

							self.set_planned_amount('F', occurences, cell, sheet, subject, team)

							continue
						if self.get_year_from_class_id(class_id) == 2:
							
							if not sheet[f"E{cell.row}"].value == "Undervisning":
								continue

							self.set_planned_amount('H', occurences, cell, sheet, subject, team)

							continue
						if self.get_year_from_class_id(class_id) == 3:
							
							if not sheet[f"E{cell.row}"].value == "Undervisning":
								continue

							self.set_planned_amount('J', occurences, cell, sheet, subject, team)

							continue
						if self.get_year_from_class_id(class_id) == 4:
							
							if not sheet[f"E{cell.row}"].value == "Undervisning":
								continue

							self.set_planned_amount('L', occurences, cell, sheet, subject, team)

							continue
			else:
				# team er ikke stamklasseundervisning
				if not 'g' in team.name.split(" ",1)[0]:
					# flerklassehold, eks. 2cf Mu
					pass
				else:
					# valgfag
					pass

	def get_subject_name(self, team_name):
		return subjects.get(team_name.lower().split(" ",2)[1])

	def get_year_from_class_id(self, class_id):

		if int(class_id[0:4]) == datetime.today().year:
			return 1
		if int(class_id[0:4]) == datetime.today().year-1:
			return 2
		if int(class_id[0:4]) == datetime.today().year-2:
			return 3
		if int(class_id[0:4]) == datetime.today().year-3:
			return 4

	def get_sheet_by_class_id(self, class_id):

		sheets = self.module_distribution.sheetnames
		for sheet in sheets:
			if class_id in sheet:
				return sheet

	def get_class_id(self, name):

		if '4' in name:
			return str(datetime.today().year-3)+name[1].lower()
		if '3' in name:
			return str(datetime.today().year-2)+name[1].lower()
		if '2' in name:
			return str(datetime.today().year-1)+name[1].lower()
		if '1' in name:
			return str(datetime.today().year-0)+name[1].lower()

	def is_name_legal(self, name):

		if "1g" in name:
			return False
		if "blå" in name:
			return False
		if "grøn" in name:
			return False
		if "rød" in name:
			return False
		if "gul" in name:
			return False
		if "gf" in name:
			return False
		if name[0] == '1':
			return False
		if "dho" in name:
			return False
		if "ff" in name:
			return False
		if "sro" in name:
			return False
		if "ssb" in name:
			return False
		if "øh" in name:
			return False
		if "hørelære" in name:
			return False
		if "sammenspil" in name:
			return False
		if "kor" in name:
			return False
		if "klassisk" in name:
			return False
		if "rytmisk" in name:
			return False
		if "mgk" in name:
			return False
		if "projekttime" in name:
			return False
		if "kt" in name:
			return False
		if "eksamen" in name:
			return False
		return True

class Team:

	def __init__(self, name, total):

		self.name = name
		self.total = total
		self.planned_amount = -1
		self.level = None


if __name__ == '__main__':
	main = Main()