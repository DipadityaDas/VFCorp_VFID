from pandas import read_excel


def vfid_data() -> None:
	try:
		excel_df = read_excel(dir_path + 'VFID.xlsx')
	except FileNotFoundError:
		print("[INFO] Please provide the Excel file.")
		exit(0)
	
	categories = {
		"SAPE": [],
		"SAP6": [],
		"SBPC": [],
		"REVA": []
	}
	
	for title in excel_df.Title:
		if "Disable" in title or "Revoke" in title:
			words: list = title.split()
			system: str = words[2]
			
			if system.__contains__("SAPE"):
				category = categories.get('SAPE')
			elif system.__contains__("SAP6"):
				category = categories.get('SAP6')
			elif system.__contains__("SBPC"):
				category = categories.get('SBPC')
			else:
				category = categories.get('REVA')
			
			if category is not None:
				category.append([words[-1], system])
	
	for category, data in categories.items():
		if data:
			print_to_file(data)
	
	print("[INFO] Successfully created the vfID.txt file.")
	print("[INFO] All data is segregated as per systems.")


def print_to_file(usage_data: list) -> None:
	with open(dir_path + 'vfID.txt', 'a') as f:
		f.write("================================================================\n")
		for row in usage_data:
			f.write(f"{row[0]}\t{row[1]}\n")


if __name__ == "__main__":
	dir_path = 'C:\\Users\\Dipaditya\\Downloads\\'
	vfid_data()
