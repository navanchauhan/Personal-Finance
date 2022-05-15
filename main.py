import typer
import xlrd

app = typer.Typer()

def icici_row2dict(row: list):
	if row[2].value == '':
		return None
	ctgry = ""
	payee = ""

	split_remarks = row[5].value.strip().split("/")
	if len(split_remarks) > 3:
		if split_remarks[0] == "MSI" or split_remarks[0] == "MIN":
			payee = split_remarks[1].strip() 
		elif split_remarks[0] == "MMT":
			payee = split_remarks[3]
			if " to " in payee.lower():
				payee = payee.split(" to ")[-1]
			else:
				payee = split_remarks[-2]

	if "RAZ" in payee:
		payee = payee.split(" ",1)[-1]

	return {
		"sno": row[1].value,
		"vdate": row[2].value,
		"tdate": row[3].value,
		"cno": row[4].value,
		"remarks": row[5].value.strip(),
		"wamt": row[6].value,
		"damt": row[7].value,
		"bal": row[8].value,
		"payee": payee.strip(),
		"ctgry": ctgry.strip(),
	}

@app.command()
def hello(input_fname: str = "data.xls", output_fname: str = "data.csv", header_idx: int = 12):
	trxs = []

	book = xlrd.open_workbook(input_fname)
	sh = book.sheet_by_index(0)
	print(sh.row(header_idx))
	for rx in range(header_idx+1, sh.nrows):
		res = icici_row2dict(sh.row(rx))
		if res is not None:
			trxs.append(res)

	print(trxs[0].keys())
	with open(output_fname,"w") as f:
		f.write("SNo.,Value Date,Transaction Date,Cheque No,Transaction Remarks,Withdrawal Amount (INR),Deposit Amount (INR), Balance (INR),Payee,Category,\n")
		for transaction in trxs:
			for key in transaction.keys():
				f.write(f"{transaction[key]},")
			f.write("\n")


@app.command()
def goodbye(name: str, formal: bool = False):
	if formal:
		typer.echo(f"Goodbye Ms. {name}")
	else:
		typer.echo(f"Bye {name}")

if __name__ == "__main__":
	app()