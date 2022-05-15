import typer
import xlrd
import decimal
from quiffen import Qif, Category
import quiffen

app = typer.Typer()


def icici_row2dict(row: list):
    if row[2].value == "":
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
        payee = payee.split(" ", 1)[-1]

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
def icici(
    input_fname: str = "data.xls", output_fname: str = "data.csv", header_idx: int = 12
):
    trxs = []

    book = xlrd.open_workbook(input_fname)
    sh = book.sheet_by_index(0)
    print(sh.row(header_idx))
    for rx in range(header_idx + 1, sh.nrows):
        res = icici_row2dict(sh.row(rx))
        if res is not None:
            trxs.append(res)

    print(trxs[0].keys())
    with open(output_fname, "w") as f:
        f.write(
            "SNo.,Value Date,Transaction Date,Cheque No,Transaction Remarks,Withdrawal Amount (INR),Deposit Amount (INR), Balance (INR),Payee,Category,\n"
        )
        for transaction in trxs:
            for key in transaction.keys():
                f.write(f"{transaction[key]},")
            f.write("\n")


@app.command()
def qif_stuff(
    input_fname: str = "stmt.qif",
    output_fname: str = "data.qif",
    day_first: bool = False,
):
    qif = Qif.parse(input_fname, day_first=day_first)
    acc = qif.accounts["Quiffen Default Account"]
    trs = acc.transactions["Bank"]

    new_qif = Qif()
    new_acc = quiffen.Account("Personal Bank Account", "My personal bank account")
    new_qif.add_account(new_acc)

    transfer = Category("Transfer")
    shopping = Category("Shopping")
    subscriptions = Category("Subscriptions")
    food = Category("Food")

    qif.add_category(food)
    new_qif.add_category(food)

    qif.add_category(transfer)
    new_qif.add_category(transfer)

    qif.add_category(shopping)
    new_qif.add_category(shopping)

    qif.add_category(subscriptions)
    new_qif.add_category(subscriptions)

    for tr in trs:
        tr.memo = tr.payee
        if "Zelle" in tr.payee:
            tr.payee = tr.payee.split(";")[-1].strip().title()
            tr.category = transfer
        elif "Amzn" in tr.payee.title():
            tr.payee = "Amazon"
            tr.category = shopping
        elif "hulu" in tr.payee.lower():
            tr.payee = "Hulu"
            tr.category = subscriptions
        elif "waffle lab" in tr.payee.lower():
            tr.payee = "The Waffle Lab"
            tr.category = food
        elif "mcdonald" in tr.payee.lower():
            tr.category = food
        elif "chegg" in tr.payee.lower():
            tr.payee = "Chegg"
            tr.category = subscriptions
        elif "Cosmos Pizza" in tr.payee.title():
            tr.category = food
            tr.payee = "Cosmos Pizza - Boulder"
        elif "Purchase" in tr.payee.title():
            details = tr.payee.title().split(" ")
            idx_for_purchase = details.index("Purchase")
            tr.payee = " ".join(details[: idx_for_purchase - 1])
        new_acc.add_transaction(tr, header="Bank")

    # new_qif.to_qif(output_fname)
    new_qif.to_csv(f"{output_fname}.csv")


if __name__ == "__main__":
    app()
