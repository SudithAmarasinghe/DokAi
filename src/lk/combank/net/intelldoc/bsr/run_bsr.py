import sys
from pathlib import Path

wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))

from final import FinalBSR

def main():
    file1 = "bsr/excel_files/balance_sheet.xlsx"
    file2 = "bsr/excel_files/income_statement.xlsx"

    bsr = FinalBSR(file1, file2)
    bsr.run(file1, file2)

if __name__ == "__main__":
    main()