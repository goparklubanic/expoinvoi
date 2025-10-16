import utils
import utils2
import sys
if __name__ == '__main__':
    files = utils.list_excel_files("resources")
    for file in files:
        buyer = utils.buyername(file)
        print(f"Processing {file}")
        raw_df = utils2.createDataframe(file)
        raw_df = utils2.addBuyer(raw_df, buyer)
        final_df = utils2.stripRows(raw_df)
        utils2.AppendToExcel(final_df)
