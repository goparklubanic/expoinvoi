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
        # print(final_df.head(15))
        utils2.AppendToExcel(final_df)
        # utils.fillempty(file)
        # df = utils.process_clearance_file(file, buyer)
        # cdf = utils.clean_clearance_dataframe(df)
        # utils.save_dataframe_to_excel(cdf)
