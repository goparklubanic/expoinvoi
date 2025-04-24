import utils

if __name__ == '__main__':
    files = utils.list_excel_files("resources")
    for file in files:
        print(f"Processing {file}")
        buyer = utils.buyername(file)
        df = utils.process_clearance_file(file, buyer)
        cdf = utils.clean_clearance_dataframe(df)
        utils.save_dataframe_to_excel(cdf)
