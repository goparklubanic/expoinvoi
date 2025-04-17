import utils

if __name__ == '__main__':
    files = utils.list_excel_files("resources")
    for file in files:
        df = utils.process_clearance_file(file)
        cdf = utils.clean_clearance_dataframe(df)
        utils.save_dataframe_to_excel(cdf)
