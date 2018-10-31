"""
python helper to format excel file output with column widths for each column set to accommodate
the longest content in any row, for all columns
"""
import xlrd               # necessary to read excel files
import xlsxwriter         # used to write excel files, much more control than xlrd


class XLSXAutoFitColumns():
    """
    CLASS to ingest an excel file and output a new excel file with
    column widths set to fit the longest content in each column
    so the user can easily read content when the open the output file manually
    """

    def __init__(self, inputfile):
        self.workbook = xlrd.open_workbook(inputfile)
        self.outputworkbook = xlsxwriter.Workbook(self.workbook)

    # noinspection PyMethodMayBeStatic
    def fit_column_widths(self, worksheet, max_width, n_columns):
        """
        Method to set the width parameter of each column in the excel file
        :param worksheet: worksheet we are processing
        :param max_width: maximum width of a column
        :param n_columns: number of columns with content in cells (non-null)
        :return:  worksheet object, with column widths set in excel
        """
        for i in range(n_columns):
            width = max_width[i]
            worksheet.set_column(i, i, width)

    def generate_from_input(self):
        input_columns = [len(x) for x in self.columns]
        columns = []
        for column in input_columns:
            if 'unnamed' in column:
                column_index = input_columns.index(column)
                column_letter = xlsxwriter.utility.xl_col_to_name(column_index)
                column_name = 'Unlabeled_column_with_data_in_column_%s' % column_letter
                columns.append(column_name)
            else:
                columns.append(column)

        assert len(columns) == len(input_columns)

        ws_selected = self.workbook.add_worksheet(name=self)
        bold = self.workbook.add_format({'bold': True})
        ws_selected.write_row('A1', columns, bold)

        ws_row = 1
        ws_col = 0
        rows = self.rows
        max_width = [len(x) for x in columns]

        for row in rows:
            for i, column in enumerate(input_columns):
                current_width = max_width[i]
                entry = self.row[column]

                ws_selected.write(ws_row, ws_col, entry)
                if isinstance(entry, float) or isinstance(entry, int):
                    str_len = len('{:0.2f}'.format(entry))
                else:
                    str_len = len(str(entry))
                max_width[i] = max(current_width, str_len)
                ws_col += 1
            ws_row += 1
            ws_col = 0
        self.fit_column_widths(ws_selected, max_width, len(columns))

    def save(self):
        """
        Save the workbook
        :return: workbook object
        """
        self.outputworkbook.close()


if __name__ == '__main__':
    XLSXAutoFitColumns(inputfile="sample_excel_data.xlsx")
