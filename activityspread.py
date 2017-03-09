import gspread
import gspread.utils
from oauth2client.service_account import ServiceAccountCredentials


class SpreadsheetHandler:
    def __init__(self):
        self.cra = ['cq','cp','cvb','cvel']
        self.scope = ['https://spreadsheets.google.com/feeds/']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('ActivitySheet-17b2f15a0207.json', self.scope)
        self.name_dict = {'oceaniaCarryQueue': 'link1',
                          'oceaniaGuildActivity': 'link2',
                          'activity_sheet': 'link3'
                          }

    def open_spreadsheet(self, spreadsheet, worksheet=0):
        return gspread.authorize(self.credentials).open_by_url(self.name_dict[spreadsheet]).get_worksheet(worksheet)

    @staticmethod
    def write_to_spreadsheet(spreadsheet, row, column, input_data):
        spreadsheet.update_cell(row, column, input_data)

    @staticmethod
    def get_column_values(spreadsheet, column):
        return list(filter(None, spreadsheet.col_values(column)))

    @staticmethod
    def get_column_values_raw(spreadsheet, column):
        return spreadsheet.col_values(column)

    @staticmethod
    def get_value(spreadsheet, row, column):
        return spreadsheet.cell(row, column).value

    def get_row(self, spreadsheet, column, index):
        try:
            return self.get_column_values(spreadsheet, column).index(index)
        except ValueError:
            return None
