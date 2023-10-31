import sys
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QDialog, QMessageBox, QTextEdit
from PyQt6.QtGui import QStandardItem, QStandardItemModel
from PyQt6 import uic
from PyQt6.QtGui import QIcon

import pandas as pd
import numpy as np

class ChangeDataSourceWindow(QDialog):
    '''window for changing data source'''
    def __init__(self):
        super().__init__()
        # load ui:
        uic.loadUi('Change_data_source.ui', self)
        self.setWindowIcon(QIcon("report_icon.png"))
        self.setWindowTitle("Change data source")

class MainWindow(QDialog):
    '''main window class'''
    def __init__(self):
        super().__init__()

        # Load the UI file
        uic.loadUi('Weekly.ui', self)
        self.setWindowIcon(QIcon("report_icon.png"))
        self.setWindowTitle("Update Weekly Report")

        # create an instance of the Change data source window:
        self.change_data_source_window = ChangeDataSourceWindow()

        # Change data source functions: Initialize a flag to track if the connection has been established
        self.delete_columns_connection_established = False
        self.create_new_column_connection_established = False
        self.rename_column_connection_established = False
        self.delete_report_specific_connection_established = False
        self.drop_all_rows_connection_established = False
        # Flag for main window:
        self.check_new_reports_connection_established = False
        self.update_report_connection_established = False
        self.change_data_source_connection_established = False

        # Connect button signals, if connections is not established
        if not self.check_new_reports_connection_established:
            self.checkNewReports.clicked.connect(self.check_new_reports)
            self.check_new_reports_connection_established = True
        if not self.update_report_connection_established:
            self.UpdateReport.clicked.connect(self.update_report)
            self.update_report_connection_established = True
        if not self.change_data_source_connection_established:
            self.ChangeDataSourceButton.clicked.connect(self.change_data_source)
            self.change_data_source_connection_established = True

    def message_box_show(self, message, title):
        '''function to show message box'''
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec()

    def read_excel_with_message(self, report_name):
        '''function to read excel and rgive message if it is not closed'''
        try:
            df = pd.read_excel(f'./data/source/{report_name}.xlsx')
            df['report_number'] = report_name
        except PermissionError:
            message = f'Не могу прочитать файл: {report_name}.xlsx. Закройте файл в Excel.'
            self.message_box_show(message, title='Access Error')
            df = None
        return df

    def delete_columns(self):
        '''function that deletes columns from dataframe'''
        # read csv file:
        df = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        # Read text from textEdit widget from UI file
        delete_text_edit = self.change_data_source_window.findChild(QTextEdit, 'textEditDelete')
        delete_text = delete_text_edit.toPlainText()
        print(f'delete_text: {delete_text}, len delete_text: {len(delete_text)}')
        if len(delete_text) > 0:
            delete_columns = [delete_text]
            test_column = delete_text
            if test_column in list(df.columns):
                df = df.drop(columns=delete_columns)
                # successful delete column message:
                self.message_box_show(message=f'Колонка: {test_column} удалена из источника', title='Column deleted')
                # save changes to file:
                df.to_csv('./data/result/df_result.csv', index=False, sep='|')
            else:
                # message: column not in data source
                self.message_box_show(message=f'Колонки: {test_column} нет в источнике', title='No column')
        else:
            # message: empty input
            self.message_box_show(message='Пустое поле', title='Empty input')

    def create_new_column(self):
        '''function to create empty column in data source'''
        # read csv file:
        df = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        # Read text from textEdit widget from UI file
        new_column_text_edit = self.change_data_source_window.findChild(QTextEdit, 'textEditNewColumn')
        new_column = new_column_text_edit.toPlainText()
        print(f'new_column: {new_column}, len new_column: {len(new_column)}')
        if len(new_column) > 0:
            # add empty column to data source:
            if new_column not in list(df.columns):
                df[new_column] = np.nan
                # successful add column message:
                self.message_box_show(message=f'Колонка: {new_column} добавлена в источник', title='Column added')
                # save changes to file:
                df.to_csv('./data/result/df_result.csv', index=False, sep='|')
            else:
                # message: column already in data source
                self.message_box_show(message=f'Колонка: {new_column} уже есть в источнике', title='Already Exist')
        else:
            # message: empty input
            self.message_box_show(message='Пустое поле', title='Empty input')

    def rename_column(self):
        '''function to rename existing column in data source'''
        # read csv file:
        df = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        # Read text from textEdit widget from UI file
        rename_text_edit = self.change_data_source_window.findChild(QTextEdit, 'textEditRename')
        rename_column_text = rename_text_edit.toPlainText()
        print(f'rename_column_text: {rename_column_text}, len rename_column_text: {len(rename_column_text)}')
        if len(rename_column_text) > 0:
            # parse columns:
            rename_columns_list = rename_column_text.split('-')
            if len(rename_columns_list) != 2:
                # message: wrong input format
                self.message_box_show(message='Неверный формат ввода. Ожидается: Старая колонка-Новая колонка (разделенные знаком - )', title='Wrong Input Format')
            else:
                print(f'rename_column_list: {rename_columns_list}')
                old_column = rename_columns_list[0]
                new_column = rename_columns_list[1]
                if old_column not in list(df.columns):
                    # message: column not in df:
                    self.message_box_show(message=f'Колонка: {old_column} не найдена в источнике', title='Not found')
                else:
                    print('rename flow')
                    df = df.rename(columns={old_column: new_column})
                    # save to data source:
                    df.to_csv('./data/result/df_result.csv', index=False, sep='|')
                    # message: successful rename
                    self.message_box_show(message=f'Переименовали: {old_column} в {new_column}', title='Rename successful')
        else:
            # message: empty input
            self.message_box_show(message='Пустое поле', title='Empty input')

    def delete_report_specific(self):
        '''function to delete one specific report
        from data source'''
        # read csv file:
        df = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        # convert report to str:
        df['report_number'] = df['report_number'].astype(str)
        # Read text from textEdit widget from UI file
        delete_report_text_edit = self.change_data_source_window.findChild(QTextEdit, 'textEditDeleteReport')
        delete_report_text = delete_report_text_edit.toPlainText()
        print(f'delete_report_text: {delete_report_text}, len delete_report_text: {len(delete_report_text)}')
        if len(delete_report_text) > 0:
            # delete report from data source:
            # unique set of reports that are already in data source:
            reports_data_list = list(df['report_number'].unique())
            # check if report is in data source:
            if delete_report_text in reports_data_list:
                # delete report
                df = df[df['report_number'] != delete_report_text]
                df.to_csv('./data/result/df_result.csv', index=False, sep='|')
                self.message_box_show(message=f'Отчет: {delete_report_text} удален из источника', title='Success')
            else:
                # message that report is not in data source:
                self.message_box_show(message=f'Вы ввели {delete_report_text}, но такого отчета нет в Источнике данных, проверьте входные данные', title='Wrong input')
        else:
            # message: empty input
            self.message_box_show(message='Пустое поле', title='Empty input')

    def drop_all_rows(self):
        '''function to drop all rows in data source'''
        # Create a confirmation dialog
        confirmation = QMessageBox()
        confirmation.setIcon(QMessageBox.Icon.Warning)
        confirmation.setWindowTitle("Confirmation")
        confirmation.setText("Уверен, что нужно прям удалить все строки?")
        confirmation.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        # Check user's choice
        result = confirmation.exec()
        if result == QMessageBox.StandardButton.Ok:
            # User clicked "OK," so proceed with dropping rows
            print("Dropping all rows...")
            # read csv file:
            df = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
            df = df.drop(df.index)  # Remove all rows
            # save to data source:
            df.to_csv('./data/result/df_result.csv', index=False, sep='|')
            # message: All rows dropped
            self.message_box_show(message='Все данные удалены из источника', title='Dropped')
        else:
            # User clicked "Cancel," do nothing or provide feedback
            print("Cancelled")


    def change_data_source(self):
        # create new window ChangeDataSource class:
        self.change_data_source_window.show()
        # check if connection has already been established and if not - create new connection
        if not self.delete_columns_connection_established:
            self.change_data_source_window.DeleteColumnsButton.clicked.connect(self.delete_columns)
            self.delete_columns_connection_established = True
        if not self.create_new_column_connection_established:
            self.change_data_source_window.CreateNewColumnButton.clicked.connect(self.create_new_column)
            self.create_new_column_connection_established = True
        if not self.rename_column_connection_established:
            self.change_data_source_window.RenameColumnButton.clicked.connect(self.rename_column)
            self.rename_column_connection_established = True
        if not self.delete_report_specific_connection_established:
            self.change_data_source_window.DeleteReportButton.clicked.connect(self.delete_report_specific)
            self.delete_report_specific_connection_established = True
        if not self.drop_all_rows_connection_established:
            self.change_data_source_window.DropAllRowsButton.clicked.connect(self.drop_all_rows)
            self.drop_all_rows_connection_established = True

    def check_what_reports_are_new(self, reports_uploaded):
        '''function to find new reports in source folder'''
        folder_path = './data/source/'
        folder = Path(folder_path)
        new_reports = []
        for file in folder.iterdir():
            if file.is_file():
                file_name = file.name
                report_name = file_name.replace('.xlsx', '')
                if report_name not in reports_uploaded:
                    new_reports.append(report_name)
        return new_reports


    def check_new_reports(self):
        '''function to display which reports
            are already uploaded
            and what reports are new'''
        # read from csv:
        df_result = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        df_result['report_number'] = df_result['report_number'].astype(str)
        reports_uploaded = sorted(list(df_result['report_number'].unique()))
        # New reports list:
        new_reports = self.check_what_reports_are_new(reports_uploaded)

        # Reports available: create list model and display in list view:
        model = QStandardItemModel()
        for item_text in reports_uploaded:
            item_text = str(item_text)
            item = QStandardItem(item_text)
            item.setText(item_text)
            model.appendRow(item)
        self.listViewRead.setModel(model)
        # New reports:
        model = QStandardItemModel()
        for item_text in new_reports:
            item_text = str(item_text)
            item = QStandardItem(item_text)
            item.setText(item_text)
            model.appendRow(item)
        self.listViewNewReports.setModel(model)

    def check_columns(self, df_1, df_2):
        '''function that checks
        what columns are not common for 2 dataframes'''
        # all columns for both dfs:
        main_columns = list(df_1.columns)
        new_columns = list(df_2.columns)
        # lists for absent columns:
        main_columns_absent = []
        new_columns_absent = []
        print(f'main_columns_absent: {main_columns_absent}, new_columns_absent: {new_columns_absent}')
        # filling list of missing columns:
        for main_column in main_columns:
            if main_column not in new_columns:
                main_columns_absent.append(main_column)
        for new_column in new_columns:
            if new_column not in main_columns:
                new_columns_absent.append(new_column)
        print(f'main_columns_absent: {main_columns_absent}, new_columns_absent: {new_columns_absent}')
        return main_columns_absent, new_columns_absent

    def update_report(self):
        '''function that concat new reports
        to the main datasource'''
        # reports to add list:
        reports_to_add = []
        reports_to_add_names = []
        # read from csv:
        df_result = pd.read_csv('./data/result/df_result.csv', sep='|', low_memory=False)
        df_result['report_number'] = df_result['report_number'].astype(str)
        reports_uploaded = sorted(list(df_result['report_number'].unique()))
        # New reports list:
        new_reports = self.check_what_reports_are_new(reports_uploaded)
        if len(new_reports) == 0:
            message_1 = f'Новых файлов нет'
            self.textEditColumns.append(message_1)
        else: # add new reports
            if len(df_result) == 0: # if data source is empty
                print('empty data source flow')
                first_report = new_reports[0]
                df_result = self.read_excel_with_message(first_report)
                message_1 = f'Отчет: {first_report}'
                self.textEditColumns.append(message_1)
                reports_to_add_names.append(first_report)
                reports_uploaded = sorted(list(df_result['report_number'].unique()))
                new_reports = self.check_what_reports_are_new(reports_uploaded)
            for new_report_name in new_reports:
                message_1 = f'Отчет: {new_report_name}'
                self.textEditColumns.append(message_1)
                df_new = self.read_excel_with_message(new_report_name)
                if df_new is not None:
                    main_columns_absent, new_columns_absent = self.check_columns(df_result, df_new)
                    print(f'main_columns_absent: {main_columns_absent}, new_columns_absent: {new_columns_absent}')
                else:
                    break
                # add new report to concat list:
                if len(main_columns_absent) == 0 and len(new_columns_absent) == 0:
                    message_1 = f'Отчет: {new_report_name} - Все Ок набор колонок совпадает'
                    self.textEditColumns.append(message_1)
                    print(f'Отчет: {new_report_name} Все Ок набор колонок совпадает')
                    reports_to_add.append(df_new)
                    reports_to_add_names.append(new_report_name)
                else:
                    if len(main_columns_absent) > 0 and len(new_columns_absent) > 0:
                        message_1 = f'В новом файле: {new_report_name} нет колонок: {main_columns_absent}'
                        message_2 = f'В новом файле: {new_report_name} появились новые колонки: {new_columns_absent}'
                        self.textEditColumns.append(message_1)
                        self.textEditColumns.append(message_2)
                        # code for changes: maybe rename
                        self.message_box_show(message=f'Отчет: {new_report_name} Проблема с колонками. Посмотрите Log и воспользуйтесь функционалом изменения источника данных', title='Report Columns Problem')
                        break
                    elif len(main_columns_absent) > 0:
                        message_1 = f'OK, но в новом файле: {new_report_name} нет колонок: {main_columns_absent}'
                        self.textEditColumns.append(message_1)
                        reports_to_add.append(df_new)
                        reports_to_add_names.append(new_report_name)
                    else:
                        message_1 = f'OK, но в новом файле: {new_report_name} появились новые колонки: {new_columns_absent}'
                        self.textEditColumns.append(message_1)
                        reports_to_add.append(df_new)
                        reports_to_add_names.append(new_report_name)
            # save result to the memory
            if len(reports_to_add) > 0:
                # concat reports to add:
                df_to_add = pd.concat(reports_to_add, ignore_index=True)
                # concat to data source:
                df_result = pd.concat([df_result, df_to_add], ignore_index=True)
                df_result.to_csv('./data/result/df_result.csv', index=False, sep='|')
                message_final = f'Отчеты, добавленные в источник: {reports_to_add_names}'
                self.textEditColumns.append(message_final)
        

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
