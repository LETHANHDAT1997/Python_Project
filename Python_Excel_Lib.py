import os
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# https://stackoverflow.com/questions/30484220/fill-cells-with-colors-using-openpyxl
# https://openpyxl.readthedocs.io/en/stable/styles.html#cell-styles-and-named-styles
# https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/borders.html#Side

class Excel_WorkBook():
    
    def __init__(self,str_path_file_excel,str_name_sheet="Sheet"):
        try:
            # Kiểm tra xem file đã tồn tại chưa
            if os.path.exists(str_path_file_excel):
                # Mở workbook hiện có
                self.Ob_workbook = openpyxl.load_workbook(str_path_file_excel)
                # # Kiểm tra xem tên sheet đã tồn tại chưa
                if self.__check_name_sheet__(str_name_sheet) == True:
                    print(f"Sheet với tên '{str_name_sheet}' đã tồn tại.") 
                else:
                    # Tạo sheet mới với tên str_name_sheet
                    self.Ob_workbook.create_sheet(title=str_name_sheet)
                    print(f"Đã tạo sheet mới với tên '{str_name_sheet}'.")
                print("Đã mở File Excel.")
            else:
                # Tạo workbook mới
                self.Ob_workbook = openpyxl.Workbook()
                # Tạo sheet mới với tên str_name_sheet
                self.Ob_workbook.create_sheet(title=str_name_sheet)
                print(f"File chưa tồn tại. Đã tạo file mới với tên Sheet là '{str_name_sheet}'.")
        except Exception as e:
            print("Đã xảy ra lỗi:", e)
            return None           


    def __check_name_sheet__(self,str_name_sheet):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if str_name_sheet in self.Ob_workbook.sheetnames:
            return True
        else:
            return False


    def Create_Sheet(self,str_name_sheet, over_write=False):
        # Kiểm tra xem có cho ghi đè hay không
        if over_write == False:
            # Kiểm tra xem tên sheet đã tồn tại chưa
            if self.__check_name_sheet__(str_name_sheet) == True:
                print(f"Sheet với tên '{str_name_sheet}' đã tồn tại.")
            else:
                # Tạo sheet mới với tên 'str_name_sheet'
                self.Ob_workbook.create_sheet(title=str_name_sheet)
                print(f"Đã tạo sheet mới với tên '{str_name_sheet}'.")
        else:
            # Tạo sheet mới với tên 'str_name_sheet'
            self.Ob_workbook.create_sheet(title=str_name_sheet)


    def Return_Sheet(self,str_name_sheet):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            return self.Ob_workbook[str_name_sheet]
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")


    def Return_List_Sheet(self):
        return self.Ob_workbook.sheetnames


    def Write_strColumn(self,str_name_sheet,str_name_Cloumn,list_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            for index, data in enumerate(list_content, start=1): 
                self.Ob_workbook[str_name_sheet][f'{str_name_Cloumn}{index}'] = data
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")   


    def Write_intColumn(self,str_name_sheet,Cloumn,list_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            for index, data in enumerate(list_content, start=1): 
                self.Ob_workbook[str_name_sheet].cell(index,Cloumn,data)
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")       


    def Write_intRow(self,str_name_sheet,Row,list_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            for index, data in enumerate(list_content, start=1): 
                self.Ob_workbook[str_name_sheet].cell(Row,index,data)
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")


    def Write_strCell(self,str_name_sheet,str_cell_index,str_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            self.Ob_workbook[str_name_sheet][str_cell_index.strip()] = str_content
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")


    def Write_intCell(self,str_name_sheet,row,column,str_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            self.Ob_workbook[str_name_sheet].cell(row,column,str_content)
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")


    def Read_strCell(self,str_name_sheet,str_cell_index):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            return self.Ob_workbook[str_name_sheet][str_cell_index.strip()].value
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")  


    def Read_intCell(self,str_name_sheet,row,column,str_content):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            return self.Ob_workbook[str_name_sheet].cell(row,column).value
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")     


    def Add_Sort_Filter(self,str_name_sheet,str_column):
        self.Ob_workbook.active.auto_filter.ref = str_column
        # self.Ob_workbook.active.auto_filter.add_filter_column(0)
        # self.Ob_workbook.active.auto_filter.add_sort_condition(str_column)
        # print(self.Ob_workbook[str_name_sheet].max_row)


    def Format_Cells(self,str_name_sheet,str_cell_index,PatternFill,Font,Border):
        # Kiểm tra xem tên sheet đã tồn tại chưa
        if self.__check_name_sheet__(str_name_sheet) == True:
            self.Ob_workbook[str_name_sheet][str_cell_index.strip()].fill    = PatternFill
            self.Ob_workbook[str_name_sheet][str_cell_index.strip()].font    = Font
            self.Ob_workbook[str_name_sheet][str_cell_index.strip()].border  = Border
        else:
            print(f"Sheet với tên '{str_name_sheet}' không tồn tại.")   


    def Merge_Cells(self,str_name_sheet,str_cell_index):
        self.Ob_workbook.active.merge_cells(str_cell_index) 

    def Save(self,path_save):
        return self.Ob_workbook.save(path_save)
    

    def Close(self):
        return self.Ob_workbook.close()


if __name__ == "__main__":
    excel_file_path = "example.xlsx"  # Đường dẫn tới file Excel
    excel_sheet_name = "Sheet"       # Tên của sheet trong file Excel

    PatternFill_Base    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type = "solid")
    Font_Base           = Font(name='Tahoma',size=11,bold=True,italic=False,vertAlign=None,underline='none',strike=False,color='FF000000')
    Border_Base         = Border(  left=Side(border_style="double",
                                    color='FF000000'),
                                    right=Side(border_style="double",
                                    color='FF000000'),
                                    top=Side(border_style="double",
                                    color='FF000000'),
                                    bottom=Side(border_style="double",
                                    color='FF000000'),
                                    diagonal=Side(border_style="double",
                                    color='FF000000'),
                                    diagonal_direction=0,
                                    outline=Side(border_style="double",
                                    color='FF000000'),
                                    vertical=Side(border_style="double",
                                    color='FF000000'),
                                    horizontal=Side(border_style="double",
                                    color='FF000000')
                                )
    
    File_example = Excel_WorkBook(excel_file_path,excel_sheet_name)
    File_example.Write_strCell(excel_sheet_name,"B7",25000)
    File_example.Write_strCell(excel_sheet_name,"C7",23000)
    File_example.Add_Sort_Filter(excel_sheet_name,"A1:D7")
    File_example.Merge_Cells(excel_sheet_name,'B2:D4')
    File_example.Format_Cells(excel_sheet_name,'B7',PatternFill_Base,Font_Base,Border_Base)
    # Phải có lệnh Save thì nội dung thay đổi mới được lưu lại trên bộ nhớ
    File_example.Save(excel_file_path)
