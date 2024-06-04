import os
import openpyxl

def read_excel_columns(file_path, sheet_name):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]

        data = []
        row = 1
        while True:
            value_A = sheet.cell(row=row, column=1).value
            value_B = sheet.cell(row=row, column=2).value
            if value_A is None and value_B is None:
                break
            data.append((value_A, value_B))
            row += 1

        workbook.close()
        return data
    except Exception as e:
        print("Đã xảy ra lỗi:", e)
        return None


class Excel_WorkBook():
    
    def __init__(self,str_path_file_excel,str_name_sheet="Sheet1"):
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


    def Save(self,path_save):
        return self.Ob_workbook.save(path_save)
    

    def Close(self):
        return self.Ob_workbook.close()


if __name__ == "__main__":
    excel_file_path = "example.xlsx"  # Đường dẫn tới file Excel
    excel_sheet_name = "Sheet1"       # Tên của sheet trong file Excel

    File_example = Excel_WorkBook(excel_file_path,excel_sheet_name)
    File_example.Write_strCell(excel_sheet_name,"C5","LETHANHDAT10051997")
    File_example.Write_strCell(excel_sheet_name,"C10","LETHANHDAT10051997")
    # Phải có lệnh Save thì nội dung thay đổi mới được lưu lại trên bộ nhớ
    File_example.Save(excel_file_path)
