import time
import os
import xlwings as xw
import datetime

SHEET_DATETIME_DICT = {"0221": "C", "0228": "D", "0307": "E", "0314": "F", "0321": "G", "0328": "H", "0402": "I", "0411": "J", "0418": "K", "0424": "L", "0425": "M", "0509": "N", "0516": "O", "0523": "P", "0530": "Q", "0606": "R", "0613": "S"}

# def string2datetime(string):
#     return datetime.datetime.strptime(string, "%Y-%m-%d %H:%M:%S")
def getFileListBySuffix(input_folder, suffix):
    raw_files = []
    # 读input_path下所有的suffix文件名
    for file in os.listdir(input_folder):
        file_path = os.path.join(file)
        if os.path.splitext(file_path)[1] == suffix:
            raw_files.append(input_folder + file_path)
    return raw_files

def process_batch(xlsx_canditates_folder_path, target_xlsx_path):
    xlsx_files = getFileListBySuffix(xlsx_canditates_folder_path, suffix=".xlsx")
    for file in xlsx_files:
        print(file)
        datetimeNumber = file.split("计算语言学")[1].split(".")[0]
        record_students(wb_source_path=file, wb_target_path=target_xlsx_path, datetimeNumber=datetimeNumber)

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

def record_students(wb_source_path, wb_target_path, datetimeNumber):
    # 0221就是表C
    flag = SHEET_DATETIME_DICT[datetimeNumber]
    wb_source = xw.Book(wb_source_path)
    wb_target = xw.Book(wb_target_path)
    sheet_source = wb_source.sheets[0]
    sheet_target = wb_target.sheets[0]
    sheet_source_length = sheet_source.used_range.last_cell.row
    sheet_target_length = sheet_target.used_range.last_cell.row
    no_col_dict = {}
    # 读输出表格，获取{名字, column_index}的dict
    target_col_start = 2
    for i in range(target_col_start, sheet_target_length + 1):
        no = sheet_target.range(f'A{i}').value
        no_col_dict[no] = i
    print(no_col_dict)
    # time.sleep(100)
    # print(sheet_target_length)
    # print(sheet_source_length)
    meeting_start_time_str = sheet_source.range('B4').value
    meeting_end_time_str = sheet_source.range('B5').value
    # print(meeting_start_time_str, type(meeting_start_time_str))
    # print(meeting_end_time_str, type(meeting_end_time_str))
    meeting_start_time_datetime = datetime.datetime.strptime(meeting_start_time_str, "%Y-%m-%d %H:%M:%S")
    meeting_end_time_datetime = datetime.datetime.strptime(meeting_end_time_str, "%Y-%m-%d %H:%M:%S")
    # print(meeting_start_time_datetime, type(meeting_start_time_datetime))
    # print(meeting_end_time_datetime, type(meeting_end_time_datetime))
    meeting_duration = meeting_end_time_datetime - meeting_start_time_datetime
    # print(meeting_duration, type(meeting_duration))
    source_col_start = 10
    for i in range(source_col_start, sheet_source_length + 1):
        # 学号、姓名
        content = sheet_source.range(f'A{i}').value
        content = content.split("(")[1:][0].split(")")[0]
        # 先看括号里有没有-
        # 如果有，分出学号、姓名
        if "-" in content:
            content_part1 = content.split("-")[0]
            content_part2 = content.split("-")[1]
            if is_number(content_part1[0]):
                student_no = content_part1.strip()
                student_name = content_part2.strip()
            else:
                student_no = content_part2
                student_name = content_part1
            print(student_no, student_name)
        # 如果没有，分出学号姓名
        else:
            for index, char in enumerate(content):
                if not is_number(char):
                    student_no = content[0:index]
                    student_name = content[index:]
                    print(student_no, student_name)
                    break
        # 入会时间
        meeting_join_time_str = sheet_source.range(f'B{i}').value
        meeting_join_time_datetime = datetime.datetime.strptime(meeting_join_time_str, "%Y-%m-%d %H:%M:%S")
        # print(f"入会时间{meeting_join_time_datetime}", type(meeting_join_time_datetime))
        # 迟到时间
        meeting_late_time = meeting_join_time_datetime - meeting_start_time_datetime
        if meeting_late_time > datetime.timedelta(0):
            meeting_late_time = str(meeting_late_time)
            result = f"迟到{meeting_late_time}"
            print(result)
        else:
            result = "✔"
            print(result)
        if student_no in no_col_dict:
            target_index = no_col_dict[student_no]
            sheet_target.range(f'{flag}{target_index}').value = result
        else:
            print(f"###{student_no}{student_name}不在上课名单上###")
    # 旷课
    for i in range(target_col_start, sheet_target_length + 1):
        state = sheet_target.range(f'{flag}{i}').value
        if state is None:
            sheet_target.range(f'{flag}{i}').value = "旷课"
    wb_source.close()

if __name__ == '__main__':
    # wb_source_path = "./xlsx_files/计算语言学0221.xlsx"
    # wb_target_path = "./选课名单.xlsx"
    # datetimeNumber = wb_source_path.split("计算语言学")[1].split(".")[0]
    # record_students(wb_source_path, wb_target_path, datetimeNumber)

    xlsx_canditates_folder_path = "./xlsx_files/"
    target_xlsx_path = "./选课名单.xlsx"
    process_batch(xlsx_canditates_folder_path, target_xlsx_path)

