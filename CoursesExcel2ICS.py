# 中国能源大学（筹）地铁12号线校区 课表Excel文件转.ics文件
# By 冷月おとわ

import AcademicAffairs
import datetime
import pytz
import json
import re
import pandas as pd
from icalendar import Calendar, Event

# 各大节课开始时间
startTime = [
    {},
    {},
    {'hour': 8, 'min': 0, 'text': '第一二节'},
    {'hour': 10, 'min': 10, 'text': '第三四节'},
    {'hour': 14, 'min': 0, 'text': '第五六节'},
    {'hour': 16, 'min': 10, 'text': '第七八节'},
    {'hour': 19, 'min': 0, 'text': '第九十节'}
    ]

# 课表Excel转换类
class CoursesExcel(object):
    def __init__(self): # 类进入接口
        # 获取开学时间(第1周周一),为基准时间
        # 两种开学时间获取方式: 1) API; 2) 手动输入
        print("\n请选择开学时间(第1周周一)写入方式：")
        print("1: 强智教务系统API")
        print("2: 手动输入开学时间")
        flag = int(input("请输入您想使用的写入方式的数字代号: "))
        startDate = self.startTermDate(flag)
        
        # 读取课表Excel文件，转换为DataFrame
        print("\n请输入您获取到的课表Excel文件的绝对路径: ")
        print("绝对路径示例: D:\Documents\SDUSTCourse2ICS\Courses.xls")
        path = input("路径: ")
        df = pd.read_excel(path)

        cal = Calendar() # 初始化日历
        date_now = datetime.datetime.now(tz = pytz.timezone('Asia/Shanghai')) # 获取现在本机时间，用于创建DTSTAMP时间戳

        cal = self.createCalendar(date_now, cal, df) # 构造日历

    # 获取开学时间相关
    def startTermDate(self, flag):
        if flag == 1:
            # 输入账号密码信息
            account = input("\n请输入您的山东科技大学教务账号: ")
            password = input("请输入您的山东科技大学教务账号密码: ")
            url = "http://jwgl.sdust.edu.cn/app.do" 
            Q = AcademicAffairs.SW(account, password, url)

            data_json = Q.get_current_time() # 获取当前周次信息，以JSON字符串存储
            data = json.loads(data_json) # 解析JSON字符串为List

            # 解析当前周开始日期
            start_date_str = data['s_time']  # 获取开始日期的字符串表示
            start_datetime = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")  # 将字符串转换为datetime对象

            # 计算第1周周一的日期
            weeks_to_subtract = data['zc'] - 1  # 需要回推的周数
            days_to_subtract = weeks_to_subtract * 7  # 转换为天数
            first_week_monday = start_datetime - datetime.timedelta(days = days_to_subtract)  # 计算第1周周一的日期

        else:
            print("\n请手动输入开学时间(第1周周一)的相关信息:")
            year = int(input("年: "))
            month = int(input("月: "))
            day = int(input("日: "))
            first_week_monday = datetime.datetime(year, month, day)

        # 给开始日期分配Asia/Shanghai时区，防止时间错乱
        sh_tz = pytz.timezone('Asia/Shanghai')
        first_week_monday = sh_tz.localize(first_week_monday)

        print("\n当前学期开学时间(第1周周一)为:", first_week_monday)
        return first_week_monday
    
    # 构造日历
    def createCalendar(self, date_now, cal, df):
        print("\n正在处理并转换读取到的课表信息")

        # 以列优先的方式遍历读取的Excel文件
        col = -1 # 列标

        for column in df:
            col += 1

            if col == 0: # 第0列无课表有效数据
                continue

            print("\n当前正在处理的星期: {}".format(df[column][1]))

            # 遍历当前日期的课程信息(2行至6行)
            for i in range(2, 7):
                print("\n当前正在处理的节次: {}{}".format(df[column][1], startTime[i]['text']))
                courses = df[column][i] # 该节次的原始字符串
                if courses != ' ': # 节次原始字符串非空则继续操作，空则跳过
                    courses_split = self.splitCourseInfo(courses) # 原始字符串切割形成的List

                    # 遍历切割形成的List，创建日历事件
                    for course in courses_split:
                        self.parseCourseInfo(course)

                else: # 此处原始字符串为空对应为仅有一个空格的串
                    print("{0}{1}无课,将跳过对{0}{1}的处理".format(df[column][1], startTime[i]['text']))

            print("\n{}处理完毕".format(df[column][1]))


    # 对当前节次的字符串进行初始切分，每个课程独立成串
    def splitCourseInfo(self, courses):
            # 使用连续的换行符(原串特征)作为分隔符来切分字符串
            courses_split = courses.strip().split('\n\n')
            # 确保每个课程信息块的格式是统一的，即以\n开始和结束
            # 这样处理之后，无课程的空串将格式化为'\n\n'串
            courses_split = ['\n' + course + '\n' for course in courses_split]
            return courses_split
    
    # 使用正则表达式匹配课程信息
    def parseCourseInfo(self, course_info):
        # 使用正则表达式匹配课程信息的各个部分
        match = re.search(r'\n(.+?)\n(.+?)\n(\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*)\[周\](?:\n(.*))?', course_info)

        if match: # 各group对号入座
            course_name = match.group(1)
            teacher_info = match.group(2)
            week_info = match.group(3)
            room_info = match.group(4) if match.group(4) is not None else "NULL_CLASSROOM"

            # 处理教师信息，删除括号但保留职称
            teacher_info = re.sub(r'\((.*?)\)', r'\1', teacher_info)
            print(f"课程名称: {course_name}, 教师信息: {teacher_info}, 上课周次: {week_info}, 教室信息: {room_info}")
            return course_name, teacher_info, week_info, room_info
            

