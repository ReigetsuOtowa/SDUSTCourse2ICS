# 胶澳地区青岛特别市经济技术开发区西南辛安矿业高等学校 课表转换ics API版
# By 冷月おとわ

import AcademicAffairs
import json
import datetime
import pytz
from icalendar import Calendar, Event

# 强智教务管理系统用户信息
########################################
account = ""
password = ""
url = "http://jwgl.sdust.edu.cn/app.do" 
########################################

# 推演当前学期开学时间（以第1周周一为基准）
def startTermDate(Q):
    data_json = Q.get_current_time() # 获取当前周次信息，以JSON字符串存储

    data = json.loads(data_json) # 解析JSON字符串为字典

    # 解析当前周开始日期
    start_date_str = data['s_time']  # 获取开始日期的字符串表示
    start_datetime = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")  # 将字符串转换为datetime对象

    # 给开始日期分配Asia/Shanghai时区，防止时间错乱
    sh_tz = pytz.timezone('Asia/Shanghai')
    start_datetime = sh_tz.localize(start_datetime)

    # 计算第1周周一的日期
    # 上述JSON字符串格式示例：{"zc":5,"e_time":"2024-03-31","s_time":"2024-03-25","xnxqh":"2023-2024-2"}
    # 假设zc字段表示当前是第zc周，则第1周周一的日期 = 当前周周一的日期 - (zc-1) * 7天
    weeks_to_subtract = data['zc'] - 1  # 需要回推的周数
    days_to_subtract = weeks_to_subtract * 7  # 转换为天数

    first_week_monday = start_datetime - datetime.timedelta(days = days_to_subtract)  # 计算第1周周一的日期

    # 输出提示信息
    print("当前读取的学期是:", data['xnxqh'])
    print("当前学期开学时间(第1周周一)为:", first_week_monday)

    return first_week_monday

# 创建课程对应的日历事件，并加入到cal日历中
def addCourse(date, date_now, course, cal):
    # 新建课程并输入基本信息
    event = Event()
    event.add('summary', course['kcmc']) # 课程名称，作summary字段

    if course['jsmc'] != None and course['jsxm'] != None: # 教室名称字段及教师姓名字段均非空，正常处理
        event.add('LOCATION', course['jsmc'] + ' ' + course['jsxm']) # 教室名称 + 教师姓名，仿WakeUp课程表LOCATION字段
        event.add('DESCRIPTION', '第' + course['kkzc'] + '周 ' + course['jsmc'] + ' ' + course['jsxm']) # 开课周次 + 教室名称 + 教师姓名，仿WakeUp课程表DESCRIPTION字段
    elif course['jsmc'] == None and course['jsxm'] != None: # 教室名称字段为空，教师姓名字段非空，仅保留教师姓名
        event.add('LOCATION', course['jsxm']) # 教师姓名，仿WakeUp课程表LOCATION字段
        event.add('DESCRIPTION', '第' + course['kkzc'] + '周 ' + course['jsxm']) # 开课周次 + 教师姓名，仿WakeUp课程表DESCRIPTION字段
    elif course['jsmc'] != None and course['jsxm'] == None: # 教室名称字段非空，教师姓名字段为空，仅保留教室名称
        event.add('LOCATION', course['jsmc']) # 教室名称，仿WakeUp课程表LOCATION字段
        event.add('DESCRIPTION', '第' + course['kkzc'] + '周 ' + course['jsmc']) # 开课周次 + 教室名称，仿WakeUp课程表DESCRIPTION字段
    elif course['jsmc'] == None and course['jsxm'] == None: # 教室名称字段和教师姓名字段均空，垃圾课程没救了，希望它能1000年生きてる
        event.add('LOCATION', "生き汚く生きて 何かを創ったらあなたの気持ちが１０００年生きられるかも しれないから") # 摘自1000年生きてる / いよわ feat.初音ミク（living millennium / Iyowa feat.Hatsune Miku） 直达链接：https://www.youtube.com/watch?v=3em-J9yYPAo
        event.add('DESCRIPTION', "生き汚く生きて 何かを創ったらあなたの気持ちが１０００年生きられるかも しれないから") # 摘自1000年生きてる / いよわ feat.初音ミク（living millennium / Iyowa feat.Hatsune Miku） 直达链接：https://www.youtube.com/watch?v=3em-J9yYPAo

    # 切分课程开始时间、结束时间字符串为小时及分钟
    start_time_str = course['kssj']
    end_time_str = course['jssj']
    start_hours, start_minutes = [int(part) for part in start_time_str.split(":")]
    end_hours, end_minutes = [int(part) for part in end_time_str.split(":")]

    # 计算该课程所在的包含年月日信息的开始时间和结束时间
    course_weekday = int(course['kcsj'][0]) # 课程时间项第一个数字为课程所在星期
    course_start_date = date + datetime.timedelta(days = (course_weekday - 1), hours = start_hours, minutes = start_minutes)
    course_end_date = date + datetime.timedelta(days = (course_weekday - 1), hours = end_hours, minutes = end_minutes)

    # 添加该课程的时间信息
    event.add('dtstart', course_start_date)
    event.add('dtend', course_end_date)
    event.add('dtstamp', date_now)

    # 将课程添加到日历中，并给与提示
    cal.add_component(event)
    print("时间从{}到{}的{}课程已成功添加到日历中".format(course_start_date, course_end_date, course['kcmc']))

# 主函数捏
if __name__ == "__main__":
    # 输入账号密码信息
    account = input("请输入您的黄岛科技大学教务账号：")
    password = input("请输入您的黄岛科技大学教务账号密码：")

    Q = AcademicAffairs.SW(account, password, url)

    startDate = startTermDate(Q) # 推演第1周周一，即开学时间，作为后续操作的基准时间
    date_now = datetime.datetime.now(tz = pytz.timezone('Asia/Shanghai')) # 获取现在本机时间，用于创建DTSTAMP时间戳

    cal = Calendar() # 创建一个日历

    # 选择生成的课表周数
    start = int(input("请输入您想生成的课表起始周："))
    end = int(input("请输入您想生成的课表结束周："))

    # 遍历各周(start周-end周)，创建相应课程，并加入日历
    for i in range(start, end + 1):
        date = startDate + datetime.timedelta(days = (i - 1) * 7) # 当前遍历周次基准时间为当前周周一

        data_json = Q.get_class_info(i) # 读取第i周的课程JSON文本文件
        data = json.loads(data_json) # JSON字符串解释为List

        if data != [None]: # 若当前周存在课程才添加课程
            for course in data:
                addCourse(date, date_now, course, cal)
        else:
            print("该周课程为空")
        
        print("第{}周创建完毕".format(i))

    # 生成.ics文件
    with open('MyCourse.ics', 'wb') as f:
        f.write(cal.to_ical())
        print(".ics文件生成完成")