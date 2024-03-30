# 青岛市西海岸新区黄岛区经济技术开发区辛安街道下庄社区东北附属学校 课程表转换.ics文件
# By 冷月おとわ
# 主体选单界面

import CoursesAPI2ICS
import CoursesExcel2ICS

# 主函数捏
if __name__ == "__main__":
    print("青岛市西海岸新区黄岛区经济技术开发区辛安街道下庄社区东北附属学校 课程表转换.ics文件小工具")
    print("请选择您生成课表的操作接口：")
    print("1: 使用强智教务系统API")
    print("2: 使用强智教务课表Excel文件")
    print("3: 使用强智教务系统网页登录")
    flag = int(input("请输入您想使用的操作接口的数字代号,0表示退出: "))

    if flag == 1:
        print("\n即将使用强智教务系统API生成课表.ics文件")
        CoursesAPI2ICS.CoursesAPI()
    elif flag == 2:
        print("\n即将使用强智教务课表Excel文件生成课表.ics文件")
        print("警告：含有单双周课程的课表暂时无法使用此方式")
        CoursesExcel2ICS.CoursesExcel()
    elif flag == 3:
        print("\n即将使用使用使用强智教务系统网页登录生成课表.ics文件")
        print("工事中,预计将和中央磁浮新干线打复活赛")
    else:
        exit(0)