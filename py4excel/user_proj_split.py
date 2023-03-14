'''
Descripttion: 人员参与项目情况表拆分1，每人每天每一个项目产生一条记录
Author: RidgeJns, ridgejns@outlook.com
Date: 2023-03-07 14:07:47
LastEditTime: 2023-03-07 18:11:29
Copyright: (c) RidgeJns
'''
import pandas as pd

# 开始/结束时间
start_date = pd.Timestamp(2023, 3, 6)
end_date = pd.Timestamp(2023, 3, 10)
input_file_path = 'source/2023大数据应用交付部-人员变更情况跟踪表.xlsx'

df = pd.read_excel(input_file_path, sheet_name='参与项目情况')

# 时间轴
timeline = df.loc[2]

start_idx = 0
end_idx = 1
for timeline_idx, idx_date in enumerate(timeline):
    print(idx_date, start_date, start_date == idx_date, end_date)
    if idx_date == start_date:
        start_idx = start_idx + timeline_idx
    elif idx_date == end_date:
        end_idx = end_idx + timeline_idx
        break

# 数据
data = []
for idx, row in df.iterrows():
    if idx < 4:  # 跳过前四行数据
        continue
    for col in range(start_idx, end_idx):
        tasks = ""
        if len(row[col]) > 0:
            if '；' in row[col]:
                tasks = str(row[col]).split('；')
            elif ';' in row[col]:
                tasks = str(row[col]).split(';')
            elif '，' in row[col]:
                tasks = str(row[col]).split('，')
            elif ',' in row[col]:
                tasks = str(row[col]).split(',')
            else:
                tasks = str(row[col]).split(';')

        if row[col] == "-" or tasks == "":
            continue

        avg_work_time = 1 / len(tasks)
        for i, task in enumerate(tasks):
            remark = ""
            if len(task) == 0:
                continue
            if task[0] == "#":
                tag_task = task.split("#")
                tag = "#" + tag_task[1] + "#"
                if len(tag_task) > 1:
                    remark = tag_task[2].strip()
            else:
                tag = task
            if tag == "L":
                tag = "#请假#"
            elif tag == "T":
                tag = "#调休#"
            tag = tag.strip()
            print(idx, row[0], timeline[col], "tag:", tag, "remark:", remark, "avg_work_time:", avg_work_time)
            data.append([row[0], timeline[col], tag, remark, avg_work_time])
df2 = pd.DataFrame(columns=["姓名", "日期", "参与项目", "备注", "平均工时"])
for i, d in enumerate(data):
    df2.loc[i] = d

output_name = "source/2023ups_" + start_date.strftime("%Y%m%d") + "-" + end_date.strftime("%Y%m%d") + ".xlsx"
df2.to_excel(output_name, index=False)
