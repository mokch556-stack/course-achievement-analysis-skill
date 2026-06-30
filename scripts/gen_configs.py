# -*- coding: utf-8 -*-
"""Skip aggregate, go straight to building configs (data already aggregated)"""
import sys, os, json

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)

SOURCE_DIR = r"C:\Users\Administrator\Desktop\大学计算机基础\大学计算机基础-教学文件-2025-2026(二）\课程档案\互金251+工管251成绩 (1)\互金251+工管251成绩"
OUTPUT_DIR = r"E:\qwenpaw生成\达成度\大学计算机基础"

MAJOR_MAP = {"互金251": "互联网金融", "工管251": "工程管理"}
INDICATOR_MAP = {
    "互联网金融": ["5-1", "5-2", "5-3", "10-1"],
    "工程管理":   ["5-3", "10-1", "5-3", "10-1"],
}

TARGET_ITEMS = [
    {"items":[{"name":"线上学习","raw_full":20,"zhe_full":5},{"name":"课堂表现","raw_full":20,"zhe_full":5},{"name":"阶段测试1","raw_full":100,"zhe_full":10},{"name":"学习手册","raw_full":20,"zhe_full":5},{"name":"期末考试","raw_full":100,"zhe_full":25}]},
    {"items":[{"name":"线上学习","raw_full":20,"zhe_full":5},{"name":"课堂表现","raw_full":20,"zhe_full":5},{"name":"阶段测试2","raw_full":100,"zhe_full":10},{"name":"学习手册","raw_full":20,"zhe_full":5},{"name":"期末考试","raw_full":100,"zhe_full":25}]},
    {"items":[{"name":"线上学习","raw_full":20,"zhe_full":5},{"name":"课堂表现","raw_full":20,"zhe_full":5},{"name":"阶段测试3","raw_full":100,"zhe_full":10},{"name":"学习手册","raw_full":20,"zhe_full":5},{"name":"期末考试","raw_full":100,"zhe_full":25}]},
    {"items":[{"name":"线上学习","raw_full":20,"zhe_full":5},{"name":"课堂表现","raw_full":20,"zhe_full":5},{"name":"阶段测试4","raw_full":100,"zhe_full":10},{"name":"学习手册","raw_full":20,"zhe_full":5},{"name":"期末考试","raw_full":100,"zhe_full":25}]},
]

from aggregate_scores import aggregate
raw_students = aggregate(SOURCE_DIR)

majors_data = {}
for s in raw_students:
    cls = s.get('class', '')
    major = MAJOR_MAP.get(cls.strip(), cls.strip())
    s['major'] = major
    majors_data.setdefault(major, []).append(s)

for major, students in majors_data.items():
    indicators = INDICATOR_MAP.get(major, ["5-1","5-2","5-3","10-1"])
    config = {
        "course_name":"大学计算机基础","course_code":"3300000005",
        "term":"2025-2026学年第2学期","assessment":"考查",
        "major":major,"classes":list({s['class'].strip() for s in students}),
        "teacher":"邹芳","hours":32,"credits":2,
        "kaohe_desc":"总评成绩=平时成绩×50%+期末考试×50%；考核方式：考查",
        "targets":[],"students":[],
    }
    for i in range(4):
        config["targets"].append({
            "id":i+1,"name":f"教学目标{i+1}","indicator":indicators[i],
            "items":TARGET_ITEMS[i]["items"],
        })
    for s in students:
        config["students"].append({
            "xh":s['xh'],"name":s['name'],"class":s['class'],
            "raw_scores":s['raw_scores'],
            "luru_ps":s.get('luru_ps',0),"luru_qm":s.get('luru_qm',0),"luru_zp":s.get('luru_zp',0),
        })
    config_path = os.path.join(OUTPUT_DIR, f"config_{major}.json")
    with open(config_path,'w',encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"✅ config_{major}.json ({len(students)}人)")
