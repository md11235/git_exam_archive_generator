# -*- coding: utf-8 -*-

import sys

import numpy as np
import pandas as pd

STU_ID_CN_STAR = "*学号"
STU_ID_CN = "学号"
STU_ID_EN = "id"

STU_NAME_STAR = "*姓名"
STU_ADMIN_CLASS="行政班"

STU_NAME= "姓名"
STU_ADMIN_CLASS2="行政班级"

LABEL_PRESENCE_SCORE = "考勤成绩"
LABEL_AVG_LAB_SCORE= "实验成绩"
LABEL_CLASS_PERF="课堂表现"

LABEL_SEM_SCORE = "*平时(50%)"
LABEL_EXAM_SCORE = "*期末(50%)"

LABEL_TOTAL_SCORE = "总成绩"

LINUX_CHAPTER_NAMES = [
    "第1章",
    "第2章",
    "第3章",
    "第4章",
    "第5章",
    "第6章",
    "第7章",
    "第8章",
    "第9章",
    "第10章",
    "第11章",
    "第12章",
    "第13章",
    "第15章",
    "第17章"]

class Submission:
    def __init__(self,
                 main_template_excel_file,
                 lab_score_csv_file,
                 # class_performance_score_file,
                 presence_score_file,
                 semester_exam_file,
                 chapter_names=None):
        if chapter_names is None:
            chapter_names = LINUX_CHAPTER_NAMES

        self.df_main = pd.read_excel(main_template_excel_file, skiprows=2)
        # print(self.df_main)
        self.df_class_performance = self.df_main[[STU_ID_CN_STAR]].copy(deep=True)

        self.df_main.set_index(STU_ID_CN_STAR, inplace=True)
        # generate class performance on the fly
        for chap_name in chapter_names:
            self.df_class_performance[chap_name] = 100
        self.df_class_performance[LABEL_CLASS_PERF] = self.df_class_performance.apply(lambda row: np.average(row[1:]), axis=1)
        self.df_class_performance.set_index(STU_ID_CN_STAR, inplace=True)


        self.df_raw_lab_score = pd.read_csv(lab_score_csv_file, delimiter='\s+', header=0)
        self.df_raw_lab_score.set_index(STU_ID_EN, inplace=True)
        self.df_lab_score = self.df_raw_lab_score.loc[:, self.df_raw_lab_score.columns.str.contains("lab")]
        self.df_lab_score[LABEL_AVG_LAB_SCORE] = self.df_lab_score.apply(lambda row: np.average(row), axis=1)

        self.df_raw_presence_score = pd.read_csv(presence_score_file, skiprows=5)
        self.df_raw_presence_score = self.df_raw_presence_score.loc[:,
                                     ~ self.df_raw_presence_score.columns.str.contains('^Unnamed')]
        self.df_raw_presence_score.set_index(STU_ID_CN, inplace=True)
        self.df_presence_score = self.df_raw_presence_score.loc[:,
                                 self.df_raw_presence_score.columns.str.contains('[0-9]')]
        # self.df_presence_score[STU_NAME_STAR] = self.df_raw_presence_score[STU_NAME].copy(deep=True)
        # self.df_presence_score[STU_ADMIN_CLASS] = self.df_raw_presence_score[STU_ADMIN_CLASS2].copy(deep=True)
        self.df_presence_score[LABEL_PRESENCE_SCORE] = self.df_presence_score.apply(
            lambda row: (10-row.str.fullmatch(r'x|X').sum())*10,
            #lambda row: (10-row.notnull().sum())*10,
            axis=1)

        print(self.df_presence_score)

        self.df_raw_semester_exam = pd.read_excel(semester_exam_file)
        self.df_raw_semester_exam.set_index(STU_ID_CN, inplace=True)

        # # self.df_calc_lab_score =
        # print(self.df_main)
        # print(self.df_raw_lab_score)
        #
        # print(self.df_raw_lab_score.loc[self.df_main.index])

        print(self.df_presence_score.loc[self.df_main.index])

        self.df_daily_perf_score = self.df_presence_score[[LABEL_PRESENCE_SCORE]].copy(deep=True).loc[self.df_main.index]
        self.df_daily_perf_score[LABEL_AVG_LAB_SCORE] = self.df_lab_score[LABEL_AVG_LAB_SCORE].loc[self.df_main.index]
        self.df_daily_perf_score[LABEL_CLASS_PERF] = self.df_class_performance[LABEL_CLASS_PERF].loc[self.df_main.index]

        # print(self.df_class_performance)

        self.df_daily_perf_score[LABEL_SEM_SCORE] = self.df_daily_perf_score[LABEL_PRESENCE_SCORE]*0.10 + \
        self.df_daily_perf_score[LABEL_AVG_LAB_SCORE]*0.80 + self.df_daily_perf_score[LABEL_CLASS_PERF]*0.10

        self.df_main[LABEL_SEM_SCORE] = self.df_daily_perf_score[LABEL_SEM_SCORE]
        self.df_main[LABEL_EXAM_SCORE] = self.df_raw_semester_exam[LABEL_TOTAL_SCORE]
        self.df_main["备注"] = self.df_main[LABEL_SEM_SCORE]*0.5 + self.df_main[LABEL_EXAM_SCORE]*0.5

        self.df_raw_semester_exam.loc[self.df_raw_semester_exam["考试状态"] == "完成", "考试状态"] = None
        self.df_main["特殊成绩标识"] = self.df_raw_semester_exam["考试状态"]

        print(self.df_main.loc[self.df_main["特殊成绩标识"].isnull()])
        print(self.df_main.loc[self.df_main["特殊成绩标识"].notnull()])

        print(self.df_main.loc[ (55.0 < self.df_main["备注"]) & (self.df_main["备注"]<60.0) ])


if __name__ == '__main__':
    submission = Submission(main_template_excel_file=sys.argv[1],
                            lab_score_csv_file=sys.argv[2],
                            presence_score_file=sys.argv[3],
                            semester_exam_file=sys.argv[4])




