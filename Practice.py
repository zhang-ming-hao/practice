#!/usr/bin/env python
# coding:utf-8

"""
小学生每日一练，自动生成每日一练题目
v0.1: 支持口算题目生成
"""

import random

import docx_util as du

# ================================================================================
class Mathematics:
    """
    数据口算题目生成器类
    """

    OP_ADDITION             = 0x01      # 加法
    OP_SUBTRACTION          = 0x02      # 减法
    OP_MULTIPLICATION       = 0x04      # 乘法
    OP_DIVISION             = 0x08      # 除法

    # --------------------------------------------------------------------------------
    def __init__(self, op, max_value, in_table=True):
        """
        Mathematics的构造函数

        :param op:        操作符
        :param max_value: 积或和的最大值
        :param in_table:  是否为表内乘除法
        """

        self.op = op
        self.maxv = max_value
        self.in_table = in_table
        
    # --------------------------------------------------------------------------------
    def get_question(self):
        """
        取得题目

        :return: 题目
        """

        while True:
            # 生成操作符
            op = [self.OP_ADDITION, self.OP_SUBTRACTION, self.OP_MULTIPLICATION, self.OP_DIVISION][random.randint(0, 3)]
            if self.op & op:
                # 操作符可用
                if op == self.OP_ADDITION:
                    # 生成加法
                    a = random.randint(0, self.maxv)
                    b = self.maxv - a
                    return '%d + %d = ' % (a, b)
                elif op == self.OP_SUBTRACTION:
                    # 生成减法
                    a = random.randint(0, self.maxv)
                    b = random.randint(0, a)
                    return '%d - %d = ' % (a, b)
                elif op == self.OP_MULTIPLICATION:
                    # 表内乘法
                    a = random.randint(1, 10)
                    b = random.randint(1, 10)
                    return '%d × %d = ' % (a, b)
                elif op == self.OP_DIVISION:
                    # 表内除法
                    a = random.randint(1, 10)
                    b = random.randint(1, 10)
                    return '%d ÷ %d = ' % (a*b, b)


    # --------------------------------------------------------------------------------
    def __iter__(self):
        """
        迭代
        :return: 题目
        """

        return self.get_question()

    # --------------------------------------------------------------------------------
    def __next__(self):
        """
        下一个
        :return: 题目
        """

        return self.get_question()


# --------------------------------------------------------------------------------
def main():

    docx = du.Docx()
    title = docx.CreateStyle(fontSize=16, align=du.Style.STYLE_ALIGN_CENTER)
    question = docx.CreateStyle(fontSize=12)

    row = 15
    col = 4

    for day in range(20):
        docx.AddParagraph(title)
        docx.AddText('日期__________ 用时__________ 错题__________', title)
        docx.AddTable(row, col, '')

        # 口算
        m = Mathematics(Mathematics.OP_ADDITION | Mathematics.OP_SUBTRACTION |  Mathematics.OP_MULTIPLICATION, 100)
        for i in range(row):
            for j in range(col):
                docx.SetCell(i, j, next(m), question)
        docx.AddParagraph(title)
        docx.AddText('\n')

    docx.Save('每日一练.docx')

if __name__ == '__main__':
    main()
