import openpyxl
from tkinter import *

team_dic = {"D": "多特蒙德CFD 13华人足球队", "BO": "波鸿原点Ppagei华人足球队",
            "N": "KSC弗兰肯足球联队", "DU": "打酱油杜伊斯堡队",
            "B": "柏林华人足球队", "S": "斯图加特华人足球队",
            "C": "开姆尼茨华人足球队", "L": "卢森堡华人足球协会",
            "M": "慕尼黑华人联合足球俱乐部", "F": "法兰克福坚强足球队",
            "DD": "德累斯顿CFC华人足球队", "SCH": "Schöneberg华人足球队"}

image_dic = {"D": "多特蒙德.png", "BO": "波鸿.png",
            "N": "纽伦堡.png", "DU": "杜伊斯堡.png",
            "B": "柏林.png", "S": "斯图加特.png",
            "C": "开姆尼茨.png", "L": "卢森堡.png",
            "M": "慕尼黑.png", "F": "法兰克福.png",
            "DD": "德累斯顿.png", "SCH": "Schoeneberg.png"}

wb = openpyxl.load_workbook("小组积分榜.xlsx")
ws_a = wb["A小组"]
ws_b = wb["B小组"]
ws_c = wb["C小组"]

# 2-image, 3-name, 4-win, 5-draw, 6-lose, 7-goal, 8-against, 9-difference, 10-score


class Team(object):
    def __init__(self, image, name, win, draw, lose, goal, against, difference, score):
        self.image = image
        self.name = name
        self.win = win
        self.draw = draw
        self.lose = lose
        self.goal = goal
        self.against = against
        self.difference = difference
        self.score = score

    def print_all_info(self):
        print("image: " + self.image)
        print("name: " + self.name)
        print("win: " + str(self.win))
        print("draw: " + str(self.draw))
        print("lose: " + str(self.lose))
        print("goal: " + str(self.goal))
        print("against: " + str(self.against))
        print("difference: " + str(self.difference))
        print("score: " + str(self.score))
'''
    @property
    def image(self):
        return self.image

    @image.setter
    def image(self, file):
        self.image = file

    @property
    def name(self):
        return self.name

    @name.setter
    def name(self, value):
        self.name = value

    @property
    def win(self):
        return self.win

    @win.setter
    def win(self, value):
        self.win = value

    @property
    def draw(self):
        return self.draw

    @draw.setter
    def draw(self, value):
        self.draw = value

    @property
    def lose(self):
        return self.lose

    @lose.setter
    def lose(self, value):
        self.lose = value

    @property
    def goal(self):
        return self.goal

    @goal.setter
    def goal(self, value):
        self.goal = value

    @property
    def against(self):
        return self.against

    @against.setter
    def against(self, value):
        self.against = value

    @property
    def difference(self):
        return self.difference

    @difference.setter
    def difference(self, value):
        self.difference = value

    @property
    def score(self):
        return self.score

    @score.setter
    def score(self, value):
        self.score = value'''


def evaluate(a, score_a,  b, score_b):
    a.goal = a.goal + score_a
    a.against = a.against + score_b
    b.goal = b.goal + score_b
    b.against = b.against + score_a
    a.difference = a.goal - a.against
    b.difference = b.goal - b.against

    if score_a > score_b:
        a.win = a.win + 1
        b.lose = b.lose + 1
        a.score = a.score + 3

    elif score_a < score_b:
        b.win = b.win + 1
        a.lose = a.lose + 1
        b.score = b.score + 3

    else:
        a.draw = a.draw + 1
        b.draw = b.draw + 1
        a.score = a.score + 1
        b.score = b.score + 1


team_a = Team(image_dic["D"], team_dic["D"], 0, 0, 0, 0, 0, 0, 0)
team_b = Team(image_dic["N"], team_dic["N"], 0, 0, 0, 0, 0, 0, 0)

evaluate(team_a, 2, team_b, 0)

team_a.print_all_info()
print()
team_b.print_all_info()


