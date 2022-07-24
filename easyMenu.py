# 基于链表的菜单管理系统
"""   C++的实现不了，限制太多   """
import collections


class EasyMenu:
    def __init__(self):
        self.__level = 0  # 顶级目录
        self.__funs = {}  # 当前各级菜单要调用的函数
        self.__nums = {}
        self.__menus = collections.OrderedDict()  # 实际上自Python3.6起，字典的keys就是”有序“的了

    def __display(self):
        for idx, title in enumerate(self.__menus):
            print('{}.{}'.format(idx + 1, title))
        # 支持返回功能
        print('{}.'.format(len(self.__menus) + 1), end='')
        if self.__level == 0:
            print('退出程序')
        else:
            print('返回上一层')

    def add(self, menus):
        for idx, (k, v) in enumerate(menus.items()):
            self.__funs[k] = v
            self.__nums[idx + 1] = k
            self.__menus[k] = EasyMenu()  # 生成一个新对象
            self.__menus[k].__level = self.__level + 1  # 层级加1
        return self  # 返回自身

    def run(self):
        while True:
            self.__display()
            flag = True
            while flag:
                line = input()
                try:
                    opt = int(line)
                except Exception:
                    print('请输入一个数字！')
                    continue
                if 1 <= opt <= len(self.__menus) + 1:
                    flag = False
                    if opt == len(self.__menus) + 1:
                        return
                    elif self.__menus[self.__nums[opt]].__menus:  # 优先调用子菜单
                        self.__menus[self.__nums[opt]].run()
                    elif self.__funs[self.__nums[opt]]:
                        self.__funs[opt - 1]()
                else:
                    print('请输入1~{}范围内的数字！'.format(len(self.__menus) + 1))

    def __getitem__(self, item):
        return self.__menus[item]


if __name__ == '__main__':
    menu = EasyMenu().add({
        '登录系统': None,
        '查询': None,
        '修改': None,
        '新增': None,
    })
    menu['登录系统'].add({
        '查看信息': None,
        '修改密码': None,
    })
    menu.run()
