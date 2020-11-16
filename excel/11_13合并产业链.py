import numpy as np
import pandas as pd

class excel_form:
    def __init__(self):
        self.filename = self.get_filename()
        self.columns_list =  self.get_column()

    def get_filename(self):

        # name = r'c:\Users\tiffa\Desktop\work11_13\样例数据.xlsx'
        name = input('输入你要操作的文件的文件名:')

        return name
    
    def get_column(self):
        self.df = pd.read_excel(self.filename)
        Columns = self.df.columns.values
        return Columns
        # print(Columns)

        
    def show(self):
        for i in self.columns_list:
            print(self.df[i])

    def engine(self):
        
        # 通过excel生成对应存放数值的列表 
        self.company_name = self.df[self.columns_list[0]].tolist()
        Len = len(self.company_name)
        self.GS1 = self.df[self.columns_list[1]].tolist()
        self.GS2 = self.df[self.columns_list[2]].tolist()
        self.GS3 = self.df[self.columns_list[3]].tolist()

        # 对最后存放 结果 的列表进行初始化
        self.com_res = []
        self.GS1_res = []
        self.GS2_res = []
        self.GS3_res = []

        # 寻找不同的公司名字，通过探路,发现不同,停止,进入内循环

        front = 0
        rear  = -100   #  故意设置
        # 开始点下标(同一公司名称的开始行数)，方便通过它作为下标进入内循环

        current_com = self.company_name[0]
        for i in range(0,Len):

            if len(self.com_res) == 0:
                self.com_res.append(current_com)
            elif current_com not in self.com_res:
                self.com_res.append(current_com)

            # 思考对nan的处理?不用思考了，程序已经走过了空白行

            # if front == rear + 2:
            #     # 这个状态是，在寻找另一个公司时，front已经跟上了rear，
            #     # 此时front与rear间没有任何数据
            #     # 若此时遇到nan,front与rear同时加一，
            #     # 继续维持这个状态，并且成功的跳过了nan
      
            #     if type(self.company_name[i]) == float: # 避免有多个空行，也就是出现连续的nan
            #         # 维持这种可供识别的状态
            #         front += 1
            #         rear += 1
            #         continue

            #     # 当前已经进入到有数据的行了，应该要开始记录了
            #     current_com = self.company_name[i]  # 记录开始
            #     front = i

            # 进入这个条件时，当前已经位于空白行了(nan),或者到达最后一行
            if current_com != self.company_name[i] or i == Len - 1:

                # print(current_com,' != ',self.company_name[i])
                # 探针发现不一样的公司
                # 确定前面一行的公司名称的 开始行数 和 结束的行数
                
                if i < Len - 1:
                    rear = i - 1  # 结束的行数
                else:
                    rear = Len - 1
                
                # print('front:',front)
                # print('rear:',rear)

                # 添加进归属1的结果列表
                self.GS1_res.append(self.deal_GS1(front,rear))
                
                # 添加进归属2的结果列表
                self.GS2_res.append(self.deal_GS2(front,rear))

                # 添加进归属3的结果列表
                self.GS3_res.append(self.deal_GS3(front,rear))

                # 已经传参完毕，更新下一轮的front(就是让front跟上rear)
                front = i + 1
                if i < Len-1:
                    current_com = self.company_name[front]  

                # current_com = self.company_name[i] # 这个赋值语句应该在下循环的开头 

    def deal_GS1(self,front,rear):
        # rear += 1
        # 处理归属1，直接遍历即可
        res = []
        for i in range(front,rear+1):
            # ,rear + 1 因为rear这一行就是结束行

            Str = self.GS1[i]
            if len(res) == 0:
                res.append(Str)
            elif Str not in res:
                res.append(Str)
        
        # print('1:',res)

        return '/n'.join(res)
    

    def deal_GS2(self,front,rear):
        # rear += 1
        # 处理归属1，直接遍历,把归属1与归属2进行拼接
        res = []
        for i in range(front,rear+1):
            # rear + 1 因为rear这一行就是结束行
            Str = self.GS1[i] +  '-'  + self.GS2[i]
            if len(res) == 0:
                res.append(Str)
            elif Str not in res:
                res.append(Str)
        
        # print('2:',res)
        return '/n'.join(res)

    
    def deal_GS3(self,front,rear):
        # rear += 1
        # 处理归属1，直接遍历,把归属1与归属2进行拼接
        res = []
        for i in range(front,rear+1):
            if self.GS3[i] == '-':
                continue
            # rear + 1 因为rear这一行就是结束行
            Str = self.GS1[i] +  '-'  + self.GS2[i] + '-' + self.GS3[i]
            if len(res) == 0:
                res.append(Str)
            elif Str not in res:
                res.append(Str)
        # print('3',res)
        return '/n'.join(res)


    def write_excel(self):
        wf = pd.DataFrame({'公司名称':self.com_res,'归属1':self.GS1_res,'归属2':self.GS2_res,'归属3':self.GS3_res})
        wf.to_excel('输出结果.xlsx',encoding='utf-8',index=False)

    def see_Len(self):
        print(self.com_res)
        print('len:',len(self.com_res))

        print(self.GS3_res)
        print('len:',len(self.GS3_res))
    
    def __del__(self):
        print('查看当前文件夹下面的 输出结果.xlsx')
        print('bye')




if __name__ == "__main__":
    Excel = excel_form()
    # Excel.value()
    # Excel.show()
    Excel.engine()
    # Excel.see_Len()
    Excel.write_excel()
    # print(Excel.GS3_res)

