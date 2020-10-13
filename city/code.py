import numpy as np
import pandas as pd
import os
import msoffcrypto
 
class merge:

    # 类的参数为当前文件夹的目录，因为我们要合并当前文件夹下面的所有excel文件，
    # 建议写绝对路径
    def __init__(self,current_directory):

        self.current_directory = current_directory
        self.content_info = []
        self.unlock_name = 'unlock.xls'
        self.result_name = 'info.xls'

        self.scan(self.current_directory) # 找出当前目录下，需要合并的excel表格

    def scan(self,file_path):
        # 扫描当前文件夹下面的所有的excel文件,为后面excel的合并做准备
        name = os.listdir(file_path)

        print('name',name)
        self.xlsx_postfix = []
        for i in name:
            if i[-5::] == '.xlsx' or i[-4::] == '.xls':
                self.xlsx_postfix.append(i)

        print(self.xlsx_postfix)
        print('将会对上述excel表格进行合并！')

    def decrypt(self,rawname):

        file = msoffcrypto.OfficeFile(open(rawname, 'rb'))  # 读取原文件
        file.load_key(password='VelvetSweatshop')  # 填入密码, 若能够直接打开, 则为默认密码'VelvetSweatshop'
        file.decrypt(open(self.unlock_name, 'wb'))  # self.unlock_name 为文件解密后的文件名字

    def get_columns(self,df):
        # 函数的参数为pandas文件(df)
        # 得到当前文件夹的  XX市、XX县 的信息，方便后续我们根据列的属性名，挑出这一列的数据
        Columns = df.columns.values

        res_columns = [] # 存放当前excel文件有 XX县、xx市

        for i in Columns:
            # 值得注意的是，我只选出了 XX县 和 XX市。（注意：在未来可能存在丢失数据情况）
            # 其他的都没有进行检测，比如出现一个XX村，这个XX村的数据是不会被选出来的。
            if i[-1] == '县' or i[-1] == '市': 
                res_columns.append(i)
        return res_columns



    def fun(self,filename):

        df = pd.read_excel(filename,header = 2) # 记住header从第2行开始，日后excel数据改变了，需要改header的值
        
        # 具体 xx县、xx市 的名字，从而挑选出这些列的数据
        city_name = self.get_columns(df)

        # 最左边的一列指标属性,单独拿出来，在最后只加一次，避免重复
        self.zhibiao = df['    指标']

        # 挑出具体的几列数据，这个就是我们需要的数据
        # 直接返回即可，后续会把它添加进类的信息列表(self.content_info)，最后再进行所有excel信息的合并
        return df.loc[:,city_name] 
         

    def Main(self):
        '''
        依次对当前文件夹内的excel文件进行解密(unlock)，加密密码是默认密码。 self.decrypt()
        可以试试如果不进行解密，编译器会报当前文件进行了加密。
        解密的文件命名为unlock.xls,下一次解密对前一次的unlock.xls进行覆盖。

        然后对每一个解密后的excel文件挑选出指定的列，把这些pandas文件添加进类的列表(self.context)  self.fun()
        '''
        print(self.xlsx_postfix)
        for i in self.xlsx_postfix:
            
            # 依次解密每个excel文件(注意:我们没有判断每个文件是否全部加密过，默认对所有的excel文件进行解密)
            self.decrypt(i)

            # self.fun(i)得到需要的数据, 把数据添加进内容列表
            self.content_info.append(self.fun(self.unlock_name)) 
            # 我们只有一个解密后的文件self.unlock_name，这个解密后的文件在新的解密文件出来后，就覆盖为新的解密excel文件了.

        print(self.content_info)
        # 把所有的列拼起来,直接拼在对应列的后面
        # new_excel = pd.concat([self.zhibiao,self.content_info],axis=1)  # self.zhibiao 是最左边的指标的一列信息
        new_excel = pd.concat(self.content_info,axis=1)  # self.zhibiao 是最左边的指标的一列信息
        
        # 输出新的excel文件
        new_excel.to_excel(self.result_name,columns = None,index = False)


if __name__ == '__main__':
    # e1 = merge("new安徽省.xls")
    # e2 = merge('new沈阳省.xlsx')
    # new = pd.concat([e1.zhibiao,e1.fun(),e2.fun()],axis=1)
    # new.to_excel('new_result.xlsx',columns = None,index = False)
    # res =  + e2.fun
    # print(res)
    e = merge(r'C:\Users\tiffa\Desktop\work_10_13')
    e.Main()
