import numpy as np
import pandas as pd
import os
import msoffcrypto
 
class merge:

    # 类的参数为当前文件夹的目录，因为我们要合并当前文件夹下面的所有excel文件，
    # 建议写绝对路径
    def __init__(self,current_directory):
        
        self.instructions()
        self.current_directory = current_directory
        self.content_info = []
        
        self.result_name = self.current_directory + '\\result\\' +'info.xls'

        # self.need_decrypt = True  # 是否需要解密
        self.is_need_decrypt()  # 让用户判断是否需要解密

        self.nan_alert = False  # nan的警告标志
        self.scan(self.current_directory) # 找出当前目录下，需要合并的excel表格
        self.create_result_catalogue(self.current_directory + '\\result')

        
        self.my_Columns = ['    指标','一、基本情况', '行政区域土地面积', '  乡(镇)个数', '  村民委员会个数', '  年末总户数', 
    '    其中：乡村户数', '  年末总人口', '    其中：乡村人口', '  年末单位从业人员数', '  乡村从业人员数', 
    '    其中：农林牧渔业', '  农业机械总动力', '  农村用电量', '  本地电话用户', '二、综合经济', '  第一产业增加值',
    '  第二产业增加值', '  地方财政预算内收入', '  财政支出', '  城乡居民储蓄存款余额', 
    '  年末金融机构各项贷款余额', '三、农业、工业及基本建设', '  粮食产量', '  棉花产量', '  油料产量', '  肉类总产量', 
    '  规模以上工业企业数', '  规模以上工业总产值(现价)', '  基本建设投资完成额', '四、教育、卫生和社会保障', 
    '  普通中学在校学生数', '  小学在校学生数', '  医院、卫生院床位数','   社会福利院数', '  社会福利院床位数']

        self.Unit = ['  单位', np.nan, '  平方公里', '    个', '    个', '    户', '    户', '    万人', '    万人',
 '    人', '    人', '    人', '  万千瓦特', '  万千瓦时', '    户', np.nan, '    万元', '    万元',
 '    万元', '    万元', '    万元', '    万元' ,np.nan, '    吨', '    吨', '    吨', '    吨',
 '    个', '    万元', '    万元', np.nan, '    人', '    人', '    床', '    个' ,'    床']

        # 必须放在后面
        self.Dict = {}
        self.init_Dict()
        
        


    def is_need_decrypt(self):
        a = input('如果您的excel表是加密文件的话，需要执行解密操作。\n输入[y/Y]，表示需要解密。\n输入其他任意字符表示不需要解密操作。\n:')
        if a in 'yY':
            self.need_decrypt = True
        else:
            self.need_decrypt = False

    def init_Dict(self):
        count = 0
        for p in self.my_Columns:
            # if p not in self.Dict.keys():
            #     self.Dict[p] = []
            self.Dict[p] = [self.Unit[count]]
            count += 1        

    def create_result_catalogue(self,path):

        # mkdir 
        # 去除首位空格
        path=path.strip()
        # 去除尾部 \ 符号
        path=path.rstrip("\\")
    
        # 判断路径是否存在
        # 存在     True
        # 不存在   False
        isExists=os.path.exists(path)
    
        # 判断结果
        if not isExists:
            # 如果不存在则创建目录
            # 创建目录操作函数
            os.makedirs(path) 
    
            print(path+' 创建成功')
            return True
        else:
            # 如果目录存在则不创建，并提示目录已存在
            print( path+' 目录已存在')
            return False

    # def create_result_catalogue():

    def instructions(self):
        # print('请确保您输入的文件夹里面全部是加密的excel表格，\n如果存在未加密的excel表格(比如其他无关的表格)，程序将会在中途退出！')
        pass

    def scan(self,file_path):
        # 扫描当前文件夹下面的所有的excel文件,为后面excel的合并做准备
        name = os.listdir(file_path)

        # print('name',name)
        self.xlsx_postfix = []
        for i in name:
            if i[-5::] == '.xlsx' or i[-4::] == '.xls':
                self.xlsx_postfix.append(i)

        
        print(self.xlsx_postfix)
        print('将会对上述excel表格进行合并！\n')

    def decrypt(self,rawname):

        file = msoffcrypto.OfficeFile(open(rawname, 'rb'))  # 读取原文件
        file.load_key(password='VelvetSweatshop')  # 填入密码, 若能够直接打开, 则为默认密码'VelvetSweatshop'
        
        # self.unlock_name为存放解密文件的地方
        file.decrypt(open(self.unlock_name, 'wb'))  # self.unlock_name 为文件解密后的文件名字

    def get_pandas_series(self,pd_series):
        # need = False  # 这个属性列是否是我们需要的，就是判断这个属性列是否是 县级的信息 
        str_flag = False  # 如果在这一列没有遇到字符串就算了，若有字符串那么就要准备判断后面是否是数字
        offset = 0
        later_check = -1
        lock = False   # 变量上锁
        for ps in pd_series:
            
            # 在没有遇到字符串时，记录此时的偏移量，方便后续根据偏移量来进行截断
            if str_flag is False:
                offset += 1
                # print('offset:',offset)
            else:
                # 此时字符串已找到，对后续的数据进行检测，从而判断出这一列是否是县级信息
                later_check += 1
            
            if str_flag is False and type(ps) is not str:
                # print('continue:',type(ps),ps)
                continue
            else:
                str_flag = True
                # 字符串已找到，此时激活后续判断的变量
                if lock is False:
                    # later_check 在赋值了一次之后，就上锁，以后不允许再被赋值
                    later_check = offset
                lock = True

            # 如果later_check一直都是 -1 说明在这一列中一直都没有找到字符串
            if later_check != -1:
                # 如果后续检测的变量已激活
                if later_check - offset == 1:

                    # if type(ps) != np.nan:  # 这个判别不行，卡了我好久
                    # 此时我们判断nan的方法是用float来检验
                    if type(ps) != float:
                        # print('type(ps)',ps)
                        # print('judge:',ps == np.nan)
                        # print('出口1')
                        return (False, None)
                if later_check - offset == 2:
                    if type(ps) == int:
                        # [1,36],[2,38] 
                        temp = pd_series[offset-1::].values  
                        # new_series = pd.Series(pd_series[offset-1::].values,index=range(2,38))
                        new_series = pd.Series(temp[:36:],index=range(2,38))  # temp长了只能截断了
                        # return (True,pd_series[offset-1::])
                        # print('出口2')
                        # print(new_series)
                        return (True,new_series)

                # 此时相差的第2格不是nan
                # 后续再出现int，int已经错过了机会，一旦int再次出现，退出循环，触发nan异常的警告
                if type(ps) == int:
                    # nan的警告只会在这个地方被触发。
                    # 既然到了这里，肯定不是正常规范的数据，但是又出现了int，触发警告通知用户
                    self.nan_alert = True
                    # print('出口3')  
                    return (False, None)  # 只有这个地方第三个返回参数是True，这说明出现了重复的nan

                if type(ps) == str and ps.isdigit() == True:
                    self.str_int_alert = True

                    
        # 当for循环结束，如果在前面没有跳出循环那么我们return
        # print('出口3')
        # print('出口4') 
        return (False,None)

    def get_per_excel_series(self,filename):
        df = pd.read_excel(filename)
        Columns = df.columns.values
        L = []
        for c in Columns:
            v = self.get_pandas_series(df[c])
            if v[0] == False:
                continue
            elif v[0] == True:
                L.append(v[1])

            # 进行nan多行的检测


        if len(L) != 0:
            return (True,L)
        else:
            return (False,None)


    def generate_Dict(self,excel_name):
        
    
        # print(len(my_Columns))

        df = pd.read_excel(excel_name)

        Columns = df.columns.values
        # L = [] 此时已经无用
        for c in Columns:
            v = self.get_pandas_series(df[c])
            if v[0] == False:
                continue
            elif v[0] == True:

                # 把这个v[1]散开进列表里面
                # L.append(v[1]) 以前的方法

                count = 0
                for p in self.my_Columns:
                    if p not in self.Dict.keys():
                        self.Dict[p] = []
                    l = v[1].values
                    self.Dict[p].append(l[count])
                    count += 1



    def Dict_excel(self,Dict,path):
        a = pd.DataFrame.from_dict(Dict)
        a.to_excel(path)



    def Main(self):
        '''
        依次对当前文件夹内的excel文件进行解密(unlock)，加密密码是默认密码。 self.decrypt()
        可以试试如果不进行解密，编译器会报当前文件进行了加密。
        解密的文件命名为unlock.xls,下一次解密对前一次的unlock.xls进行覆盖。

        然后对每一个解密后的excel文件挑选出指定的列，把这些pandas文件添加进类的列表(self.context)  self.fun()
        '''
        # print(self.xlsx_postfix)
        for i in self.xlsx_postfix:

            # 每个excel表开始时，初始化时，关掉nan的警告
            self.nan_alert = False

            self.str_int_alert = False

            # 依次解密每个excel文件(注意:我们没有判断每个文件是否全部加密过，默认对所有的excel文件进行解密)
            
            # 为原始原件的存放路径
            self.file_path = self.current_directory + '\\' + i
            
            if self.need_decrypt == True:
                self.unlock_name = self.current_directory + '\\result\\' + 'unlock.xls'

                # 传原始的excel表的路径进去
                # self.decrypt(self.unlock_name)

                self.decrypt(self.file_path)
                self.generate_Dict(self.unlock_name)

            elif self.need_decrypt == False:
                # 不需要解密的文件
                
                self.generate_Dict(self.file_path)
            
            # 判断nan多行重复的报警是否触发
            if self.nan_alert is True:
                print('警告:在',i,'中存在多列的空格，此处绝对有错误，请人工修正！程序并未添加的数据信息。')
            
            if self.str_int_alert == True:
                print('警告:在',i,'中的整数为字符串型，此处绝对有错误，请人工修正！程序并未添加这些数据信息。')


        self.Dict_excel(self.Dict,self.result_name)  # 把字典文件导出成excel表

        print('查看您输入路径里面的result内的info.xls，即为最后的输出文件。\n(您可以删除没有用的中途文件unlock.xls)')


if __name__ == '__main__':

    go = True

    while(go):
        # e = merge(r'C:\Users\tiffa\Desktop\work_10_13')
        folder_path = input('请输入您存放excel的文件夹的绝对路径，我们将会把此文件夹内的所有excel进行合并。\n:')
        e = merge(folder_path)
        e.Main()
        print('程序运行结束！\n如果您希望继续运行，请输入[y/Y],\n希望退出程序输入其他任何字符即可！')
        go = input(':')
        if go not in 'yY':
            go = False

        

